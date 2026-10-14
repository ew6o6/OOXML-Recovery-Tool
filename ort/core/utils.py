"""Output, image-validation, filename, and metadata helpers."""

from __future__ import annotations

import binascii
import hashlib
import io
import os
import re
import struct
import xml.etree.ElementTree as ET
import zlib
from pathlib import Path

from PIL import Image, ImageFile


def safe_print(message):
    try:
        print(message)
    except UnicodeEncodeError:
        print(str(message).encode("ascii", errors="replace").decode("ascii"))


def safe_output_stem(file_path):
    stem = os.path.splitext(os.path.basename(file_path))[0].rstrip(" .")
    return stem or "recovered"


def _coerce_xml_text(value):
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="ignore")
    if not isinstance(value, str):
        return ""
    if value.lstrip().startswith("<"):
        return value
    try:
        return binascii.unhexlify(value).decode("utf-8", errors="ignore")
    except (binascii.Error, ValueError):
        return value


def extract_file_name(decoded_name):
    pattern = re.search(r"[\x00-\x1F\x7F]", decoded_name[::-1])
    if pattern:
        decoded_name = decoded_name[-pattern.end() + 1 :]
    return decoded_name


def sanitize_part_basename(name):
    basename = os.path.basename(name.replace("\\", "/")) or "unverified.bin"
    basename = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", basename).rstrip(" .")
    return basename[:180] or "unverified.bin"


def detect_image_format(raw):
    if raw.startswith(b"\xff\xd8\xff"):
        return "jpeg"
    if raw.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if raw.startswith((b"GIF87a", b"GIF89a")):
        return "gif"
    if raw.startswith(b"BM"):
        return "bmp"
    if raw.startswith((b"II*\x00", b"MM\x00*")):
        return "tiff"
    if raw.startswith(b"RIFF") and raw[8:12] == b"WEBP":
        return "webp"
    stripped = raw.lstrip(b"\xef\xbb\xbf \t\r\n")
    if stripped.startswith(b"<svg") or (stripped.startswith(b"<?xml") and b"<svg" in stripped[:4096]):
        return "svg"
    if len(raw) >= 44 and raw[:4] == b"\x01\x00\x00\x00" and raw[40:44] == b" EMF":
        return "emf"
    if raw.startswith((b"\xd7\xcd\xc6\x9a", b"\x01\x00\x09\x00")):
        return "wmf"
    if raw.startswith((b"II\xbc\x01", b"MM\x01\xbc")):
        return "wdp"
    if raw.startswith(b"RIFF") and raw[8:12] == b"WAVE":
        return "wav"
    return None


def _png_complete(raw):
    if not raw.startswith(b"\x89PNG\r\n\x1a\n"):
        return False
    position = 8
    saw_iend = False
    while position + 12 <= len(raw):
        length = struct.unpack_from(">I", raw, position)[0]
        end = position + 12 + length
        if end > len(raw):
            return False
        chunk_type = raw[position + 4 : position + 8]
        chunk_data = raw[position + 8 : position + 8 + length]
        expected_crc = struct.unpack_from(">I", raw, position + 8 + length)[0]
        if zlib.crc32(chunk_type + chunk_data) & 0xFFFFFFFF != expected_crc:
            return False
        position = end
        if chunk_type == b"IEND":
            saw_iend = length == 0
            break
    return saw_iend and position == len(raw)


def _terminal_structure_complete(raw, image_format):
    if image_format == "jpeg":
        return raw.rstrip(b"\x00").endswith(b"\xff\xd9")
    if image_format == "png":
        return _png_complete(raw)
    if image_format == "gif":
        return raw.endswith(b";")
    if image_format == "bmp":
        return len(raw) >= 6 and int.from_bytes(raw[2:6], "little") <= len(raw)
    if image_format == "webp":
        return len(raw) >= 12 and int.from_bytes(raw[4:8], "little") + 8 <= len(raw)
    return True


def validate_image_bytes(raw, image_format):
    if not raw:
        return "undecodable", "empty decoded media"
    if image_format == "svg":
        try:
            root = ET.fromstring(raw)
        except ET.ParseError as exc:
            return "undecodable", f"invalid SVG XML: {exc}"
        if root.tag.rsplit("}", 1)[-1].lower() != "svg":
            return "undecodable", "XML root is not svg"
        return "decodable_structurally_complete", "well-formed SVG"
    if image_format not in {"jpeg", "png", "gif", "bmp", "tiff", "webp"}:
        return "unverified_format", "no supported decoder for detected/original format"

    ImageFile.LOAD_TRUNCATED_IMAGES = False
    try:
        with Image.open(io.BytesIO(raw)) as image:
            image.load()
    except Exception as exc:
        return "undecodable", f"Pillow load failed: {exc}"
    if not _terminal_structure_complete(raw, image_format):
        return "decodable_structurally_incomplete", "terminal structure missing or inconsistent"
    return "decodable_structurally_complete", "full decode and terminal structure"


def content_type_map(entries):
    xml = next(
        (
            item.get("decoded_xml", "")
            for item in entries
            if item.get("local_file_name") == "[Content_Types].xml"
        ),
        "",
    )
    defaults = {}
    overrides = {}
    if xml:
        try:
            root = ET.fromstring(xml)
            for node in root:
                local = node.tag.rsplit("}", 1)[-1]
                if local == "Default":
                    defaults[node.attrib.get("Extension", "").lower()] = node.attrib.get("ContentType")
                elif local == "Override":
                    overrides[node.attrib.get("PartName", "").lstrip("/")] = node.attrib.get("ContentType")
        except ET.ParseError:
            pass
    result = {}
    for item in entries:
        name = item.get("local_file_name", "")
        extension = Path(name).suffix.lower().lstrip(".")
        result[name] = overrides.get(name) or defaults.get(extension)
    return result


_FORMAT_VALIDATION = {
    "decodable_structurally_complete": "decodable_complete",
    "decodable_structurally_incomplete": "decodable_incomplete",
    "undecodable": "undecodable",
    "unverified_format": "unsupported_format",
}
_IMAGE_BUCKET = {
    "preserved_verified": "verified",
    "preserved_decodable": "decodable",
    "preserved_unverified": "unverified",
    "preserved_incomplete": "incomplete",
}


def _image_summary_status(integrity, format_validation, recognized):
    """Summary image status (evaluation_definitions_v2 §2, per-image status)."""
    if integrity == "crc_verified":
        return "preserved_verified"
    if integrity == "crc_unavailable" and format_validation == "decodable_complete":
        return "preserved_decodable"
    if integrity == "crc_unavailable" and format_validation == "unsupported_format" and not recognized:
        return "preserved_unverified"
    return "preserved_incomplete"


def extract_img_file(entries, output_path):
    """Validate and preserve decoded media without inventing an extension.

    Rejected bytes (content_validation_status=rejected) are never written; only their
    payload offset and size are recorded.
    """
    output_root = Path(output_path)
    types = content_type_map(entries)
    records = []
    used_paths = set()
    for item in entries:
        name = item.get("local_file_name", "")
        if "/media/" not in name or name.endswith("/"):
            continue
        validation = item.get("content_validation_status")
        record = {
            "name": name,
            "declared_content_type": types.get(name),
            "decode_status": item.get("decode_status"),
            "content_validation_status": validation,
            "source_offset": item.get("source_offset"),
            "payload_offset": item.get("data_start"),
            "output": None,
        }
        hex_data = item.get("decoded_media_hex", "")
        if not hex_data:
            rejected_bytes = item.get("rejected_bytes", 0)
            record.update(
                status="rejected" if rejected_bytes else "lost",
                integrity=None,
                format_validation=None,
                detected_format=None,
                bytes=0,
                rejected_bytes=rejected_bytes,
                sha256=None,
                reason=item.get("content_validation_reason") or item.get("decode_error") or "no decoded media bytes",
                diagnostic_output=None,
                diagnostic_sha256=None,
            )
            # Decision 2026-10-13: rejected STORED image bytes are kept as a diagnostic .bin.
            # The record stays "rejected" and is not counted as a preserved image.
            diagnostic_hex = item.get("rejected_stored_image_hex")
            if diagnostic_hex:
                diagnostic = bytes.fromhex(diagnostic_hex)
                target_dir = output_root / "rejected_stored"
                target_dir.mkdir(parents=True, exist_ok=True)
                filename = sanitize_part_basename(name) + ".bin"
                candidate = target_dir / filename
                counter = 1
                while candidate.as_posix().lower() in used_paths or candidate.exists():
                    candidate = target_dir / f"{Path(filename).stem}_{counter}.bin"
                    counter += 1
                candidate.write_bytes(diagnostic)
                used_paths.add(candidate.as_posix().lower())
                record.update(
                    diagnostic_output=candidate.relative_to(output_root.parent).as_posix(),
                    diagnostic_sha256=hashlib.sha256(diagnostic).hexdigest(),
                )
            records.append(record)
            continue
        raw = bytes.fromhex(hex_data)
        detected = detect_image_format(raw)
        validation_detail, reason = validate_image_bytes(raw, detected)
        format_validation = _FORMAT_VALIDATION[validation_detail]
        integrity = {"crc_verified": "crc_verified", "crc_mismatch": "crc_mismatch"}.get(validation, "crc_unavailable")
        recognized = detected in {"jpeg", "png", "gif", "bmp", "tiff", "webp", "svg"}
        status = _image_summary_status(integrity, format_validation, recognized)

        target_dir = output_root / _IMAGE_BUCKET[status]
        target_dir.mkdir(parents=True, exist_ok=True)
        filename = sanitize_part_basename(name)
        candidate = target_dir / filename
        counter = 1
        while candidate.as_posix().lower() in used_paths or candidate.exists():
            candidate = target_dir / f"{Path(filename).stem}_{counter}{Path(filename).suffix}"
            counter += 1
        candidate.write_bytes(raw)
        used_paths.add(candidate.as_posix().lower())
        record.update(
            status=status,
            integrity=integrity,
            format_validation=format_validation,
            detected_format=detected,
            bytes=len(raw),
            sha256=hashlib.sha256(raw).hexdigest(),
            reason=reason,
            output=candidate.relative_to(output_root.parent).as_posix(),
        )
        records.append(record)
    return records


def aggregate_image_status(records, inventory_status, reported_missing_media=()):
    """Aggregate image status (evaluation_definitions_v2 §2)."""
    known_missing = len(reported_missing_media)
    if not records and not known_missing:
        return "unknown" if inventory_status == "unknown" else "none"
    preserved = [record for record in records if str(record.get("status", "")).startswith("preserved_")]
    if not preserved:
        return "failed"
    if len(preserved) == len(records) and not known_missing:
        return "all_preserved"
    return "partial"


def image_counts(records, reported_missing_media=()):
    total = len(records) + len(reported_missing_media)
    preserved = sum(str(record.get("status", "")).startswith("preserved_") for record in records)
    verified = sum(record.get("integrity") == "crc_verified" for record in records)
    return {
        "images_total": total,
        "images_preserved": preserved,
        "images_integrity_verified": verified,
        "verified_ratio": (verified / total) if total else None,
    }


def extract_metadata(entries, output_dir):
    tags = {
        "dc:creator": "creator",
        "cp:lastModifiedBy": "lastModifiedBy",
        "cp:revision": "revision",
        "dcterms:created": "created",
        "dcterms:modified": "modified",
    }
    results = ["[Document Metadata]"]
    for item in entries:
        if item.get("local_file_name") != "docProps/core.xml":
            continue
        # Only validated decoded XML is used; compressed bytes are never parsed as metadata.
        xml_raw = item.get("decoded_xml") or ""
        found = False
        for tag, label in tags.items():
            match = re.search(f"<{tag}.*?>(.*?)</{tag}>", xml_raw)
            if match and match.group(1):
                found = True
                value = match.group(1)
                results.append(f"{label} : {value} (<{tag}>{value}</{tag}>)")
        if not found:
            results = ["Metadata is severely damaged or missing."]
    metadata_path = Path(output_dir) / "metadata" / "metadata.txt"
    metadata_path.parent.mkdir(parents=True, exist_ok=True)
    metadata_path.write_text("\n".join(results), encoding="utf-8")
    safe_print(f"[INFO] Metadata saved to {metadata_path}")
    return str(metadata_path)
