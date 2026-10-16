"""Output, image-validation, filename, XML, and metadata helpers."""

from __future__ import annotations

import hashlib
import io
import os
import re
import struct
import xml.etree.ElementTree as ET
import zlib
from pathlib import Path

from lxml import etree
from PIL import Image, ImageFile


def safe_print(message):
    try:
        print(message)
    except UnicodeEncodeError:
        print(str(message).encode("ascii", errors="replace").decode("ascii"))


# ------------------------------------------------------------------- output folders

_WINDOWS_RESERVED = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}


def _safe_segment(segment):
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", segment).rstrip(" .") or "_"
    if cleaned.split(".")[0].lower() in _WINDOWS_RESERVED:
        cleaned = "_" + cleaned
    return cleaned[:180]


def new_output_directory(parent, file_path):
    """Create output_<file name with extension>; never reuse an existing folder (adds _2, _3, ...)."""
    parent = Path(parent)
    base = "output_" + _safe_segment(os.path.basename(file_path) or "recovered")
    counter = 1
    while True:
        candidate = parent / (base if counter == 1 else f"{base}_{counter}")
        try:
            candidate.mkdir(parents=True)
            return candidate
        except FileExistsError:
            counter += 1


def write_part(parts_root, name, data, used):
    """Write a recovered part below parts_root under its original path; return the written path.

    Segments are made safe for Windows. Names that collide case-insensitively get _2, _3, ...;
    a path blocked by an existing file of the same name is written flat under _conflicts/."""
    parts_root = Path(parts_root)
    segments = [_safe_segment(segment) for segment in name.split("/") if segment]
    target = parts_root.joinpath(*segments)
    stem, suffix = target.stem, target.suffix
    counter = 2
    while target.as_posix().lower() in used or target.exists():
        target = target.with_name(f"{stem}_{counter}{suffix}")
        counter += 1
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
    except (FileExistsError, NotADirectoryError):
        target = parts_root / "_conflicts" / "_".join(segments)
        while target.as_posix().lower() in used or target.exists():
            target = target.with_name(f"{stem}_{counter}{suffix}")
            counter += 1
        target.parent.mkdir(parents=True, exist_ok=True)
    target.resolve().relative_to(parts_root.resolve())  # raises if the path left parts_root
    target.write_bytes(data)
    used.add(target.as_posix().lower())
    return target


# ---------------------------------------------------------------------------- XML

_XML_PARSER = etree.XMLParser(recover=True, resolve_entities=False, no_network=True, huge_tree=True)


def parse_xml(text):
    """Parse decoded XML leniently (validated prefixes are incomplete); None if nothing parses."""
    if not text:
        return None
    data = text.encode("utf-8") if isinstance(text, str) else text
    try:
        return etree.fromstring(data, _XML_PARSER)
    except (etree.XMLSyntaxError, ValueError):
        return None


def local_name(element):
    """Element name without namespace or prefix ('' for comments and processing instructions).

    In recovery mode an undeclared prefix stays in the tag ('p:sld'), so it is removed as well."""
    if not isinstance(element.tag, str):
        return ""
    return element.tag.rsplit("}", 1)[-1].rsplit(":", 1)[-1]


STATUS_NOTES = {
    "crc_verified": "byte-identical to the original (CRC-32 verified)",
    "structure_verified": "no CRC-32 reference; complete and well-formed",
    "prefix_validated": "start of the part only; the end may contain non-original text",
    "crc_mismatch": "CRC-32 differs from the original; content may be altered (invalid UTF-8 shown as U+FFFD)",
    "unverified": "no CRC-32 reference and not well-formed; content may be altered",
}


def source_line(item, label="source"):
    """One-line provenance header for text output: part name, validation status, and its meaning."""
    status = item.get("content_validation_status")
    return f"# {label}: {item.get('local_file_name')} | {status} | {STATUS_NOTES.get(status, '')}".rstrip(" |")


def first_with_content(entries, name):
    """First entry with decoded XML for a part name (duplicates can occur in damaged files)."""
    return next((item for item in entries if item.get("local_file_name") == name and item.get("decoded_xml")), None)


def write_relationships(entries, output_dir):
    """Write every recovered .rels part, grouped by source part, to relationships.txt."""
    blocks = []
    for item in entries:
        if not item.get("local_file_name", "").endswith(".rels") or not item.get("decoded_xml"):
            continue
        lines = [source_line(item)]
        root = parse_xml(item["decoded_xml"])
        for relationship in root.iter() if root is not None else ():
            if local_name(relationship) != "Relationship":
                continue
            mode = relationship.get("TargetMode")
            lines.append(f"{relationship.get('Id')}: {relationship.get('Target')} ({relationship.get('Type')})"
                         + (f" [{mode}]" if mode else ""))
        blocks.append("\n".join(lines))
    if blocks:
        (Path(output_dir) / "relationships.txt").write_text("\n\n".join(blocks) + "\n", encoding="utf-8")


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

    Rejected bytes (content_validation_status=rejected) are not written; their payload offset and
    size are recorded. A rejected STORED image is the exception: its bytes are written as a
    diagnostic .bin, unless they are one repeated byte value (e.g. a zero-filled area), which carries
    no content and is recorded in the manifest as uniform_fill_byte instead.
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
                uniform_fill_byte=None,
            )
            # Decision 2026-10-13: rejected STORED image bytes are kept as a diagnostic .bin.
            # The record stays "rejected" and is not counted as a preserved image.
            # Decision 2026-10-16: bytes that are one repeated value are not written (no content);
            # the value is recorded so that the fill itself remains visible in the manifest.
            diagnostic_hex = item.get("rejected_stored_image_hex")
            diagnostic = bytes.fromhex(diagnostic_hex) if diagnostic_hex else b""
            if diagnostic and diagnostic.count(diagnostic[:1]) == len(diagnostic):
                record["uniform_fill_byte"] = f"0x{diagnostic[0]:02X}"
            elif diagnostic:
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


_METADATA_FIELDS = {
    "docProps/core.xml": ("title", "subject", "creator", "keywords", "description", "lastModifiedBy", "revision",
                          "created", "modified", "lastPrinted", "category", "contentStatus", "identifier",
                          "language", "version"),
    "docProps/app.xml": ("Application", "AppVersion", "Company", "Manager", "Template", "TotalTime", "Pages",
                         "Words", "Characters", "Lines", "Paragraphs", "Slides", "Notes", "HiddenSlides"),
}


def extract_metadata(entries, output_dir):
    """Write metadata.txt from docProps/core.xml and docProps/app.xml; return the values per part.

    Only decoded XML that passed content validation is read; each block names its source status."""
    lines = []
    result = {}
    for part, fields in _METADATA_FIELDS.items():
        item = first_with_content(entries, part)
        if item is None:
            lines += [f"# {part}: not recovered", ""]
            continue
        values = {}
        root = parse_xml(item["decoded_xml"])
        for element in root if root is not None else ():
            name = local_name(element)
            if name in fields and (element.text or "").strip() and name not in values:
                values[name] = element.text
        lines.append(source_line(item))
        lines += [f"{name}: {value}" for name, value in values.items()]
        lines.append("")
        result[part] = {"content_validation_status": item.get("content_validation_status"), "values": values}
    metadata_path = Path(output_dir) / "metadata.txt"
    metadata_path.write_text("\n".join(lines), encoding="utf-8")
    return result
