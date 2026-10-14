"""CD-tolerant OOXML extraction with explicit spans, two-axis validation, and status v2.

Pipeline (see evaluation_definitions_v2.md):
1. Parse central-directory (CD) records wherever they survive; trust a record for a
   local header only when its offset and name match.
2. Discover local file headers (LFH) sequentially. Jump over a payload only when its
   end is verified (CRC-32 or declared size); otherwise keep searching inside it.
   Afterwards re-scan the whole file and merge plausible LFHs that were skipped,
   excluding those inside CRC-verified payloads (e.g. embedded packages).
3. Bound STORED payloads by CD -> LFH sizes -> data descriptor -> next header,
   preferring the first boundary whose CRC-32 verifies.
4. Decode, validate content, write recovered parts, text, and media, and write the
   recovery manifest (schema v3).
"""

from __future__ import annotations

import hashlib
import io
import json
import re
import struct
import zipfile
import zlib
from collections import namedtuple
from pathlib import Path

from .. import __version__
from .decoder import (
    CONTENT_YIELDING_VALIDATION,
    MAX_DECODED_BYTES,
    OK_VALIDATION,
    decode_local_file_data,
    inflate_probe,
    is_media_part,
)
from .handlers import docx, pptx, xlsx
from .fragment_salvage import salvage_sources
from .utils import (
    aggregate_image_status,
    extract_img_file,
    extract_metadata,
    image_counts,
    new_output_directory,
    safe_print,
    write_part,
    write_relationships,
)


LFH_SIG = b"PK\x03\x04"
CD_SIG = b"PK\x01\x02"
EOCD_SIG = b"PK\x05\x06"
DD_SIG = b"PK\x07\x08"
MAX_NAME_BYTES = 1024
_FALLBACK_NAME_BYTES = 512

# Case-insensitive known extensions. Used to (a) cut a declared name that absorbed
# payload bytes and (b) recover a name when the declared length is invalid.
_KNOWN_EXTENSIONS = (
    "xml.rels|rels|xml|png|jpeg|jpg|jpe|jfif|gif|bmp|dib|svg|emf|wmf|emz|wmz|wdp|jxr|hdp|"
    "tif|tiff|webp|ico|wav|mp3|m4a|wma|aac|mp4|m4v|mov|avi|wmv|mpg|mpeg|bin|fntdata|odttf|"
    "ttf|otf|vml|xlsx|xlsm|xlsb|xls|docx|docm|doc|pptx|pptm|ppt|vsdx|sldx|pdf|txt|json|sigs|"
    "psdsxs|data|dat|odt|ods|odp|html|htm|css|js"
)
_SEGMENT = rb'[^\x00-\x1f\x7f/\\:*?"<>|]'
_FALLBACK_NAME = re.compile(
    rb"(?:\[Content_Types\]\.xml|_rels/\.rels|ppt/metadata|(?:" + _SEGMENT + rb"+/)*"
    + _SEGMENT + rb"+?\.(?:" + _KNOWN_EXTENSIONS.encode() + rb"))",
    re.IGNORECASE,
)
_KNOWN_EXT_SET = set(_KNOWN_EXTENSIONS.split("|"))
_SPECIAL_NAMES = {"[Content_Types].xml", "_rels/.rels", "ppt/metadata"}

_CORE_PATTERNS = {
    ".docx": (re.compile(r"^word/document\.xml$"),),
    ".xlsx": (re.compile(r"^xl/(?:worksheets/sheet\d+|sharedStrings)\.xml$"),),
    ".pptx": (re.compile(r"^ppt/slides/slide\d+\.xml$"),),
}
_CONTENT_PATTERNS = {
    ".docx": (
        re.compile(r"^word/document\.xml$"),
        re.compile(r"^word/header\d+\.xml$"),
        re.compile(r"^word/footer\d+\.xml$"),
        re.compile(r"^word/(?:footnotes|endnotes|comments)\.xml$"),
        re.compile(r"^word/glossary/document\.xml$"),
    ),
    ".xlsx": (re.compile(r"^xl/(?:worksheets/sheet\d+|sharedStrings)\.xml$"),),
    ".pptx": (
        re.compile(r"^ppt/slides/slide\d+\.xml$"),
        re.compile(r"^ppt/notesSlides/notesSlide\d+\.xml$"),
    ),
}


# --------------------------------------------------------------------------- names


def decode_part_name(raw, flags=0):
    """Return a safe part name decoded per the ZIP UTF-8 flag (bit 11) or CP437, else None."""
    if not raw or len(raw) > MAX_NAME_BYTES:
        return None
    if any(byte < 0x20 or byte == 0x7F for byte in raw):
        return None
    try:
        name = raw.decode("utf-8")
    except UnicodeDecodeError:
        if flags & 0x800:
            return None
        name = raw.decode("cp437")
    if name.startswith(("/", "\\")) or re.match(r"^[A-Za-z]:", name) or "\\" in name:
        return None
    if any(char in '<>:"|?*' for char in name):
        return None
    segments = name.split("/")
    if any(segment in ("..", ".") for segment in segments) or "" in segments[:-1]:
        return None
    return name


def _extension_known(name):
    if name in _SPECIAL_NAMES or name.endswith("/"):
        return True
    lowered = name.lower()
    if lowered.endswith(".xml.rels"):
        return True
    if "." not in lowered.rsplit("/", 1)[-1]:
        return False
    return lowered.rsplit(".", 1)[-1] in _KNOWN_EXT_SET


def recover_filename_length(part, cd_name_raw=None):
    """Return the filename length for an LFH slice (0 if no plausible name).

    Order: trusted CD name -> declared length (UTF-8/CP437, safe path) -> grammar
    fallback. A declared name with an unknown extension is cut back only when a
    known-extension prefix exists (absorbed payload bytes).
    """
    if len(part) < 30:
        return 0
    flags = int.from_bytes(part[6:8], "little")
    if cd_name_raw and part[30 : 30 + len(cd_name_raw)] == cd_name_raw:
        return len(cd_name_raw)
    declared = int.from_bytes(part[26:28], "little")
    if 0 < declared <= MAX_NAME_BYTES and 30 + declared <= len(part):
        declared_raw = part[30 : 30 + declared]
        name = decode_part_name(declared_raw, flags)
        if name is not None:
            if _extension_known(name):
                return declared
            match = _FALLBACK_NAME.match(declared_raw)
            if match and len(match.group()) < declared:
                return len(match.group())
            return declared
    match = _FALLBACK_NAME.match(part[30 : 30 + _FALLBACK_NAME_BYTES])
    if match and decode_part_name(match.group(), flags) is not None:
        return len(match.group())
    return 0


# ------------------------------------------------------------------ central directory


def parse_central_directory(data):
    """Return CD records keyed by (adjusted) local-header offset.

    Records are read wherever a CD signature survives. If an EOCD is present, the
    offset delta of a prefixed archive is applied (as zipfile does).
    """
    delta = 0
    eocd = data.rfind(EOCD_SIG)
    if eocd >= 0 and eocd + 22 <= len(data):
        cd_size, cd_offset = struct.unpack_from("<II", data, eocd + 12)
        if cd_offset + cd_size <= eocd:
            delta = eocd - cd_size - cd_offset
    records = {}
    position = 0
    while True:
        index = data.find(CD_SIG, position)
        if index < 0 or index + 46 > len(data):
            break
        position = index + 4
        (_, _, flags, method, _, _, crc, csize, usize, name_len, extra_len, comment_len, _, _, _,
         local_offset) = struct.unpack_from("<HHHHHHIIIHHHHHII", data, index + 4)
        raw_name = data[index + 46 : index + 46 + name_len]
        name = decode_part_name(raw_name, flags)
        if name is None or method not in (0, 8):
            continue
        records.setdefault(local_offset + delta, []).append(
            {"name": name, "raw_name": raw_name, "flags": flags, "method": method, "crc": crc, "csize": csize,
             "usize": usize}
        )
    return records


def _trusted_cd(cd_records, offset, data):
    """A CD record is trusted for an LFH only if offset and name bytes match."""
    if data[offset : offset + 4] != LFH_SIG:
        return None
    for record in cd_records.get(offset, []):
        raw_name = record["raw_name"]
        if data[offset + 30 : offset + 30 + len(raw_name)] == raw_name:
            return record
    return None


def _central_directory_inventory(data, discovered_names):
    if EOCD_SIG not in data or CD_SIG not in data:
        return "unknown", [], []
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            reported = [info.filename for info in archive.infolist()]
    except (OSError, ValueError, zipfile.BadZipFile):
        return "unknown", [], []
    missing = sorted(set(reported) - set(discovered_names))
    extras = sorted(set(discovered_names) - set(reported))
    status = "complete" if not missing and not extras and len(reported) == len(discovered_names) else "incomplete"
    return status, missing, extras


# --------------------------------------------------------------------------- headers


def _valid_size(value, start, limit):
    return 0 < value < 0xFFFFFFFF and start + value <= limit


def _parse_lfh(data, offset, cd_records):
    part = data[offset : offset + 30 + MAX_NAME_BYTES]
    cd = _trusted_cd(cd_records, offset, data)
    name_length = recover_filename_length(part, cd["raw_name"] if cd else None)
    if not name_length:
        return None
    flags = int.from_bytes(part[6:8], "little")
    name = decode_part_name(part[30 : 30 + name_length], flags)
    if name is None:
        return None
    extra_length = int.from_bytes(part[28:30], "little")
    data_start = offset + 30 + name_length + extra_length
    if data_start > len(data):
        return None
    return {
        "offset": offset,
        "name": name,
        "flags": flags,
        "method": int.from_bytes(part[8:10], "little"),
        "crc": int.from_bytes(part[14:18], "little"),
        "csize": int.from_bytes(part[18:22], "little"),
        "usize": int.from_bytes(part[22:26], "little"),
        "data_start": data_start,
        "cd": cd,
    }


def _plausible_lfh_offsets(data, start=0):
    position = start
    while position <= len(data) - 30:
        offset = data.find(LFH_SIG, position)
        if offset < 0:
            return
        position = offset + 4
        if int.from_bytes(data[offset + 8 : offset + 10], "little") not in (0, 8):
            continue
        if recover_filename_length(data[offset : offset + 30 + MAX_NAME_BYTES]) == 0:
            continue
        yield offset


def _next_lfh_offset(data, start):
    return next(_plausible_lfh_offsets(data, start), None)


def _crc_refs(header, descriptor=None):
    refs = set()
    if header.get("cd"):
        refs.add(header["cd"]["crc"])
    lfh_crc_usable = not (header["flags"] & 0x8 and header["crc"] == 0)
    if lfh_crc_usable:
        refs.add(header["crc"])
    if descriptor:
        refs.add(descriptor["crc"])
    return refs


def _size_refs(header, descriptor=None):
    refs = set()
    if header.get("cd") and header["cd"]["usize"] < 0xFFFFFFFF:
        refs.add(header["cd"]["usize"])
    if 0 < header["usize"] < 0xFFFFFFFF:
        refs.add(header["usize"])
    if descriptor:
        refs.add(descriptor["usize"])
    return refs


def _deflate_descriptor(data, end, consumed):
    """Data descriptor right after a DEFLATE stream end (optional signature)."""
    if data[end : end + 4] == DD_SIG and end + 16 <= len(data):
        crc, csize, usize = struct.unpack_from("<III", data, end + 4)
        if csize == consumed:
            return {"crc": crc, "csize": csize, "usize": usize, "length": 16}
    if end + 12 <= len(data):
        crc, csize, usize = struct.unpack_from("<III", data, end)
        if csize == consumed:
            return {"crc": crc, "csize": csize, "usize": usize, "length": 12}
    return None


def _stored_descriptors(data, data_start, limit):
    """Candidate STORED descriptors: signed ones anywhere, unsigned ones just before limit."""
    found = []
    position = data_start
    while True:
        index = data.find(DD_SIG, position, limit)
        if index < 0:
            break
        position = index + 1
        if index + 16 <= len(data):
            crc, csize, usize = struct.unpack_from("<III", data, index + 4)
            if csize == usize == index - data_start:
                found.append({"crc": crc, "csize": csize, "usize": usize, "length": 16})
    for back in (12, 16):
        index = limit - back
        if index >= data_start and index + 12 <= len(data):
            crc, csize, usize = struct.unpack_from("<III", data, index if back == 12 else index + 4)
            if csize == usize == index - data_start:
                found.append({"crc": crc, "csize": csize, "usize": usize, "length": back})
    return found


def stored_boundary(data, header, candidate_end):
    """Choose a STORED payload length (evaluation_definitions_v2 §1, boundary_status)."""
    start = header["data_start"]
    candidates = []
    cd = header.get("cd")
    if cd and start + cd["csize"] <= len(data) and cd["csize"] == cd["usize"]:
        candidates.append(("central_directory", cd["csize"]))
    csize, usize = header["csize"], header["usize"]
    if _valid_size(csize, start, len(data)) and _valid_size(usize, start, len(data)) and csize == usize:
        candidates.append(("lfh_consistent_sizes", csize))
    else:
        if _valid_size(csize, start, len(data)):
            candidates.append(("lfh_compressed_size_only", csize))
        if _valid_size(usize, start, len(data)):
            candidates.append(("lfh_uncompressed_size_only", usize))
    descriptors = _stored_descriptors(data, start, candidate_end) if header["flags"] & 0x8 else []
    for descriptor in descriptors:
        candidates.append(("data_descriptor", descriptor["csize"]))
    candidates.append(("next_header_fallback", max(0, candidate_end - start)))

    refs = _crc_refs(header) | {descriptor["crc"] for descriptor in descriptors}
    if refs:
        for label, length in candidates:
            if zlib.crc32(data[start : start + length]) & 0xFFFFFFFF in refs:
                return length, label, True, descriptors
    for label, length in candidates:
        if start + length <= candidate_end or label == "next_header_fallback":
            return length, label, False, descriptors
    return 0, "stored_unbounded", False, descriptors


# --------------------------------------------------------------------------- discovery


def discover_entries(data, cd_records=None):
    """Sequential discovery with verified jumps, then a merging re-scan."""
    cd_records = parse_central_directory(data) if cd_records is None else cd_records
    headers = {}
    verified_spans = []
    offset = _next_lfh_offset(data, 0)
    while offset is not None:
        header = _parse_lfh(data, offset, cd_records)
        if header is None:
            offset = _next_lfh_offset(data, offset + 4)
            continue
        header["discovery"] = "sequential"
        headers[offset] = header
        start = header["data_start"]
        search_from = start
        if header["method"] == 8 and start < len(data):
            probe = inflate_probe(data[start:])
            if probe["status"] == "complete":
                end = start + probe["consumed"]
                descriptor = _deflate_descriptor(data, end, probe["consumed"]) if header["flags"] & 0x8 else None
                crc_ok = probe["crc"] in _crc_refs(header, descriptor)
                size_ok = probe["consumed"] in {
                    value for value in (header["csize"], (header["cd"] or {}).get("csize"))
                    if value and value < 0xFFFFFFFF
                }
                if crc_ok or size_ok:
                    search_from = end + (descriptor["length"] if descriptor else 0)
                if crc_ok:
                    verified_spans.append((start, end))
        elif header["method"] == 0:
            next_guess = _next_lfh_offset(data, start)
            limit = next_guess if next_guess is not None else len(data)
            length, label, crc_ok, _ = stored_boundary(data, header, limit)
            if crc_ok:
                verified_spans.append((start, start + length))
                search_from = start + length
            elif label in ("central_directory", "lfh_consistent_sizes"):
                search_from = start + length
        next_offset = _next_lfh_offset(data, max(search_from, offset + 4))
        if next_offset is None or next_offset <= offset:
            break
        offset = next_offset

    # Re-scan: merge plausible LFHs skipped by unverified jumps.
    for candidate in _plausible_lfh_offsets(data, 0):
        if candidate in headers:
            continue
        if any(start <= candidate < end for start, end in verified_spans):
            continue
        header = _parse_lfh(data, candidate, cd_records)
        if header is not None:
            header["discovery"] = "rescan"
            headers[candidate] = header

    offsets = sorted(headers)
    cd_start = data.find(CD_SIG, offsets[-1] + 30) if offsets else -1
    entries = []
    for index, offset in enumerate(offsets):
        header = headers[offset]
        candidate_end = offsets[index + 1] if index + 1 < len(offsets) else len(data)
        if index + 1 == len(offsets) and cd_start >= 0:
            candidate_end = min(candidate_end, cd_start)
        start = header["data_start"]
        candidate_end = max(candidate_end, start)
        descriptor = None
        if header["method"] == 0:
            length, boundary_status, _, descriptors = stored_boundary(data, header, candidate_end)
            payload = data[start : start + length]
            descriptor = next((item for item in descriptors if item["csize"] == length), None)
        else:
            payload = data[start:candidate_end]
            boundary_status = "candidate_to_next_lfh"
            if header["flags"] & 0x8:
                probe = inflate_probe(payload)
                if probe["status"] == "complete":
                    descriptor = _deflate_descriptor(data, start + probe["consumed"], probe["consumed"])
        cd = header["cd"]
        declared_valid = 0 < header["csize"] < 0xFFFFFFFF and header["csize"] <= candidate_end - start
        entries.append(
            {
                "local_file_name": header["name"],
                "local_file_data": payload.hex(),
                "compressed_data_hex": payload.hex(),
                "extension": Path(header["name"]).suffix.lower() or None,
                "compression_method": header["method"],
                "general_purpose_flags": header["flags"],
                "declared_compressed_size": header["csize"],
                "uncompressed_size": header["usize"],
                "declared_crc32": header["crc"],
                "trusted_central_directory": bool(cd),
                "discovery": header["discovery"],
                "source_offset": offset,
                "data_start": start,
                "candidate_end": candidate_end,
                "declared_span": [start, start + header["csize"]] if declared_valid else None,
                "decoded_span": None,
                "residual_span": None,
                "boundary_status": boundary_status,
                "declared_size_consistency": "not_checked",
                "crc_refs": sorted(_crc_refs(header, descriptor)),
                "size_refs": sorted(_size_refs(header, descriptor)),
                "_header": header,
            }
        )
    return entries, offsets


def _finalize_spans(entries):
    orphan_spans = []
    for item in entries:
        data_start = item["data_start"]
        candidate_end = item["candidate_end"]
        consumed = item.get("compressed_consumed", 0)
        method = item.get("compression_method")
        used_end = data_start
        if method == 8 and consumed:
            used_end = min(candidate_end, data_start + consumed)
            item["decoded_span"] = [data_start, used_end]
            if item.get("decode_status") == "complete":
                item["boundary_status"] = "decoded_eof"
        elif method == 0:
            used_end = min(candidate_end, data_start + consumed)
            item["decoded_span"] = [data_start, used_end]

        declared = item.get("declared_span")
        if declared and item.get("decoded_span"):
            item["declared_size_consistency"] = "match" if declared[1] == item["decoded_span"][1] else "mismatch"
        elif item.get("declared_compressed_size") in (0, 0xFFFFFFFF):
            item["declared_size_consistency"] = "unavailable"
        else:
            item["declared_size_consistency"] = "out_of_bounds"

        if used_end < candidate_end:
            item["residual_span"] = [used_end, candidate_end]
            orphan_spans.append(
                {"kind": "entry_residual", "entry": item["local_file_name"], "span": item["residual_span"],
                 "bytes": candidate_end - used_end}
            )
    return orphan_spans


# --------------------------------------------------------------------------- statuses


def _matches(name, patterns):
    return any(pattern.match(name) for pattern in patterns)


def _infer_extension(names):
    for prefix, extension in (("word/", ".docx"), ("xl/", ".xlsx"), ("ppt/", ".pptx")):
        if any(name.startswith(prefix) for name in names):
            return extension
    return None


def _core_status(entries, extension):
    core = [item for item in entries if _matches(item["local_file_name"], _CORE_PATTERNS.get(extension, ()))]
    if not core:
        return "missing"
    validations = [item.get("content_validation_status") for item in core]
    if all(value in OK_VALIDATION for value in validations):
        return "complete"
    if any(value in CONTENT_YIELDING_VALIDATION for value in validations):
        return "partial"
    return "failed"


def _document_content_status(entries, extension, inventory_missing):
    patterns = _CONTENT_PATTERNS.get(extension, ())
    content = [item for item in entries if _matches(item["local_file_name"], patterns)]
    reported_missing = [name for name in inventory_missing if _matches(name, patterns)]
    if not content and not reported_missing:
        return "missing"
    validations = [item.get("content_validation_status") for item in content]
    if content and not reported_missing and all(value in OK_VALIDATION for value in validations):
        return "complete"
    if any(value in CONTENT_YIELDING_VALIDATION for value in validations):
        return "partial"
    return "failed"


def _entry_failed(item):
    return item.get("content_validation_status") not in OK_VALIDATION + ("directory",)


def _overall_status(inventory_status, entries, any_content, any_image):
    if inventory_status == "complete" and not any(_entry_failed(item) for item in entries):
        return "complete"
    if any_content or any_image:
        return "partial"
    return "indeterminate" if inventory_status == "unknown" else "failed"


def _json_entry(item):
    excluded = {"local_file_data", "compressed_data_hex", "decoded_xml", "decoded_media_hex", "candidate_end",
                "data_start", "rejected_stored_image_hex"}
    return {key: value for key, value in item.items() if key not in excluded and not key.startswith("_")}


def tool_code_sha256():
    """SHA-256 over sorted ort/**/*.py relative paths and contents."""
    root = Path(__file__).resolve().parents[1]
    digest = hashlib.sha256()
    for path in sorted(root.rglob("*.py")):
        relative = path.relative_to(root).as_posix()
        if "__pycache__" in relative:
            continue
        digest.update(relative.encode())
        digest.update(b"\0")
        digest.update(path.read_bytes())
        digest.update(b"\0")
    return digest.hexdigest()


# --------------------------------------------------------------------------- main entry

RecoveryResult = namedtuple("RecoveryResult", "entries extension output_directory overall_status")
OUTPUT_LAYOUT = 2


def _write_parts(entries, output_directory):
    """Write every content-yielding non-media part to parts/<original path> (rejected bytes never)."""
    used = set()
    for item in entries:
        raw = item.get("_decoded_bytes")
        name = item["local_file_name"]
        if not raw or is_media_part(name) or item.get("content_validation_status") not in CONTENT_YIELDING_VALIDATION:
            continue
        target = write_part(output_directory / "parts", name, raw, used)
        item["output"] = target.relative_to(output_directory).as_posix()
        item["output_sha256"] = hashlib.sha256(raw).hexdigest()


def get_file_hex(file_path, *, output_root=None, salvage_fragments=False, dump_entries=False,
                 max_part_size=MAX_DECODED_BYTES):
    """Recover one file: write its output folder and recovery manifest (schema v3).

    Returns RecoveryResult(entries, extension, output_directory, overall_status). When nothing
    is written, overall_status is "missing_input" or "not_ooxml" and the other fields are None.
    """
    input_path = Path(file_path)
    if not input_path.is_file():
        return RecoveryResult(None, None, None, "missing_input")
    data = input_path.read_bytes()
    cd_records = parse_central_directory(data)
    entries, offsets = discover_entries(data, cd_records)
    if not entries:
        safe_print("[PREPROCESSING] No plausible local ZIP header found.")
        return RecoveryResult(None, None, None, "not_ooxml")

    names = [item["local_file_name"] for item in entries]
    inferred_ext = _infer_extension(names)
    if not inferred_ext:
        safe_print("[PREPROCESSING] No OOXML package prefix found.")
        return RecoveryResult(None, None, None, "not_ooxml")
    safe_print(f"[PREPROCESSING] OOXML parts found. Inferred type: {inferred_ext}")

    inventory_status, inventory_missing, inventory_extras = _central_directory_inventory(data, names)
    decode_local_file_data(entries, max_output=max_part_size)
    orphan_spans = _finalize_spans(entries)
    if offsets and offsets[0] > 0:
        orphan_spans.insert(0, {"kind": "unowned_prefix", "entry": None, "span": [0, offsets[0]], "bytes": offsets[0]})
    for item in entries:
        if item.get("content_validation_status") == "prefix_validated":
            # The decoder can emit non-original output between the damage and the point where it fails.
            item["tail_may_be_non_original"] = True

    output_directory = new_output_directory(Path(output_root) if output_root else input_path.parent, input_path.name)
    _write_parts(entries, output_directory)

    reported_missing_media = [name for name in inventory_missing if is_media_part(name)]
    image_records = extract_img_file(entries, output_directory / "img")
    image_status = aggregate_image_status(image_records, inventory_status, reported_missing_media)
    metadata = extract_metadata(entries, output_directory)
    write_relationships(entries, output_directory)

    content_patterns = _CONTENT_PATTERNS[inferred_ext]
    handler = {".docx": docx.process_extracted_docx_data, ".xlsx": xlsx.process_extracted_xlsx_data,
               ".pptx": pptx.process_extracted_pptx_data}[inferred_ext]
    text_records = {record["name"]: record for record in handler(entries, output_directory)}

    content_parts = []
    for item in entries:
        name = item["local_file_name"]
        if not _matches(name, content_patterns) or item.get("content_validation_status") not in CONTENT_YIELDING_VALIDATION:
            continue
        record = {"name": name, "decode_status": item.get("decode_status"),
                  "content_validation_status": item.get("content_validation_status")}
        record.update({key: value for key, value in text_records.get(name, {}).items() if key != "name"})
        content_parts.append(record)

    core_status = _core_status(entries, inferred_ext)
    document_content_status = _document_content_status(entries, inferred_ext, inventory_missing)
    any_image = any(str(record.get("status", "")).startswith("preserved_") for record in image_records)
    overall_status = _overall_status(inventory_status, entries, bool(content_parts), any_image)

    failed_entries = [item for item in entries if _entry_failed(item)]
    content_failed = [item for item in failed_entries if _matches(item["local_file_name"], content_patterns)]
    media_failed = [item for item in failed_entries if is_media_part(item["local_file_name"])]

    fragment_records = []
    if salvage_fragments:
        sources = []
        for item in entries:
            if item.get("content_validation_status") != "rejected":
                continue
            if not item["local_file_name"].endswith((".xml", ".rels")):
                continue
            start, end = item["data_start"], item["candidate_end"]
            sources.append({"category": "failed_payload", "part": item["local_file_name"], "start": start,
                            "data": data[start:end]})
        for orphan in orphan_spans:
            start, end = orphan["span"]
            sources.append({"category": "orphan_span", "part": None, "start": start, "data": data[start:end]})
        fragment_records = salvage_sources(sources, extension=inferred_ext, output_directory=output_directory)

    manifest = {
        "schema_version": 3,
        "output_layout": OUTPUT_LAYOUT,
        "tool_version": __version__,
        "evaluation_definitions": "evaluation_definitions_v2_1",
        "tool_code_sha256": tool_code_sha256(),
        "input": str(input_path.resolve()),
        "input_sha256": hashlib.sha256(data).hexdigest(),
        "inferred_type": inferred_ext,
        "salvage_fragments": bool(salvage_fragments),
        "max_part_size_bytes": max_part_size,
        "inventory_status": inventory_status,
        "inventory_missing_entries": inventory_missing,
        "inventory_extra_entries": inventory_extras,
        "core_part_detected": any(_matches(name, _CORE_PATTERNS[inferred_ext]) for name in names),
        "core_content_status": core_status,
        "document_content_status": document_content_status,
        "image_status": image_status,
        **image_counts(image_records, reported_missing_media),
        "overall_status": overall_status,
        "auxiliary_failed_entries": len(failed_entries) - len(content_failed) - len(media_failed),
        "content_failed_entries": len(content_failed),
        "media_failed_entries": len(media_failed),
        "metadata": metadata,
        "entries": [_json_entry(item) for item in entries],
        "orphan_spans": orphan_spans,
        "content_parts": content_parts,
        "images": image_records,
        "fragments": fragment_records,
    }
    (output_directory / "recovery_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    if dump_entries:
        dump = [{key: value for key, value in item.items() if not key.startswith("_")} for item in entries]
        (output_directory / "entries.json").write_text(
            json.dumps(dump, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
    safe_print(f"[INFO] Recovery status: {overall_status}; manifest: {output_directory / 'recovery_manifest.json'}")
    return RecoveryResult(entries, inferred_ext, output_directory, overall_status)
