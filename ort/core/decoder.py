"""Bounded decoding and content validation for damaged OOXML local-file entries.

Two axes are kept separate (evaluation_definitions_v2):
- ``decode_status``: whether the DEFLATE stream ended (``complete``), was STORED,
  produced only a prefix, or produced nothing.
- ``content_validation_status``: whether the bytes are evidenced to be the claimed
  content (CRC-32 against a trusted reference, or a structural check).
"""

from __future__ import annotations

import hashlib
import re
import xml.etree.ElementTree as ET
import zlib


MAX_DECODED_BYTES = 16 * 1024 * 1024
_PROBE_CHUNK = 1 << 16
_XML_WHITESPACE = b" \t\r\n"
OK_VALIDATION = ("crc_verified", "structure_verified")
CONTENT_YIELDING_VALIDATION = OK_VALIDATION + ("prefix_validated", "crc_mismatch", "unverified")

# evaluation_definitions_v2_1 §2: namespace URI + complete start tag of a part-appropriate element.
_NAMESPACE_EVIDENCE = re.compile(rb"schemas\.openxmlformats\.org|schemas\.microsoft\.com/office")
_ANY_PREFIX = rb"(?:[A-Za-z_][\w.-]*:)?"
_EXPECTED_ELEMENTS = (
    (re.compile(r"^word/document\.xml$"), (b"w:document", b"w:body", b"w:p", b"w:r", b"w:t")),
    (re.compile(r"^word/header\d+\.xml$"), (b"w:hdr", b"w:p")),
    (re.compile(r"^word/footer\d+\.xml$"), (b"w:ftr", b"w:p")),
    (re.compile(r"^word/footnotes\.xml$"), (b"w:footnotes", b"w:footnote")),
    (re.compile(r"^word/endnotes\.xml$"), (b"w:endnotes", b"w:endnote")),
    (re.compile(r"^word/comments\.xml$"), (b"w:comments", b"w:comment")),
    (re.compile(r"^word/glossary/document\.xml$"), (b"w:glossaryDocument", b"w:docParts", b"w:docPart", b"w:p")),
    (re.compile(r"^xl/worksheets/sheet\d+\.xml$"), (_ANY_PREFIX + b"worksheet", _ANY_PREFIX + b"sheetData",
                                                    _ANY_PREFIX + b"row", _ANY_PREFIX + b"c")),
    (re.compile(r"^xl/sharedStrings\.xml$"), (_ANY_PREFIX + b"sst", _ANY_PREFIX + b"si", _ANY_PREFIX + b"t")),
    (re.compile(r"^ppt/slides/slide\d+\.xml$"), (b"p:sld", b"p:cSld", b"p:sp", b"a:p")),
    (re.compile(r"^ppt/notesSlides/notesSlide\d+\.xml$"), (b"p:notes", b"p:cSld", b"p:sp", b"a:p")),
    (re.compile(r"^\[Content_Types\]\.xml$"), (_ANY_PREFIX + b"Types", _ANY_PREFIX + b"Default", _ANY_PREFIX + b"Override")),
    (re.compile(r"\.rels$"), (_ANY_PREFIX + b"Relationships", _ANY_PREFIX + b"Relationship")),
)
_GENERIC_ELEMENT = (rb"[A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?",)


def _complete_start_tag(names):
    """A start tag closed by '>' (or '/>'), e.g. <w:p> or <w:document xmlns:w="...">."""
    alternatives = b"|".join(name if b"(" in name or b"[" in name else re.escape(name) for name in names)
    return re.compile(rb"<(?:" + alternatives + rb")(?:\s[^<>]*)?/?>")


_GENERIC_TAG = _complete_start_tag(_GENERIC_ELEMENT)
_EXPECTED_TAGS = tuple((pattern, _complete_start_tag(names)) for pattern, names in _EXPECTED_ELEMENTS)


def valid_utf8_prefix(raw):
    """(longest strictly valid UTF-8 prefix, offset of the first decode error or None).

    Bytes are never skipped inside the prefix: it ends at the start of the first error
    (evaluation_definitions_v2_1 §1)."""
    try:
        raw.decode("utf-8")
        return raw, None
    except UnicodeDecodeError as exc:
        return raw[:exc.start], exc.start


def _count_replacements(raw):
    """Decode with U+FFFD replacement and count the replaced invalid sequences."""
    count = 0
    position = 0
    pieces = []
    while position < len(raw):
        try:
            pieces.append(raw[position:].decode("utf-8"))
            break
        except UnicodeDecodeError as exc:
            pieces.append(raw[position:position + exc.start].decode("utf-8"))
            pieces.append("�")
            count += 1
            position += max(exc.end, exc.start + 1)
    return "".join(pieces), count


def is_xml_part(name):
    return name.endswith((".xml", ".rels")) or name == "[Content_Types].xml"


def is_media_part(name):
    return "/media/" in name and not name.endswith("/")


_IMAGE_EXTENSIONS = {
    "png", "jpg", "jpeg", "jpe", "jfif", "gif", "bmp", "dib", "tif", "tiff", "webp", "ico",
    "svg", "emf", "wmf", "emz", "wmz", "wdp", "jxr", "hdp",
}


def is_stored_image_part(name):
    """Media part whose extension names an image format (audio/video are excluded)."""
    if not is_media_part(name) or "." not in name.rsplit("/", 1)[-1]:
        return False
    return name.rsplit(".", 1)[-1].lower() in _IMAGE_EXTENSIONS


def recover_deflate_prefix(data, max_output=MAX_DECODED_BYTES):
    """Recover only the output produced before the first raw-DEFLATE error."""
    decoder = zlib.decompressobj(-zlib.MAX_WBITS)
    chunks = []
    total = 0
    for byte in data:
        try:
            chunk = decoder.decompress(bytes([byte]), max_output + 1 - total)
        except zlib.error:
            break
        chunks.append(chunk)
        total += len(chunk)
        if total > max_output:
            return b""
        if decoder.eof:
            break
    return b"".join(chunks)


def inflate_probe(data):
    """Stream-inflate without keeping output; report end, consumed input, size, and CRC-32.

    Used by the scanner to locate a stream's end cheaply. ``status`` is ``complete``
    (end-of-stream reached), ``truncated`` (input ended first), or ``error``.
    """
    decoder = zlib.decompressobj(-zlib.MAX_WBITS)
    crc = 0
    size = 0
    position = 0
    try:
        while True:
            if decoder.unconsumed_tail:
                out = decoder.decompress(decoder.unconsumed_tail, _PROBE_CHUNK)
            else:
                if position >= len(data):
                    break
                chunk = data[position : position + _PROBE_CHUNK]
                position += len(chunk)
                out = decoder.decompress(chunk, _PROBE_CHUNK)
            crc = zlib.crc32(out, crc)
            size += len(out)
            if decoder.eof:
                break
    except zlib.error:
        return {"status": "error", "consumed": 0, "size": size, "crc": None}
    consumed = position - len(decoder.unused_data) - len(decoder.unconsumed_tail)
    status = "complete" if decoder.eof else "truncated"
    return {"status": status, "consumed": consumed, "size": size, "crc": crc & 0xFFFFFFFF}


def decompress_with_status(data, max_output=MAX_DECODED_BYTES):
    """Return decoded bytes, status, error, consumed input, and unused input."""
    decoder = zlib.decompressobj(-zlib.MAX_WBITS)
    try:
        raw = decoder.decompress(data, max_output + 1)
        if len(raw) > max_output or decoder.unconsumed_tail:
            return b"", "failed", "decoded output limit exceeded", 0, b""
        consumed = len(data) - len(decoder.unused_data)
        status = "complete" if decoder.eof else ("partial_prefix" if raw else "failed")
        return raw, status, None, consumed, decoder.unused_data
    except zlib.error as exc:
        fallback = zlib.decompressobj(-zlib.MAX_WBITS)
        chunks = []
        total = 0
        consumed = 0
        for index, byte in enumerate(data):
            try:
                chunk = fallback.decompress(bytes([byte]), max_output + 1 - total)
            except zlib.error:
                consumed = index
                break
            chunks.append(chunk)
            total += len(chunk)
            consumed = index + 1
            if total > max_output:
                return b"", "failed", "decoded output limit exceeded", consumed, b""
            if fallback.eof:
                break
        raw = b"".join(chunks)
        return raw, ("partial_prefix" if raw else "failed"), str(exc), consumed, b""


def validate_xml_prefix(raw, part_name=""):
    """OOXML XML evidence on the strict UTF-8 prefix (evaluation_definitions_v2_1 §2).

    Rejects short zlib artifacts, text without an OOXML namespace URI, and fragments such as
    '<w:document' without a complete start tag."""
    if not raw:
        return False, "empty"
    prefix, error_offset = valid_utf8_prefix(raw)
    body = prefix[3:] if prefix.startswith(b"\xef\xbb\xbf") else prefix
    body = body.lstrip(_XML_WHITESPACE)
    if not body and error_offset is not None:
        return False, "invalid_utf8"
    if not body.startswith(b"<"):
        return False, "not_xml_start"
    if not _NAMESPACE_EVIDENCE.search(body):
        return False, "no_ooxml_namespace"
    tag = next((compiled for pattern, compiled in _EXPECTED_TAGS if pattern.search(part_name or "")), _GENERIC_TAG)
    if not tag.search(body):
        return False, "no_complete_expected_start_tag"
    return True, "ooxml_evidence"


def _well_formed(raw):
    try:
        ET.fromstring(raw)
        return True
    except ET.ParseError:
        return False


def assess_content(name, raw, decode_status, crc_refs, size_refs):
    """Classify decoded bytes on the content axis (see evaluation_definitions_v2 §1)."""
    # Imported lazily: utils imports Pillow and must not be required by the decoder tests.
    from .utils import detect_image_format, validate_image_bytes

    if not raw:
        return "rejected", "no decoded bytes"
    media = is_media_part(name)
    xml = is_xml_part(name)
    size_consistent = not size_refs or len(raw) in size_refs

    if decode_status in ("complete", "stored"):
        crc = zlib.crc32(raw) & 0xFFFFFFFF
        if crc_refs:
            if crc in crc_refs:
                return "crc_verified", "crc32 matches trusted reference"
            if media and detect_image_format(raw):
                return "crc_mismatch", "crc32 differs; recognized image signature kept"
            if xml and validate_xml_prefix(raw, name)[0]:
                return "crc_mismatch", "crc32 differs; OOXML evidence on valid prefix kept"
            return "rejected", "crc32 differs and no content evidence"
        if xml:
            if _well_formed(raw):
                return "structure_verified", "no crc reference; well-formed XML"
            if validate_xml_prefix(raw, name)[0]:
                return "unverified", "no crc reference; not well-formed, OOXML evidence kept"
            return "rejected", "no crc reference; XML structure invalid"
        if media:
            image_format = detect_image_format(raw)
            if image_format:
                status, _ = validate_image_bytes(raw, image_format)
                if status == "decodable_structurally_complete":
                    return "structure_verified", "no crc reference; image fully decodable"
                return "unverified", "no crc reference; recognized but not structurally complete"
            if size_consistent:
                return "unverified", "no crc reference; size consistent"
            return "rejected", "no crc reference; unrecognized bytes with inconsistent size"
        if size_consistent:
            return "unverified", "no crc reference or validator"
        return "rejected", "no crc reference; inconsistent size"

    if decode_status == "partial_prefix":
        if media:
            if detect_image_format(raw):
                kind = "crc_mismatch" if crc_refs else "unverified"
                return kind, "truncated media with recognized signature"
            return "rejected", "truncated media without recognized signature"
        accepted, reason = validate_xml_prefix(raw, name)
        return ("prefix_validated", reason) if accepted else ("rejected", reason)

    return "rejected", "no decoded bytes"


def decode_local_file_data(entries, max_output=MAX_DECODED_BYTES):
    """Decode entries and attach both status axes.

    Each entry may carry ``crc_refs`` / ``size_refs`` (trusted CRC-32 values and
    uncompressed sizes) gathered by the extractor. Rejected bytes are dropped and only
    their size is kept (``rejected_bytes``). Bytes of other content-yielding parts are kept
    in ``_decoded_bytes`` (XML after prefix trimming) so that they can be written to disk.
    """
    for item in entries:
        name = item.get("local_file_name", "")
        hex_data = item.get("compressed_data_hex", item.get("local_file_data", ""))
        item["compressed_data_hex"] = hex_data
        item.setdefault("decode_status", "not_attempted")
        item.setdefault("decode_error", None)
        item.setdefault("recovered_bytes", 0)
        item.setdefault("compressed_consumed", 0)
        crc_refs = set(item.get("crc_refs") or [])
        size_refs = set(item.get("size_refs") or [])

        if name.endswith("/"):
            item["decode_status"] = "stored" if item.get("compression_method") == 0 else "failed"
            item["content_validation_status"] = "directory"
            continue

        try:
            payload = bytes.fromhex(hex_data) if hex_data else b""
        except ValueError as exc:
            item.update(decode_status="failed", decode_error=str(exc),
                        content_validation_status="rejected", content_validation_reason="invalid payload hex")
            continue

        method = item.get("compression_method")
        if method == 0:
            raw, status, error, consumed = payload, "stored", None, len(payload)
        elif method == 8 and payload:
            raw, status, error, consumed, _ = decompress_with_status(payload, max_output=max_output)
        elif method == 8:
            raw, status, error, consumed = b"", "failed", "empty payload", 0
        else:
            raw, status, error, consumed = b"", "failed", f"unsupported compression method: {method}", 0

        if status == "partial_prefix" and not is_media_part(name):
            accepted, reason = validate_xml_prefix(raw, name)
            item["prefix_validation"] = reason
        validation, reason = assess_content(name, raw, status, crc_refs, size_refs)
        if validation == "rejected":
            if status == "partial_prefix":
                status = "failed"
                error = error or reason
            item["rejected_bytes"] = len(raw)
            # Decision 2026-10-13: rejected bytes of STORED images are kept for diagnosis
            # (written as .bin, never counted as a preserved image). Bytes after a damaged
            # region of an uncompressed image are still original data.
            if method == 0 and raw and is_stored_image_part(name):
                item["rejected_stored_image_hex"] = raw.hex()
            raw = b""

        item["decode_status"] = status if (raw or status != "partial_prefix") else "failed"
        item["decode_error"] = error
        item["content_validation_status"] = validation
        item["content_validation_reason"] = reason
        item["recovered_bytes"] = len(raw)
        item["compressed_consumed"] = consumed

        if not raw:
            continue
        if is_xml_part(name) or (status == "partial_prefix" and not is_media_part(name)):
            # evaluation_definitions_v2_1 §3: the text representation follows the decision above
            # and never overturns it.
            if validation == "prefix_validated":
                prefix, error_offset = valid_utf8_prefix(raw)
                item["utf8_first_error_offset"] = error_offset
                item["utf8_trailing_bytes_removed"] = len(raw) - len(prefix)
                raw = prefix
                item["recovered_bytes"] = len(raw)
                decoded = raw.decode("utf-8")
            elif validation in ("crc_mismatch", "unverified"):
                item["decoded_bytes_sha256"] = hashlib.sha256(raw).hexdigest()
                item["utf8_first_error_offset"] = valid_utf8_prefix(raw)[1]
                decoded, item["utf8_replaced_sequences"] = _count_replacements(raw)
            else:
                try:
                    decoded = raw.decode("utf-8")
                except UnicodeDecodeError as exc:
                    item.update(decode_status="failed", decode_error=f"invalid UTF-8: {exc}",
                                content_validation_status="rejected", content_validation_reason="invalid UTF-8",
                                rejected_bytes=len(raw), recovered_bytes=0)
                    continue
            item["decoded_xml"] = decoded
            item["local_file_data"] = decoded
            item["_decoded_bytes"] = raw
        elif is_media_part(name):
            item["decoded_media_hex"] = raw.hex()
        else:
            item["_decoded_bytes"] = raw

    return entries
