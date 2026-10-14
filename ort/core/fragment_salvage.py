"""Opt-in, marker-based recovery of independently decodable DEFLATE suffixes."""

from __future__ import annotations

import hashlib
import html
import re
import zlib
from pathlib import Path


MARKER = b"\x00\x00\xff\xff"
MAX_OUTPUT = 16 * 1024 * 1024
MAX_CANDIDATES_PER_SOURCE = 64
_TEXT_TAGS = {
    ".docx": re.compile(rb"<w:t(?:\s[^>]*)?>([^<]*)</w:t>"),
    ".xlsx": re.compile(rb"<(?:t|v)(?:\s[^>]*)?>([^<]*)</(?:t|v)>"),
    ".pptx": re.compile(rb"<a:t(?:\s[^>]*)?>([^<]*)</a:t>"),
}
_CLOSING_EVIDENCE = re.compile(
    rb"</(?:w:(?:document|body|p|tbl|hdr|ftr)|a:(?:p|r|txBody)|worksheet|sst)>"
)


def _trim_split_utf8(raw):
    body = raw
    removed = 0
    while removed < 3 and body and body[0] & 0xC0 == 0x80:
        body = body[1:]
        removed += 1
    return body, removed


def validate_fragment(raw, extension):
    body, leading_bytes_removed = _trim_split_utf8(raw)
    try:
        body.decode("utf-8")
    except UnicodeDecodeError:
        return False, "invalid_utf8", leading_bytes_removed, []
    pattern = _TEXT_TAGS.get(extension)
    text_nodes = pattern.findall(body) if pattern else []
    if text_nodes:
        return True, "text_tag", leading_bytes_removed, text_nodes
    if _CLOSING_EVIDENCE.search(body):
        return True, "closing_ooxml_tag", leading_bytes_removed, []
    return False, "no_ooxml_evidence", leading_bytes_removed, []


def recover_deflate_fragments(
    data,
    *,
    extension,
    source_start=0,
    max_output=MAX_OUTPUT,
    max_candidates=MAX_CANDIDATES_PER_SOURCE,
):
    """Yield independently decodable suffix candidates after full-flush markers."""
    position = 0
    tried = 0
    while tried < max_candidates:
        marker_offset = data.find(MARKER, position)
        if marker_offset < 0:
            break
        position = marker_offset + 1
        tried += 1
        compressed_offset = marker_offset + len(MARKER)
        decoder = zlib.decompressobj(-zlib.MAX_WBITS)
        try:
            raw = decoder.decompress(data[compressed_offset:], max_output + 1)
        except zlib.error:
            continue
        if not decoder.eof or not raw or len(raw) > max_output or decoder.unconsumed_tail:
            continue
        accepted, evidence, removed, text_nodes = validate_fragment(raw, extension)
        if not accepted:
            continue
        normalized_raw = raw[removed:]
        text = "".join(
            html.unescape(node.decode("utf-8", "replace")) for node in text_nodes
        )
        consumed = len(data[compressed_offset:]) - len(decoder.unused_data)
        yield {
            "absolute_compressed_offset": source_start + compressed_offset,
            "compressed_end": source_start + compressed_offset + consumed,
            "decoded": normalized_raw,
            "decoded_sha256": hashlib.sha256(normalized_raw).hexdigest(),
            "decoded_bytes": len(normalized_raw),
            "validation_evidence": evidence,
            "leading_split_utf8_bytes_removed": removed,
            "text": text,
            "text_codepoints": len(text),
            "non_whitespace_codepoints": sum(not char.isspace() for char in text),
        }


def _is_core_part(name, extension):
    if extension == ".docx":
        return name == "word/document.xml"
    if extension == ".xlsx":
        return bool(re.match(r"^xl/(?:worksheets/sheet\d+|sharedStrings)\.xml$", name or ""))
    if extension == ".pptx":
        return bool(re.match(r"^ppt/slides/slide\d+\.xml$", name or ""))
    return False


def _deduplicate(candidates, extension):
    candidates = sorted(
        candidates,
        key=lambda item: (
            0
            if item.get("source_category") == "failed_payload"
            and _is_core_part(item.get("source_part"), extension)
            else 1
            if item.get("source_category") == "orphan_span"
            else 2
        ),
    )
    exact = {}
    for candidate in candidates:
        key = (candidate["absolute_compressed_offset"], candidate["decoded_sha256"])
        exact.setdefault(key, candidate)
    ordered = sorted(exact.values(), key=lambda item: item["decoded_bytes"], reverse=True)
    kept = []
    for candidate in ordered:
        if any(
            larger["decoded"].endswith(candidate["decoded"])
            and larger["compressed_end"] == candidate["compressed_end"]
            for larger in kept
        ):
            continue
        kept.append(candidate)
    return sorted(kept, key=lambda item: item["absolute_compressed_offset"])


def salvage_sources(sources, *, extension, output_directory):
    candidates = []
    for source in sources:
        for candidate in recover_deflate_fragments(
            source["data"], extension=extension, source_start=source["start"]
        ):
            candidate.update(
                source_category=source["category"],
                source_part=source.get("part"),
            )
            candidates.append(candidate)
    kept = _deduplicate(candidates, extension)
    fragment_dir = Path(output_directory) / "fragments"
    records = []
    for index, candidate in enumerate(kept, 1):
        fragment_dir.mkdir(parents=True, exist_ok=True)
        offset = candidate["absolute_compressed_offset"]
        stem = f"fragment_{index:03d}_offset_{offset}"
        raw_path = fragment_dir / f"{stem}.xml.fragment"
        text_path = fragment_dir / f"{stem}.txt"
        raw_path.write_bytes(candidate["decoded"])
        text_path.write_text(candidate["text"], encoding="utf-8")
        record = {key: value for key, value in candidate.items() if key not in {"decoded", "text"}}
        record.update(
            reported_part=candidate.get("source_part") or f"unknown@{offset}",
            raw_output=raw_path.relative_to(output_directory).as_posix(),
            text_output=text_path.relative_to(output_directory).as_posix(),
        )
        records.append(record)
    return records
