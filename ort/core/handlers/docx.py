"""Independent extraction of DOCX text-bearing parts."""

from __future__ import annotations

import os
import re
from pathlib import Path

from .common import extract_xml_text_for_docx, parse_rels_file, parse_styles_xml, save_text_to_docx
from ..utils import safe_output_stem, safe_print


TEXT_PARTS = (
    re.compile(r"^word/document\.xml$"),
    re.compile(r"^word/header\d+\.xml$"),
    re.compile(r"^word/footer\d+\.xml$"),
    re.compile(r"^word/(?:footnotes|endnotes|comments)\.xml$"),
    re.compile(r"^word/glossary/document\.xml$"),
)


def is_docx_text_part(name):
    return any(pattern.match(name) for pattern in TEXT_PARTS)


def _part_filename(name):
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", name).strip("_") + ".txt"


def process_extracted_docx_data(entries, file_path, output_directory=None):
    out_dir = Path(output_directory or os.path.join(os.path.dirname(file_path), f"output_{safe_output_stem(file_path)}"))
    out_dir.mkdir(parents=True, exist_ok=True)

    styles_xml = next(
        (item.get("decoded_xml", "") for item in entries if item.get("local_file_name") == "word/styles.xml"),
        "",
    )
    rels_xml = next(
        (
            item.get("decoded_xml", "")
            for item in entries
            if item.get("local_file_name") == "word/_rels/document.xml.rels"
        ),
        "",
    )
    styles = parse_styles_xml(styles_xml) if styles_xml else {}

    text_dir = out_dir / "text_parts"
    recovered = []
    combined = []
    body_text = ""
    for item in entries:
        name = item.get("local_file_name", "")
        if not is_docx_text_part(name):
            continue
        xml = item.get("decoded_xml", "")
        if not xml:
            continue
        text = extract_xml_text_for_docx(xml, styles)
        if not text.strip():
            continue
        text_dir.mkdir(parents=True, exist_ok=True)
        target = text_dir / _part_filename(name)
        target.write_text(text, encoding="utf-8")
        recovered.append(
            {
                "name": name,
                "decode_status": item.get("decode_status"),
                "content_validation_status": item.get("content_validation_status"),
                "text_codepoints": len(text),
                "non_whitespace_codepoints": sum(not char.isspace() for char in text),
                "output": target.relative_to(out_dir).as_posix(),
            }
        )
        combined.append(f"[{name}]\n{text}")
        if name == "word/document.xml":
            body_text = text

    if combined:
        (out_dir / "document_content.txt").write_text("\n\n".join(combined), encoding="utf-8")
    if body_text.strip():
        save_text_to_docx(body_text, str(out_dir / "output_damaged.docx"))
    if rels_xml:
        (out_dir / "relationships.txt").write_text("\n".join(parse_rels_file(rels_xml)), encoding="utf-8")

    if recovered:
        safe_print(f"\n[DOCX output path] {out_dir}")
    else:
        safe_print("[DOCX] No recoverable document text found.")
    return {
        "recovered_parts": recovered,
        "body_text_recovered": bool(body_text.strip()),
        "any_text_recovered": bool(recovered),
    }
