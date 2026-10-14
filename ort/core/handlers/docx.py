"""Independent extraction of DOCX text-bearing parts."""

from __future__ import annotations

import re
from pathlib import Path

from .common import extract_xml_text_for_docx, save_text_to_docx
from ..utils import safe_print, source_line


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


def process_extracted_docx_data(entries, output_directory):
    """Write text/<part>.txt per text part, text/document_content.txt, and text/body_text.docx.

    Every file starts with a provenance line naming the source part and its validation status.
    Returns one record per written part."""
    text_dir = Path(output_directory) / "text"
    recovered = []
    combined = []
    body = None
    for item in entries:
        name = item.get("local_file_name", "")
        if not is_docx_text_part(name):
            continue
        xml = item.get("decoded_xml", "")
        if not xml:
            continue
        text = extract_xml_text_for_docx(xml)
        if not text.strip():
            continue
        text_dir.mkdir(parents=True, exist_ok=True)
        header = source_line(item)
        target = text_dir / _part_filename(name)
        target.write_text(f"{header}\n{text}", encoding="utf-8")
        recovered.append(
            {
                "name": name,
                "text_codepoints": len(text),
                "non_whitespace_codepoints": sum(not char.isspace() for char in text),
                "output": target.relative_to(output_directory).as_posix(),
            }
        )
        combined.append(f"{header}\n{text}")
        if name == "word/document.xml" and body is None:
            body = (header, text)

    if combined:
        (text_dir / "document_content.txt").write_text("\n\n".join(combined), encoding="utf-8")
    if body:
        save_text_to_docx(body, str(text_dir / "body_text.docx"))
    if not recovered:
        safe_print("[DOCX] No recoverable document text found.")
    return recovered
