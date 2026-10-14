"""PPTX slide and speaker-note text, in document order."""

import re
from pathlib import Path

from ..utils import local_name, parse_xml, safe_print, source_line


SLIDE_PART = re.compile(r"^ppt/(slides/slide|notesSlides/notesSlide)(\d+)\.xml$")
_TITLE_TYPES = {"title", "ctrTitle"}


def paragraph_text(paragraph):
    """Text of one a:p: runs and fields joined without separators, a:br as a line feed."""
    pieces = []
    for element in paragraph.iter():
        name = local_name(element)
        if name == "t":
            pieces.append(element.text or "")
        elif name == "br":
            pieces.append("\n")
    return "".join(pieces)


def _paragraphs(node):
    return [paragraph_text(element) for element in node.iter() if local_name(element) == "p"]


def _shape_label(shape):
    for element in shape.iter():
        if local_name(element) == "ph":
            return "TITLE" if element.get("type") in _TITLE_TYPES else "BODY"
    return "BODY"


def _table_rows(frame):
    rows = []
    for row in frame.iter():
        if local_name(row) != "tr":
            continue
        cells = [" ".join(_paragraphs(cell)) for cell in row if local_name(cell) == "tc"]
        rows.append("\t".join(cells))
    return rows


def _walk(container, blocks):
    """Shapes in document order; groups are entered, alternate content uses its first choice only."""
    for child in container:
        name = local_name(child)
        if name == "sp":
            paragraphs = _paragraphs(child)
            if "".join(paragraphs).strip():
                while paragraphs and not paragraphs[-1].strip():
                    paragraphs.pop()
                blocks.append(f"[{_shape_label(child)}]\n" + "\n".join(paragraphs))
        elif name == "graphicFrame":
            rows = _table_rows(child)
            if "".join(rows).strip():
                blocks.append("[TABLE]\n" + "\n".join(rows))
        elif name in ("grpSp", "spTree"):
            _walk(child, blocks)
        elif name == "AlternateContent":
            choice = next((element for element in child if local_name(element) == "Choice"), None)
            if choice is not None:
                _walk(choice, blocks)


def slide_text(xml):
    root = parse_xml(xml)
    if root is None:
        return ""
    blocks = []
    for tree in root.iter():
        if local_name(tree) == "spTree":
            _walk(tree, blocks)
            break
    return "\n".join(blocks)


def process_extracted_pptx_data(entries, output_directory):
    """Write text/slideN.txt, text/notesSlideN.txt, and text/document_content.txt.

    N is the part number. Each file starts with a provenance line. Returns one record per part."""
    out_dir = Path(output_directory)
    text_dir = out_dir / "text"
    parts = {}
    for item in entries:
        match = SLIDE_PART.match(item.get("local_file_name", ""))
        if match and item.get("decoded_xml") and item["local_file_name"] not in parts:
            kind = 0 if match.group(1).startswith("slides") else 1
            parts[item["local_file_name"]] = ((kind, int(match.group(2))), item)

    records = []
    combined = []
    for name, (_, item) in sorted(parts.items(), key=lambda pair: pair[1][0]):
        text = slide_text(item["decoded_xml"])
        if not text.strip():
            continue
        text_dir.mkdir(parents=True, exist_ok=True)
        header = source_line(item)
        target = text_dir / (Path(name).stem + ".txt")
        target.write_text(f"{header}\n{text}", encoding="utf-8")
        combined.append(f"{header}\n{text}")
        records.append({"name": name, "text_codepoints": len(text),
                        "output": target.relative_to(out_dir).as_posix()})
    if combined:
        (text_dir / "document_content.txt").write_text("\n\n".join(combined), encoding="utf-8")
    if not records:
        safe_print("[PPTX] No recoverable slide text found.")
    return records
