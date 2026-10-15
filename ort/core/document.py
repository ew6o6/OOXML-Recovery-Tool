"""recovered_document.md: one readable document per input, built from the recovered parts.

Sections:
  Document information  input hash, tool, statuses, document properties
  Recovered content     reading order; every section names the part and its validation status;
                        images appear where they are referenced ("[Image N]") and link to the files in img/
  Labels                an index by validation status: part, byte location, size, and the section above

Reading order: DOCX body, headers, footers, footnotes, endnotes, comments; XLSX sheets in workbook order,
with images at their anchor cells; PPTX slides in presentation order, with pictures and speaker notes.
The document is a reading aid. The exact recovered text and parts remain in text/ and parts/.
"""

import posixpath
import re
from pathlib import Path

from .handlers.common import column_letters, parse_shared_strings, parse_sheet_cells, resolve_cell
from .utils import STATUS_NOTES, local_name, parse_xml

FILENAME = "recovered_document.md"
LABEL_ORDER = ["crc_verified", "structure_verified", "prefix_validated", "crc_mismatch", "unverified", "rejected"]
NOTES = {**STATUS_NOTES, "rejected": "no evidence of being the claimed content; not written"}
WARNINGS = {
    "prefix_validated": "**Damage point:** this part is recovered up to here; the text just above may be non-original.",
    "crc_mismatch": "**Warning:** this part differs from the original (CRC-32 mismatch); where it differs is unknown.",
    "unverified": "**Warning:** whether this part is original cannot be confirmed.",
    "structure_verified": "**Note:** complete and well-formed, but there is no CRC-32 record to confirm it is original.",
}
CORE_KEYS = ["title", "subject", "creator", "lastModifiedBy", "created", "modified", "lastPrinted", "revision"]
APP_KEYS = ["Application", "AppVersion", "Company", "TotalTime", "Pages", "Words", "Slides"]
DOCX_GROUPS = [("Header", r"^word/header\d*\.xml$"), ("Footer", r"^word/footer\d*\.xml$"),
               ("Footnotes", r"^word/footnotes\.xml$"), ("Endnotes", r"^word/endnotes\.xml$"),
               ("Comments", r"^word/comments\.xml$")]
DOCX_CONTAINERS = {"sdt", "sdtContent", "body", "customXml", "hdr", "ftr", "footnote", "endnote", "comment"}


def _attribute(element, name):
    """Attribute by local name (namespace or an undeclared prefix removed)."""
    for key, value in element.attrib.items():
        if key.rsplit("}", 1)[-1].rsplit(":", 1)[-1] == name:
            return value
    return None


def _rels_name(part):
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def _number(part):
    match = re.search(r"(\d+)\.xml$", part)
    return int(match.group(1)) if match else 0


def _link(path):
    return f"<{path}>"


class _Builder:
    def __init__(self, entries, manifest):
        self.manifest = manifest
        self.entries = {}
        for item in entries:
            self.entries.setdefault(item.get("local_file_name", ""), item)
        self.images = {record["name"]: record for record in manifest.get("images", [])}
        self.numbers, self.embedded = {}, set()
        self.headings, self.chars = {}, {}
        self.lines = []

    # ------------------------------------------------------------------ parts
    def status(self, part):
        item = self.entries.get(part)
        return item.get("content_validation_status") if item else None

    def xml(self, part):
        item = self.entries.get(part)
        return parse_xml(item.get("decoded_xml")) if item and item.get("decoded_xml") else None

    def relationships(self, part):
        """Relationship id -> (target part, type) from the recovered .rels part of a source part."""
        root = self.xml(_rels_name(part))
        result = {}
        for element in root.iter() if root is not None else ():
            if local_name(element) != "Relationship" or element.get("TargetMode") == "External":
                continue
            target = element.get("Target") or ""
            target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(
                posixpath.join(posixpath.dirname(part), target))
            result[element.get("Id")] = (target, element.get("Type") or "")
        return result

    # ------------------------------------------------------------------ images
    def marker(self, media):
        if media not in self.numbers:
            self.numbers[media] = len(self.numbers) + 1
        return f"[Image {self.numbers[media]}]"

    def image_lines(self, media):
        number = self.numbers[media]
        if media in self.embedded:
            return [f"*Image {number} (shown above).*"]
        self.embedded.add(media)
        record = self.images.get(media)
        if record is None:
            status = self.status(media)
            return [f"> Image {number}: `{media}` was not recovered" + (f" ({status})." if status else
                                                                         " (not found in the file).")]
        status = record.get("content_validation_status")
        if record.get("output"):
            check = record.get("format_validation")
            extra = "" if check == "decodable_complete" else f"; format check: {check}"
            return [f"![Image {number}]({_link(record['output'])})",
                    f"*Image {number} · {status}{extra} · `{media}`*"]
        if record.get("uniform_fill_byte"):
            detail = f"its data area is filled with {record['uniform_fill_byte']} ({record.get('rejected_bytes', 0):,} bytes)"
        elif record.get("diagnostic_output"):
            detail = f"not an image; bytes kept for examination in `{record['diagnostic_output']}`"
        else:
            detail = "no image was recovered"
        return [f"> Image {number}: `{media}` — {status or record.get('status')}; {detail}."]

    # ------------------------------------------------------------------ text
    def runs(self, node, rels, pending):
        """Text of a node in document order with image markers; mc:Fallback copies are skipped."""
        pieces = []
        for child in node:
            name = local_name(child)
            if name == "Fallback":
                continue
            if name == "t" and child.text:
                pieces.append(child.text)
            elif name == "tab":
                pieces.append("\t")
            elif name in ("br", "cr"):
                pieces.append("\n")
            elif name in ("blip", "imagedata"):
                relationship = _attribute(child, "embed" if name == "blip" else "id")
                if relationship in rels:
                    media = rels[relationship][0]
                    pieces.append(self.marker(media))
                    pending.append(media)
            pieces.append(self.runs(child, rels, pending))
        return "".join(pieces)

    def with_images(self, text, pending):
        lines = [text] if text.strip() else []
        for media in pending:
            lines.extend(self.image_lines(media))
        pending.clear()
        return lines

    def section(self, heading, part, body):
        status = self.status(part)
        self.lines.append(f"### {heading} (`{part}`) — {status or 'not found'}")
        self.headings[part] = heading
        if status in (None, "rejected") or not body:
            self.lines += ["> No recovered content." if status != "rejected" else "> Rejected; not written.", ""]
            return
        if status in ("crc_mismatch", "unverified", "structure_verified"):
            self.lines += [f"> {WARNINGS[status]}", ""]
        # Markdown joins adjacent lines: paragraphs are separated by a blank line, list items are not.
        for index, line in enumerate(body):
            if index and not (line.startswith("- ") and body[index - 1].startswith("- ")):
                self.lines.append("")
            self.lines.append(line)
        if status == "prefix_validated":
            self.lines += ["", f"> {WARNINGS[status]}"]
        self.lines.append("")
        self.chars[part] = sum(len(line) for line in body if not line.startswith(("![", "*Image", "> Image")))

    # ------------------------------------------------------------------ formats
    def docx(self):
        def blocks(container, rels):
            lines = []
            for child in container:
                name = local_name(child)
                pending = []
                if name == "p":
                    lines += self.with_images(self.runs(child, rels, pending), pending)
                elif name == "tbl":
                    for row in (element for element in child.iter() if local_name(element) == "tr"):
                        cells = [self.runs(cell, rels, pending).replace("\n", " ").strip()
                                 for cell in row if local_name(cell) == "tc"]
                        lines += self.with_images("| " + " | ".join(cells) + " |", pending)
                elif name in DOCX_CONTAINERS:
                    lines += blocks(child, rels)
            return lines

        root = self.xml("word/document.xml")
        body = next((e for e in root.iter() if local_name(e) == "body"), root) if root is not None else None
        self.section("Body", "word/document.xml",
                     blocks(body, self.relationships("word/document.xml")) if body is not None else [])
        for heading, pattern in DOCX_GROUPS:
            for part in sorted((name for name in self.entries if re.match(pattern, name)), key=_number):
                root = self.xml(part)
                lines = blocks(root, self.relationships(part)) if root is not None else []
                if lines or self.status(part) == "rejected":
                    self.section(heading, part, lines)

    def xlsx(self):
        workbook_rels = self.relationships("xl/workbook.xml")
        workbook = self.xml("xl/workbook.xml")
        sheets = [(sheet.get("name"), workbook_rels[_attribute(sheet, "id")][0])
                  for sheet in (workbook.iter() if workbook is not None else ())
                  if local_name(sheet) == "sheet" and _attribute(sheet, "id") in workbook_rels]
        known = {part for _, part in sheets}
        sheets += [(None, part) for part in sorted(self.entries, key=_number)
                   if re.match(r"^xl/worksheets/sheet\d+\.xml$", part) and part not in known]
        shared_part = "xl/sharedStrings.xml"
        shared_item = self.entries.get(shared_part)
        shared = parse_shared_strings(shared_item["decoded_xml"]) if shared_item and shared_item.get("decoded_xml") else []
        if shared_part in self.entries:
            self.lines += [f"> Shared strings (`{shared_part}`): {self.status(shared_part)}. "
                           "String cell values are taken from this part.", ""]
        for index, (name, part) in enumerate(sheets, 1):
            anchors = {}
            for target, kind in self.relationships(part).values():
                if not kind.endswith("/drawing"):
                    continue
                drawing, drawing_rels = self.xml(target), self.relationships(target)
                for anchor in drawing.iter() if drawing is not None else ():
                    if local_name(anchor) not in ("twoCellAnchor", "oneCellAnchor"):
                        continue
                    start = next((e for e in anchor if local_name(e) == "from"), None)
                    cells = {local_name(e): e.text for e in start} if start is not None else {}
                    row = int(cells.get("row") or 0) + 1
                    column = int(cells.get("col") or 0) + 1
                    for blip in (e for e in anchor.iter() if local_name(e) == "blip"):
                        if _attribute(blip, "embed") in drawing_rels:
                            anchors.setdefault(row, []).append(
                                (f"{column_letters(column)}{row}", drawing_rels[_attribute(blip, "embed")][0]))
            item = self.entries.get(part)
            rows = {}
            for cell in parse_sheet_cells(item["decoded_xml"]) if item and item.get("decoded_xml") else []:
                value, _ = resolve_cell(cell, shared)
                if value != "" or cell["formula"]:
                    text = f"{cell['ref']}={value}".replace("\n", " / ")
                    rows.setdefault(cell["row"], []).append(text + (f" (formula ={cell['formula']})" if cell["formula"] else ""))
            lines = []
            for row in sorted(set(rows) | set(anchors)):
                if row in rows:
                    lines.append(f"- Row {row}: " + " | ".join(rows[row]))
                for ref, media in anchors.get(row, []):
                    lines.append(f"- At {ref}: {self.marker(media)}")
                    lines += self.image_lines(media)
            self.section(f"Sheet {index}" + (f" '{name}'" if name else " (name unknown)"), part, lines)

    def pptx(self):
        presentation_rels = self.relationships("ppt/presentation.xml")
        presentation = self.xml("ppt/presentation.xml")
        slides = [presentation_rels[_attribute(e, "id")][0]
                  for e in (presentation.iter() if presentation is not None else ())
                  if local_name(e) == "sldId" and _attribute(e, "id") in presentation_rels]
        slides += [part for part in sorted(self.entries, key=_number)
                   if re.match(r"^ppt/slides/slide\d+\.xml$", part) and part not in slides]

        def shapes(container, rels, lines):
            for child in container:
                name = local_name(child)
                pending = []
                if name == "sp":
                    paragraphs = [self.runs(p, rels, pending) for p in child.iter() if local_name(p) == "p"]
                    lines += self.with_images("\n".join(p for p in paragraphs if p.strip()), pending)
                elif name == "pic":
                    self.runs(child, rels, pending)
                    for media in pending:
                        lines.append(self.marker(media))
                        lines += self.image_lines(media)
                elif name == "graphicFrame":
                    for row in (e for e in child.iter() if local_name(e) == "tr"):
                        cells = [self.runs(cell, rels, pending).replace("\n", " ").strip()
                                 for cell in row if local_name(cell) == "tc"]
                        lines += self.with_images("| " + " | ".join(cells) + " |", pending)
                elif name in ("grpSp", "spTree"):
                    shapes(child, rels, lines)
                elif name == "AlternateContent":
                    choice = next((e for e in child if local_name(e) == "Choice"), None)
                    if choice is not None:
                        shapes(choice, rels, lines)

        for index, part in enumerate(slides, 1):
            rels = self.relationships(part)
            root = self.xml(part)
            lines = []
            tree = next((e for e in root.iter() if local_name(e) == "spTree"), None) if root is not None else None
            if tree is not None:
                shapes(tree, rels, lines)
            self.section(f"Slide {index}", part, lines)
            for target, kind in rels.values():
                if not kind.endswith("/notesSlide"):
                    continue
                notes = self.xml(target)
                texts = []
                for shape in (e for e in notes.iter() if local_name(e) == "sp") if notes is not None else ():
                    placeholder = next((e for e in shape.iter() if local_name(e) == "ph"), None)
                    if placeholder is not None and placeholder.get("type") == "body":
                        texts += [self.runs(p, {}, []) for p in shape.iter() if local_name(p) == "p"]
                texts = [text for text in texts if text.strip()]
                if texts or self.status(target) == "rejected":
                    self.section(f"Slide {index} notes", target, texts)

    # ------------------------------------------------------------------ document
    def build(self):
        m = self.manifest
        self.lines += [f"# Recovered document: {Path(m.get('input', '')).name}", "", "## Document information", "",
                       f"- Input SHA-256: `{m.get('input_sha256')}`",
                       f"- Tool: ORT {m.get('tool_version')} (code SHA-256 `{m.get('tool_code_sha256')}`)",
                       f"- Type: {m.get('inferred_type')} · overall: {m.get('overall_status')} · "
                       f"core content: {m.get('core_content_status')} · images: {m.get('image_status')}",
                       "- Every section below names its part and validation status; the index by status is under "
                       "\"Labels\". Exact text and parts are in `text/` and `parts/`."]
        for part, keys in (("docProps/core.xml", CORE_KEYS), ("docProps/app.xml", APP_KEYS)):
            record = (m.get("metadata") or {}).get(part)
            if record:
                values = [f"{key}: {record['values'][key]}" for key in keys if record["values"].get(key)]
                if values:
                    self.lines.append(f"- Properties (`{part}`, {record.get('content_validation_status')}): "
                                      + "; ".join(values))
        self.lines += ["", "## Recovered content", ""]
        {".docx": self.docx, ".xlsx": self.xlsx, ".pptx": self.pptx}[m["inferred_type"]]()
        unplaced = [name for name in self.images if name not in self.numbers]
        if unplaced:
            self.lines += ["### Images outside the content or without a known position", ""]
            for media in unplaced:
                self.marker(media)
                users = [part for part in self.entries if not part.endswith(".rels")
                         and any(target == media for target, _ in self.relationships(part).values())]
                where = ("referenced by " + ", ".join(f"`{u}`" for u in users)) if users else \
                    "no recovered part references it"
                self.lines += [f"- Image {self.numbers[media]}: {where}"] + self.image_lines(media) + [""]
        self.labels()
        return "\n".join(self.lines) + "\n"

    def labels(self):
        self.lines += ["## Labels", ""]
        for status in LABEL_ORDER:
            parts = [part for part, item in self.entries.items()
                     if item.get("content_validation_status") == status and part not in self.images]
            images = [media for media, record in self.images.items() if record.get("content_validation_status") == status]
            if not parts and not images:
                continue
            self.lines += [f"### {status}", f"> {NOTES[status]}", ""]
            for part in parts:
                if part not in self.chars:
                    continue
                item = self.entries[part]
                span = item.get("decoded_span") or item.get("declared_span")
                where = f"header at byte {item.get('source_offset', 0):,}" + (
                    f", data {span[0]:,}–{span[1]:,}" if span else "")
                self.lines.append(f"- `{part}`: \"{self.headings[part]}\" ({where}; {self.chars[part]:,} characters)")
            for media in images:
                record = self.images[media]
                label = f"Image {self.numbers[media]}" if media in self.numbers else "Image"
                target = f" → `{record['output']}`" if record.get("output") else ""
                self.lines.append(f"- {label}: `{media}` (header at byte {record.get('source_offset') or 0:,}){target}")
            others = [part for part in parts if part not in self.chars]
            if others:
                self.lines.append(f"- Other parts ({len(others)}): " + ", ".join(f"`{p}`" for p in others))
            self.lines.append("")


def write_recovered_document(entries, manifest, output_directory):
    """Write recovered_document.md into the output folder and return its name."""
    text = _Builder(entries, manifest).build()
    (Path(output_directory) / FILENAME).write_text(text, encoding="utf-8")
    return FILENAME
