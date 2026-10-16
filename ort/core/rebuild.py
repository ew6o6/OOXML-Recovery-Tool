"""recovered_<name>.<docx|xlsx|pptx>: the recovered content written anew in the input's own format.

The rebuilt file is not a repair of the original package: content is written into a new document
(python-docx, openpyxl, python-pptx), so the original formatting is not reproduced and nothing that failed
validation is restored. Its first page, sheet, or slide states that it was rebuilt by ORT and is not the
original, with the input SHA-256; every part that is not shown to be original is marked where it appears,
and the last page, sheet, or slide indexes parts and images by validation status.

DOCX  reading order of document.py; consecutive table rows form one table; images embedded in place.
XLSX  first sheet "ORT recovery"; every sheet keeps its name and order, values are written to their original
      cells with the original number format, formulas are kept as cell comments (the stored value is shown,
      so that Excel does not recalculate it). Damaged sheets get a coloured tab and an A1 comment; no row is
      inserted, so cell addresses never shift. Images are placed at their anchor cells.
PPTX  slide size and order of the original; shapes, pictures, and tables at their recorded positions;
      speaker notes copied; slides that are not shown to be original get a status band.
Image formats the libraries cannot embed (e.g. EMF) are replaced by a note with the path of the file in img/.
"""

import re
from pathlib import Path

from .document import LABEL_ORDER, NOTES, STATUS_COLOURS, attribute
from .handlers.common import parse_number_formats, parse_sheet_cells
from .utils import local_name

BANNER = ("Rebuilt by ORT from the recovered content. This is NOT the original document: formatting is not "
          "reproduced, and parts marked below may be incomplete or non-original.")
SHORT = {"structure_verified": "Not shown to be original (no CRC-32 record).",
         "prefix_validated": "Damaged: recovered up to the damage point; the end may be non-original.",
         "crc_mismatch": "Altered: differs from the original (CRC-32 mismatch).",
         "unverified": "Not confirmed: may differ from the original.", "rejected": "Rejected: no content written.",
         None: "Not recovered."}
EMBEDDABLE = {"png", "jpeg", "jpg", "gif", "bmp", "tiff", "tif"}
_XML_INVALID = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff￾￿]")


def rebuilt_name(manifest):
    return f"recovered_{Path(manifest['input']).stem}{manifest['inferred_type']}"


def clean(text):
    """Characters that OOXML cannot store become U+FFFD."""
    return _XML_INVALID.sub("�", text)


def info_lines(manifest):
    lines = [BANNER, f"Input: {Path(manifest['input']).name}", f"Input SHA-256: {manifest['input_sha256']}",
             f"Tool: ORT {manifest['tool_version']} (code SHA-256 {manifest['tool_code_sha256']})",
             f"Overall: {manifest['overall_status']} · core content: {manifest['core_content_status']} · "
             f"images: {manifest['image_status']}"]
    for part, record in (manifest.get("metadata") or {}).items():
        values = "; ".join(f"{key}: {value}" for key, value in record.get("values", {}).items() if value)
        if values:
            lines.append(f"Properties ({part}, {record.get('content_validation_status')}): {values}")
    return lines


def label_lines(content):
    lines = []
    for status in LABEL_ORDER:
        rows = content.label_rows(status)
        if rows:
            lines.append(f"[{status}] {NOTES[status]}")
            lines += [f"    {item} ({where})" for item, _, where in rows]
    return lines


def embeddable(record):
    return bool(record and record.get("output") and record.get("format_validation") == "decodable_complete"
                and (record.get("detected_format") or "").lower() in EMBEDDABLE)


# ----------------------------------------------------------------------------- DOCX
def rebuild_docx(content, manifest, out_dir, target):
    from docx import Document
    from docx.enum.text import WD_COLOR_INDEX
    from docx.shared import Cm, Pt, RGBColor

    blocks = content.build()
    doc = Document()
    for line in info_lines(manifest):
        run = doc.add_paragraph().add_run(clean(line))
        run.font.size = Pt(9)
        if line == BANNER:
            run.bold = True
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.add_paragraph()
    table_rows = []

    def flush_table():
        if table_rows:
            table = doc.add_table(rows=len(table_rows), cols=max(len(row) for row in table_rows))
            table.style = "Table Grid"
            for row, cells in zip(table.rows, table_rows):
                for cell, text in zip(row.cells, cells):
                    cell.text = clean(text)
            table_rows.clear()

    for block in blocks:
        kind = block[0]
        if kind != "table_row":
            flush_table()
        if kind == "section":
            _, heading, part, status, _ = block
            run = doc.add_paragraph().add_run(f"── {heading} ({part}) · {status} ──")
            run.font.size = Pt(8.5)
            run.font.color.rgb = RGBColor.from_string(STATUS_COLOURS.get(status, "#000000")[1:])
        elif kind == "heading":
            doc.add_heading(block[1], level=2)
        elif kind == "warning":
            run = doc.add_paragraph().add_run(block[2])
            run.bold = True
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif kind == "note":
            doc.add_paragraph().add_run(block[1]).italic = True
        elif kind == "text":
            doc.add_paragraph(clean(block[1]))
        elif kind == "table_row":
            table_rows.append(block[1])
        elif kind == "image":
            _, number, media, status, source, image_format, problem = block
            record = content.images.get(media)
            if embeddable(record):
                doc.add_picture(str(out_dir / source), width=Cm(12))
            else:
                reason = "does not decode completely" if problem else f"format {image_format or 'unknown'} cannot be embedded"
                doc.add_paragraph(f"[Image {number}: {reason}; file {source}]")
            caption = doc.add_paragraph().add_run(f"Image {number} · {status} · {media}")
            caption.italic = True
            caption.font.size = Pt(8)
        elif kind == "image_ref":
            doc.add_paragraph().add_run(f"[Image {block[1]} (shown above)]").italic = True
        elif kind == "image_note":
            doc.add_paragraph().add_run(f"[Image {block[1]}: {block[2]} — {block[3] or ''}; {block[4]}]").italic = True
    flush_table()
    doc.add_page_break()
    doc.add_heading("ORT labels (index by validation status)", level=1)
    for line in label_lines(content):
        paragraph = doc.add_paragraph(clean(line.strip()), style=None if line.startswith("[") else "List Bullet")
        if line.startswith("["):
            paragraph.runs[0].bold = True
    doc.save(target)


# ----------------------------------------------------------------------------- XLSX
def rebuild_xlsx(content, manifest, out_dir, target):
    from openpyxl import Workbook
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.comments import Comment
    from openpyxl.drawing.image import Image as SheetImage
    from openpyxl.styles import Font, PatternFill

    def safe(value):
        return ILLEGAL_CHARACTERS_RE.sub("�", value) if isinstance(value, str) else value

    def add_comment(cell, text):
        cell.comment = Comment(text + (f"\n{cell.comment.text}" if cell.comment else ""), "ORT")

    content.build()  # numbers the images as in recovered_report.html
    workbook = Workbook()
    info = workbook.active
    info.title = "ORT recovery"
    info.sheet_properties.tabColor = "FFC000"
    lines = info_lines(manifest) + [""]
    shared = content.shared_strings()
    styles = content.entries.get("xl/styles.xml")
    formats = parse_number_formats(styles["decoded_xml"]) if styles and styles.get("decoded_xml") else {}
    used_titles = {"ORT recovery"}
    lines.append("Sheets")
    for index, (name, part) in enumerate(content.sheets(), 1):
        status = content.status(part)
        title = re.sub(r"[\[\]:*?/\\]", "_", name or f"Sheet{index}")[:31]
        while title in used_titles:
            title = f"{title[:28]}_{index}"
        used_titles.add(title)
        lines.append(f"{title} ← {part}: {status or 'not found'} {SHORT.get(status, '')}")
        sheet = workbook.create_sheet(title)
        if status != "crc_verified":
            sheet.sheet_properties.tabColor = STATUS_COLOURS.get(status or "not found")[1:]
        item = content.entries.get(part)
        last_row = 0
        for cell in parse_sheet_cells(item["decoded_xml"]) if item and item.get("decoded_xml") else []:
            kind, raw = cell["type"], cell["raw"]
            if kind == "s" and raw is not None and raw.strip().isdigit():
                value = shared[int(raw)] if int(raw) < len(shared) else f"[shared string #{raw} not recovered]"
            elif kind == "inlineStr":
                value = cell["inline"] or ""
            elif kind == "b":
                value = raw == "1"
            elif kind in ("str", "e"):
                value = raw or ""
            else:
                try:
                    number = float(raw) if raw not in (None, "") else None
                    value = int(number) if number is not None and number.is_integer() else number
                except ValueError:
                    value = raw
            if value in (None, "") and not cell["formula"]:
                continue
            target_cell = sheet[cell["ref"]]
            target_cell.value = safe(value)
            if cell.get("style") in formats and not isinstance(value, str):
                target_cell.number_format = formats[cell["style"]]
            if cell["formula"]:
                add_comment(target_cell, f"Formula in the original: ={cell['formula']}\n(the stored value is shown)")
            last_row = max(last_row, cell["row"])
        if SHORT.get(status) and status != "crc_verified":
            add_comment(sheet["A1"], f"ORT: {part} is {status or 'not found'}. {SHORT.get(status)}")
        if status == "prefix_validated":
            marker = sheet.cell(row=last_row + 2, column=1,
                                value="ORT: damage point — the rows above may end with non-original values.")
            marker.font = Font(bold=True, color="C00000")
            marker.fill = PatternFill("solid", fgColor="FFF2CC")
        for row, anchored in content.anchored_images(part).items():
            for reference, media in anchored:
                record = content.images.get(media)
                if embeddable(record):
                    image = SheetImage(str(out_dir / record["output"]))
                    image.anchor = reference
                    sheet.add_image(image)
                else:
                    where = record.get("output") if record else None
                    add_comment(sheet[reference], f"ORT: Image {content.numbers.get(media, '?')} ({media}) is anchored here "
                                                  "but cannot be embedded" + (f"; file {where}" if where else "; not recovered"))
    lines += ["", "Labels (index by validation status)"] + label_lines(content)
    for row, line in enumerate(lines, 1):
        cell = info.cell(row=row, column=1, value=safe(line))
        cell.font = Font(bold=line in (BANNER, "Sheets", "Labels (index by validation status)") or line.startswith("["),
                         color="C00000" if line == BANNER else None)
    info.column_dimensions["A"].width = 120
    workbook.save(target)


# ----------------------------------------------------------------------------- PPTX
def rebuild_pptx(content, manifest, out_dir, target):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Emu, Pt

    content.build()
    presentation = Presentation()
    original = content.xml("ppt/presentation.xml")
    size = next((e for e in original.iter() if local_name(e) == "sldSz"), None) if original is not None else None
    if size is not None and (size.get("cx") or "").isdigit() and (size.get("cy") or "").isdigit():
        presentation.slide_width, presentation.slide_height = Emu(int(size.get("cx"))), Emu(int(size.get("cy")))
    blank = presentation.slide_layouts[6]
    width, height = presentation.slide_width, presentation.slide_height

    def textbox(slide, left, top, box_width, box_height, lines, size_pt=14, colour=None, fill=None, bold=False):
        box = slide.shapes.add_textbox(left, top, box_width, box_height)
        frame = box.text_frame
        frame.word_wrap = True
        for index, line in enumerate(lines):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            run = paragraph.add_run()
            run.text = clean(line)
            run.font.size = Pt(size_pt)
            run.font.bold = bold
            if colour:
                run.font.color.rgb = RGBColor.from_string(colour)
        if fill:
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor.from_string(fill)

    def text_slides(title, lines, per_slide=16):
        for start in range(0, max(len(lines), 1), per_slide):
            slide = presentation.slides.add_slide(blank)
            textbox(slide, Emu(width // 20), Emu(height // 30), Emu(width * 9 // 10), Emu(height // 10), [title], 22, bold=True)
            textbox(slide, Emu(width // 20), Emu(height // 6), Emu(width * 9 // 10), Emu(height * 4 // 5),
                    lines[start:start + per_slide], 11)

    def position(element):
        node = next((e for e in element.iter() if local_name(e) == "xfrm"), None)
        offset = next((e for e in node if local_name(e) == "off"), None) if node is not None else None
        extent = next((e for e in node if local_name(e) == "ext"), None) if node is not None else None
        if offset is None or extent is None:
            return None
        try:
            values = tuple(int(v) for v in (offset.get("x", 0), offset.get("y", 0), extent.get("cx", 0), extent.get("cy", 0)))
        except ValueError:
            return None
        return values if values[2] > 0 and values[3] > 0 else None

    text_slides("Rebuilt by ORT — not the original", info_lines(manifest), per_slide=12)
    for index, part in enumerate(content.slides(), 1):
        status = content.status(part)
        slide = presentation.slides.add_slide(blank)
        rels, root = content.relationships(part), content.xml(part)
        tree = next((e for e in root.iter() if local_name(e) == "spTree"), None) if root is not None else None
        cursor = height // 8
        stack = [iter(tree)] if tree is not None else []
        while stack:  # shapes in document order, without recursion
            child = next(stack[-1], None)
            if child is None:
                stack.pop()
                continue
            name = local_name(child)
            if name in ("grpSp",):
                stack.append(iter(child))  # group children keep their own (group-space) positions
                continue
            if name == "AlternateContent":
                choice = next((e for e in child if local_name(e) == "Choice"), None)
                if choice is not None:
                    stack.append(iter(choice))
                continue
            if name not in ("sp", "pic", "graphicFrame"):
                continue
            box = position(child)
            if box is None:  # e.g. a placeholder that inherits its position from the layout
                box = (width // 20, cursor, width * 9 // 10, height // 8)
                cursor += height // 7
            left, top, box_width, box_height = (Emu(value) for value in box)
            if name == "sp":
                lines, sizes = [], []
                for paragraph in (e for e in child.iter() if local_name(e) == "p"):
                    lines.append(content.runs(paragraph, rels, []))
                    run_properties = next((e for e in paragraph.iter()
                                           if local_name(e) in ("rPr", "endParaRPr") and (e.get("sz") or "").isdigit()), None)
                    sizes.append(int(run_properties.get("sz")) / 100 if run_properties is not None else 18)
                if any(line.strip() for line in lines):
                    textbox(slide, left, top, box_width, box_height, lines, size_pt=max(8, min(sizes)))
            elif name == "pic":
                pending = []
                content.runs(child, rels, pending)
                for media in pending:
                    record = content.images.get(media)
                    if embeddable(record):
                        slide.shapes.add_picture(str(out_dir / record["output"]), left, top, box_width, box_height)
                    else:
                        textbox(slide, left, top, box_width, box_height,
                                [f"[Image {content.numbers.get(media, '?')}: {media} cannot be embedded"
                                 + (f"; file {record['output']}" if record and record.get("output") else "") + "]"],
                                10, fill="EEEEEE")
            else:
                rows = [[content.runs(c, rels, []).replace("\n", " ") for c in r if local_name(c) == "tc"]
                        for r in child.iter() if local_name(r) == "tr"]
                rows = [row for row in rows if row]
                if rows:
                    table = slide.shapes.add_table(len(rows), max(len(r) for r in rows), left, top, box_width, box_height).table
                    for r, cells in enumerate(rows):
                        for c, text in enumerate(cells):
                            table.cell(r, c).text = clean(text)
        if status != "crc_verified":
            textbox(slide, Emu(0), Emu(0), width, Emu(height // 14),
                    [f"ORT · slide {index} ({part}): {status or 'not found'} — {SHORT.get(status, '')}"],
                    12, colour="FFFFFF", fill=STATUS_COLOURS.get(status or "not found")[1:], bold=True)
        notes = content.notes(part, rels) + [f"[ORT] source part {part}: {status or 'not found'}"]
        slide.notes_slide.notes_text_frame.text = clean("\n".join(notes))
    text_slides("ORT labels (index by validation status)", label_lines(content))
    presentation.save(target)


def write_rebuilt_document(content, manifest, output_directory):
    """Write recovered_<name><ext> into the output folder and return its file name."""
    name = rebuilt_name(manifest)
    writer = {".docx": rebuild_docx, ".xlsx": rebuild_xlsx, ".pptx": rebuild_pptx}[manifest["inferred_type"]]
    writer(content, manifest, Path(output_directory), Path(output_directory) / name)
    return name
