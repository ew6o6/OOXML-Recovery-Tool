"""Reading outputs of one input: recovered_report.html, the rebuilt document, and its PDF preview.

recovered_report.html (static: no script, no external resource; all recovered text HTML-escaped)
  Header            the "not the original" notice, input hash, statuses, links to the rebuilt document and PDF
  Guide             what each label means, its evidence, and how to treat it in a report
  Overview          label counts and document properties
  Contents          reading order with the validation status of every part
  Recovered content reading order; every section names its part and status and warns where content may be
                    non-original; images appear where they are referenced and link to their files in img/
  Label index       parts and images by validation status

Reading order: DOCX body, headers, footers, footnotes, endnotes, comments; XLSX sheets in workbook order,
with images at their anchor cells; PPTX slides in presentation order, with pictures and speaker notes.
XML is walked with explicit stacks: damaged XML recovered by the lenient parser can nest deeper than
Python's recursion limit. The outputs are reading aids; the exact text and parts remain in text/ and parts/.
"""

import html
import posixpath
import re
from collections import Counter
from pathlib import Path

from .handlers.common import column_letters, column_number, parse_shared_strings, parse_sheet_cells, resolve_cell
from .utils import STATUS_NOTES, local_name, parse_xml

HTML_FILENAME = "recovered_report.html"
LABEL_ORDER = ["crc_verified", "structure_verified", "prefix_validated", "crc_mismatch", "unverified", "rejected"]
NOTES = {**STATUS_NOTES, "rejected": "no evidence of being the claimed content; not written"}
WARNINGS = {
    "prefix_validated": "Damage point: this part is recovered up to here; the text just above may be non-original.",
    "crc_mismatch": "Warning: this part differs from the original (CRC-32 mismatch); where it differs is unknown.",
    "unverified": "Warning: whether this part is original cannot be confirmed.",
    "structure_verified": "Note: complete and well-formed, but there is no CRC-32 record to confirm it is original.",
}
CORE_KEYS = ["title", "subject", "creator", "lastModifiedBy", "created", "modified", "lastPrinted", "revision"]
APP_KEYS = ["Application", "AppVersion", "Company", "TotalTime", "Pages", "Words", "Slides"]
DOCX_GROUPS = [("Header", r"^word/header\d*\.xml$"), ("Footer", r"^word/footer\d*\.xml$"),
               ("Footnotes", r"^word/footnotes\.xml$"), ("Endnotes", r"^word/endnotes\.xml$"),
               ("Comments", r"^word/comments\.xml$")]
DOCX_CONTAINERS = {"sdt", "sdtContent", "body", "customXml", "hdr", "ftr", "footnote", "endnote", "comment"}
BROWSER_IMAGES = {"png", "jpeg", "jpg", "gif", "bmp", "webp", "svg"}
LABEL_GUIDE = {  # what the label means · evidence · how to treat the content in a report
    "crc_verified": (
        "Byte-identical to the original.",
        "The CRC-32 of the recovered bytes equals the value recorded in the ZIP structure (local file header, "
        "central directory, or data descriptor). CRC-32 detects accidental damage; it is not proof against "
        "deliberate recomputation.",
        "May be presented as original content."),
    "structure_verified": (
        "Complete and well-formed, but not shown to be original.",
        "No CRC-32 value was available to compare. The part is well-formed XML, or the image decodes completely.",
        "State that the content is complete but could not be compared with the original."),
    "prefix_validated": (
        "Only the start of the part was recovered.",
        "DEFLATE is decoded sequentially (RFC 1951), so output before the damaged bytes equals the original. Where "
        "the damage begins in the output is unknown, so the end may contain non-original text.",
        "Use the beginning; treat the text just before the marked damage point with caution."),
    "crc_mismatch": (
        "Decoded completely, but altered.",
        "The CRC-32 of the recovered bytes differs from the recorded value: some content differs from the original. "
        "Where it differs is unknown.",
        "Indicative only; do not present as original content."),
    "unverified": (
        "Not confirmed.",
        "No CRC-32 value was available, and the part is not well-formed (or no validator exists for the image format).",
        "Indicative only; may differ from the original."),
    "rejected": (
        "Not recovered.",
        "There is no evidence that the bytes are the claimed content, so nothing was written. Uncompressed images may "
        "be kept as diagnostic files.",
        "Report only that the part existed at the recorded location."),
}


def attribute(element, name):
    """Attribute by local name (namespace or an undeclared prefix removed)."""
    for key, value in element.attrib.items():
        if key.rsplit("}", 1)[-1].rsplit(":", 1)[-1] == name:
            return value
    return None


def rels_name(part):
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def part_number(part):
    match = re.search(r"(\d+)\.xml$", part)
    return int(match.group(1)) if match else 0


class Content:
    """The recovered content of one input as typed blocks, in reading order.

    Blocks: ("section", heading, part, status, anchor), ("warning", status, text), ("note", text),
    ("text", text), ("table_row", [cells]), ("row", ["Row n", "A1=value", ...]),
    ("image", number, media, status, file, format, problem), ("image_ref", number),
    ("image_note", number, media, status, detail), ("heading", text). Image markers stay in text as [Image N].
    """

    def __init__(self, entries, manifest):
        self.manifest = manifest
        self.entries = {}
        for item in entries:
            self.entries.setdefault(item.get("local_file_name", ""), item)
        self.images = {record["name"]: record for record in manifest.get("images", [])}
        self.numbers, self.embedded, self.anchors, self.chars = {}, set(), {}, {}
        self.blocks = []

    # ------------------------------------------------------------------ parts and relationships
    def status(self, part):
        item = self.entries.get(part)
        return item.get("content_validation_status") if item else None

    def xml(self, part):
        item = self.entries.get(part)
        return parse_xml(item.get("decoded_xml")) if item and item.get("decoded_xml") else None

    def relationships(self, part):
        """Relationship id -> (target part, type) from the recovered .rels part of a source part."""
        root = self.xml(rels_name(part))
        result = {}
        for element in root.iter() if root is not None else ():
            if local_name(element) != "Relationship" or element.get("TargetMode") == "External":
                continue
            target = element.get("Target") or ""
            target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(
                posixpath.join(posixpath.dirname(part), target))
            result[element.get("Id")] = (target, element.get("Type") or "")
        return result

    def sheets(self):
        """[(display name or None, sheet part)] in workbook order, then sheets the workbook does not list."""
        workbook_rels, workbook = self.relationships("xl/workbook.xml"), self.xml("xl/workbook.xml")
        sheets = [(sheet.get("name"), workbook_rels[attribute(sheet, "id")][0])
                  for sheet in (workbook.iter() if workbook is not None else ())
                  if local_name(sheet) == "sheet" and attribute(sheet, "id") in workbook_rels]
        known = {part for _, part in sheets}
        return sheets + [(None, part) for part in sorted(self.entries, key=part_number)
                         if re.match(r"^xl/worksheets/sheet\d+\.xml$", part) and part not in known]

    def slides(self):
        presentation_rels, presentation = self.relationships("ppt/presentation.xml"), self.xml("ppt/presentation.xml")
        slides = [presentation_rels[attribute(element, "id")][0]
                  for element in (presentation.iter() if presentation is not None else ())
                  if local_name(element) == "sldId" and attribute(element, "id") in presentation_rels]
        return slides + [part for part in sorted(self.entries, key=part_number)
                         if re.match(r"^ppt/slides/slide\d+\.xml$", part) and part not in slides]

    def shared_strings(self):
        item = self.entries.get("xl/sharedStrings.xml")
        return parse_shared_strings(item["decoded_xml"]) if item and item.get("decoded_xml") else []

    def anchored_images(self, sheet_part):
        """{row: [(cell reference, media part)]} from the drawings of a worksheet."""
        anchors = {}
        for target, kind in self.relationships(sheet_part).values():
            if not kind.endswith("/drawing"):
                continue
            drawing, drawing_rels = self.xml(target), self.relationships(target)
            for anchor in drawing.iter() if drawing is not None else ():
                if local_name(anchor) not in ("twoCellAnchor", "oneCellAnchor"):
                    continue
                start = next((element for element in anchor if local_name(element) == "from"), None)
                cells = {local_name(element): element.text for element in start} if start is not None else {}
                row, column = int(cells.get("row") or 0) + 1, int(cells.get("col") or 0) + 1
                for blip in (element for element in anchor.iter() if local_name(element) == "blip"):
                    if attribute(blip, "embed") in drawing_rels:
                        anchors.setdefault(row, []).append(
                            (f"{column_letters(column)}{row}", drawing_rels[attribute(blip, "embed")][0]))
        return anchors

    # ------------------------------------------------------------------ images
    def marker(self, media):
        if media not in self.numbers:
            self.numbers[media] = len(self.numbers) + 1
        return f"[Image {self.numbers[media]}]"

    def image_blocks(self, media):
        number = self.numbers[media]
        if media in self.embedded:
            return [("image_ref", number)]
        self.embedded.add(media)
        record = self.images.get(media)
        if record is None:
            return [("image_note", number, media, self.status(media), "was not recovered")]
        status = record.get("content_validation_status")
        if record.get("output"):
            check = record.get("format_validation")
            # unsupported_format (e.g. EMF) is not a decoding failure; the renderer names the format instead
            problem = "" if check in ("decodable_complete", "unsupported_format") else f"format check: {check}"
            return [("image", number, media, status, record["output"], record.get("detected_format") or "", problem)]
        if record.get("uniform_fill_byte"):
            detail = f"its data area is filled with {record['uniform_fill_byte']} ({record.get('rejected_bytes', 0):,} bytes)"
        elif record.get("diagnostic_output"):
            detail = f"not an image; bytes kept for examination in {record['diagnostic_output']}"
        else:
            detail = "no image was recovered"
        return [("image_note", number, media, status or record.get("status"), detail)]

    # ------------------------------------------------------------------ text
    def runs(self, node, rels, pending):
        """Text of a node in document order with image markers; mc:Fallback copies are skipped."""
        pieces, stack = [], list(reversed(list(node)))
        while stack:
            child = stack.pop()
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
                relationship = attribute(child, "embed" if name == "blip" else "id")
                if relationship in rels:
                    pieces.append(self.marker(rels[relationship][0]))
                    pending.append(rels[relationship][0])
            stack.extend(reversed(list(child)))
        return "".join(pieces)

    def with_images(self, block, pending):
        has_text = bool(block[1].strip()) if block[0] == "text" else any(cell.strip() for cell in block[1])
        out = [block] if has_text else []
        for media in pending:
            out += self.image_blocks(media)
        pending.clear()
        return out

    def section(self, heading, part, body):
        status = self.status(part)
        anchor = f"part-{len(self.anchors) + 1}"
        self.anchors[part] = (anchor, heading)
        self.blocks.append(("section", heading, part, status or "not found", anchor))
        if status in (None, "rejected") or not body:
            self.blocks.append(("note", "Rejected; not written." if status == "rejected" else "No recovered content."))
            return
        if status in ("crc_mismatch", "unverified", "structure_verified"):
            self.blocks.append(("warning", status, WARNINGS[status]))
        self.blocks += body
        if status == "prefix_validated":
            self.blocks.append(("warning", status, WARNINGS[status]))
        self.chars[part] = sum(len(b[1]) if b[0] == "text" else sum(map(len, b[1])) if b[0] in ("table_row", "row")
                               else 0 for b in body)

    # ------------------------------------------------------------------ formats
    def docx(self):
        def blocks(container, rels):
            out, stack = [], [iter(container)]  # containers are entered in place, without recursion
            while stack:
                child = next(stack[-1], None)
                if child is None:
                    stack.pop()
                    continue
                name, pending = local_name(child), []
                if name == "p":
                    out += self.with_images(("text", self.runs(child, rels, pending)), pending)
                elif name == "tbl":
                    for row in (element for element in child.iter() if local_name(element) == "tr"):
                        cells = [self.runs(cell, rels, pending).replace("\n", " ").strip()
                                 for cell in row if local_name(cell) == "tc"]
                        out += self.with_images(("table_row", cells), pending)
                elif name in DOCX_CONTAINERS:
                    stack.append(iter(child))
            return out

        root = self.xml("word/document.xml")
        body = next((e for e in root.iter() if local_name(e) == "body"), root) if root is not None else None
        self.section("Body", "word/document.xml",
                     blocks(body, self.relationships("word/document.xml")) if body is not None else [])
        for heading, pattern in DOCX_GROUPS:
            for part in sorted((name for name in self.entries if re.match(pattern, name)), key=part_number):
                root = self.xml(part)
                content = blocks(root, self.relationships(part)) if root is not None else []
                if content or self.status(part) == "rejected":
                    self.section(heading, part, content)

    def xlsx(self):
        shared = self.shared_strings()
        if "xl/sharedStrings.xml" in self.entries:
            self.blocks.append(("note", f"Shared strings (xl/sharedStrings.xml): {self.status('xl/sharedStrings.xml')}. "
                                        "String cell values are taken from this part."))
        for index, (name, part) in enumerate(self.sheets(), 1):
            anchors = self.anchored_images(part)
            item, rows = self.entries.get(part), {}
            for cell in parse_sheet_cells(item["decoded_xml"]) if item and item.get("decoded_xml") else []:
                value, _ = resolve_cell(cell, shared)
                if value != "" or cell["formula"]:
                    rows.setdefault(cell["row"], []).append(
                        f"{cell['ref']}={value}" + (f" (formula ={cell['formula']})" if cell["formula"] else ""))
            body = [("row", [f"Row {row}"] + rows[row]) for row in sorted(rows)]
            for row in sorted(anchors):  # anchored images follow the grid of the sheet
                for reference, media in anchors[row]:
                    body.append(("text", f"At {reference}: {self.marker(media)}"))
                    body += self.image_blocks(media)
            self.section(f"Sheet {index}" + (f" '{name}'" if name else " (name unknown)"), part, body)

    def pptx(self):
        def shapes(container, rels, out):
            stack = [iter(container)]  # groups are entered in place, without recursion
            while stack:
                child = next(stack[-1], None)
                if child is None:
                    stack.pop()
                    continue
                name, pending = local_name(child), []
                if name == "sp":
                    paragraphs = [self.runs(p, rels, pending) for p in child.iter() if local_name(p) == "p"]
                    out += self.with_images(("text", "\n".join(p for p in paragraphs if p.strip())), pending)
                elif name == "pic":
                    self.runs(child, rels, pending)
                    for media in pending:
                        out.append(("text", self.marker(media)))
                        out += self.image_blocks(media)
                elif name == "graphicFrame":
                    for row in (element for element in child.iter() if local_name(element) == "tr"):
                        cells = [self.runs(cell, rels, pending).replace("\n", " ").strip()
                                 for cell in row if local_name(cell) == "tc"]
                        out += self.with_images(("table_row", cells), pending)
                elif name in ("grpSp", "spTree"):
                    stack.append(iter(child))
                elif name == "AlternateContent":
                    choice = next((element for element in child if local_name(element) == "Choice"), None)
                    if choice is not None:
                        stack.append(iter(choice))

        for index, part in enumerate(self.slides(), 1):
            rels, root, body = self.relationships(part), self.xml(part), []
            tree = next((e for e in root.iter() if local_name(e) == "spTree"), None) if root is not None else None
            if tree is not None:
                shapes(tree, rels, body)
            self.section(f"Slide {index}", part, body)
            notes = self.notes(part, rels)
            if notes or any(kind.endswith("/notesSlide") and self.status(target) == "rejected"
                            for target, kind in rels.values()):
                target = next(t for t, k in rels.values() if k.endswith("/notesSlide"))
                self.section(f"Slide {index} notes", target, [("text", text) for text in notes])

    def notes(self, slide_part, rels=None):
        """Speaker-note paragraphs of a slide (body placeholder of its notes slide)."""
        texts = []
        for target, kind in (rels or self.relationships(slide_part)).values():
            if not kind.endswith("/notesSlide"):
                continue
            notes = self.xml(target)
            for shape in (e for e in notes.iter() if local_name(e) == "sp") if notes is not None else ():
                placeholder = next((e for e in shape.iter() if local_name(e) == "ph"), None)
                if placeholder is not None and placeholder.get("type") == "body":
                    texts += [self.runs(p, {}, []) for p in shape.iter() if local_name(p) == "p"]
        return [text for text in texts if text.strip()]

    def build(self):
        {".docx": self.docx, ".xlsx": self.xlsx, ".pptx": self.pptx}[self.manifest["inferred_type"]]()
        unplaced = [name for name in self.images if name not in self.numbers]
        if unplaced:
            self.blocks.append(("heading", "Images outside the content or without a known position"))
            for media in unplaced:
                self.marker(media)
                users = [part for part in self.entries if not part.endswith(".rels")
                         and any(target == media for target, _ in self.relationships(part).values())]
                self.blocks.append(("text", f"Image {self.numbers[media]}: " + (
                    "referenced by " + ", ".join(users) if users else "no recovered part references it")))
                self.blocks += self.image_blocks(media)
        return self.blocks

    def label_rows(self, status):
        """[(part or image, anchor or None, location)] of one validation status, for the label index."""
        rows = []
        for part, item in self.entries.items():
            if part in self.images or item.get("content_validation_status") != status:
                continue
            span = item.get("decoded_span") or item.get("declared_span")
            where = f"header at byte {item.get('source_offset', 0):,}" + (f", data {span[0]:,}–{span[1]:,}" if span else "")
            if part in self.chars:
                anchor, heading = self.anchors[part]
                rows.append((f"{part}: “{heading}”", anchor, f"{where}; {self.chars[part]:,} characters"))
            else:
                rows.append((part, None, where))
        for media, record in self.images.items():
            if record.get("content_validation_status") == status:
                target = f" → {record['output']}" if record.get("output") else ""
                rows.append((f"Image {self.numbers.get(media, '')}: {media}", None,
                             f"header at byte {record.get('source_offset') or 0:,}{target}"))
        return rows


# ----------------------------------------------------------------------------- HTML
STATUS_COLOURS = {"crc_verified": "#2F7D57", "structure_verified": "#3A6EA5", "prefix_validated": "#EE7A3C",
                  "crc_mismatch": "#DE4B55", "unverified": "#8A847D", "rejected": "#3D3936", "not found": "#B9B3AD"}
OVERALL_COLOURS = {"complete": "#2F7D57", "partial": "#EE7A3C", "failed": "#DE4B55"}
CSS = """
:root{--red:#DE4B55;--red-dark:#C23440;--orange:#EE7A3C;--peach:#F6D3BE;--pink:#FCEDEE;--beige:#FAF5F0;
--line:#ECE5DE;--text:#1A1A1A;--sub:#5F5A55;--muted:#8A847D;--dark:#1E1C1B}
*{box-sizing:border-box}html{scroll-behavior:smooth;scroll-padding-top:72px}
body{margin:0;background:#fff;color:var(--text);font:15px/1.65 "Pretendard Variable",Pretendard,"Malgun Gothic","Apple SD Gothic Neo","Noto Sans KR",sans-serif;-webkit-font-smoothing:antialiased}
a{color:inherit;text-decoration:none}.inner{max-width:1200px;margin:0 auto;padding:0 32px}
header.top{position:sticky;top:0;z-index:10;background:rgba(255,255,255,.96);border-bottom:1px solid var(--line)}
header.top .inner{display:flex;align-items:center;gap:28px;height:60px}
.brand{font-weight:800;letter-spacing:.02em;font-size:20px}.brand small{display:block;font-size:10px;font-weight:600;color:var(--muted);letter-spacing:.08em;margin-top:-4px}
nav{margin-left:auto;display:flex;gap:24px;font-size:13px;font-weight:600;color:var(--sub)}nav a:hover{color:var(--red)}
.pill{display:inline-block;padding:4px 12px;border-radius:999px;font-size:12px;font-weight:700;color:#fff;white-space:nowrap}
.hero{background:radial-gradient(1200px 400px at 85% 0%,#0a1a33 0%,var(--dark) 60%);color:#fff;padding:56px 0 48px}
.eyebrow{font-size:12px;font-weight:800;letter-spacing:.12em;color:var(--red);text-transform:uppercase}
.hero h1{font-size:34px;line-height:1.3;margin:10px 0 6px;word-break:break-all}.hero .warn-line{color:#FFB79C;font-weight:600}
.hero .hash{font:12px Consolas,monospace;color:#B9B3AD;margin-top:10px;word-break:break-all}
.hero .files{margin-top:14px;font-size:13px;color:#ECE5DE}.hero .files a{color:#fff;text-decoration:underline}
.hero .files .how{color:#B9B3AD;font-size:12px}
.stats{display:grid;grid-template-columns:repeat(4,1fr);margin:30px 0 0;border-top:1px solid #3d3936}
.stats div{padding:16px 18px 0 0}.stats dt{font-size:12px;color:#B9B3AD}.stats dd{margin:4px 0 0;font-size:22px;font-weight:800}
section.block{padding:56px 0}section.beige{background:var(--beige)}
.sec-head{border-bottom:1px solid var(--peach);padding-bottom:12px;margin-bottom:28px;font-size:12px;font-weight:800;letter-spacing:.12em;color:var(--red);text-transform:uppercase}
h2{font-size:28px;line-height:1.35;margin:0 0 8px}h2 em{font-style:normal;color:var(--red)}.lead{color:var(--sub);margin:0 0 24px}
.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
.card{background:#fff;border:1px solid var(--line);padding:18px 20px}
.card .k{font-size:11px;font-weight:800;letter-spacing:.1em}.card .n{font-size:26px;font-weight:800;margin:2px 0}
.card p{margin:0;color:var(--sub);font-size:13px}
.props{display:grid;grid-template-columns:160px 1fr;gap:6px 18px;margin-top:28px;font-size:14px}
table.guide{width:100%;font-size:14px}table.guide td,table.guide th{white-space:normal;padding:10px 12px;border-left:none;border-right:none}
table.guide td:first-child{white-space:nowrap}.legend dt{display:flex;align-items:center}.props dt{color:var(--muted)}.props dd{margin:0;word-break:break-word}
ol.steps{list-style:none;margin:0;padding:0;display:grid;grid-template-columns:repeat(4,1fr)}
ol.steps li{border-left:1px solid var(--line);padding:4px 18px 18px}ol.steps li:nth-child(4n+1){border-left:none;padding-left:0}
ol.steps .num{color:var(--red);font-size:12px;font-weight:800}ol.steps .t{font-weight:700;margin:2px 0}
ol.steps .s{font-size:12px;color:var(--sub);display:flex;align-items:center;gap:6px}.dot{width:8px;height:8px;border-radius:50%}
article.part{border:1px solid var(--line);margin:0 0 22px;background:#fff}
article.part>.ph{display:flex;align-items:center;gap:12px;padding:14px 20px;border-bottom:1px solid var(--line);background:#fffdfb}
article.part h3{margin:0;font-size:18px}.path{font:12px Consolas,monospace;color:var(--muted)}
article.part .ph .pill{margin-left:auto}article.part .pb{padding:16px 20px}
.text{white-space:pre-wrap;margin:0 0 10px;word-break:break-word}
.callout{border-radius:4px;padding:10px 14px;margin:0 0 14px;font-size:14px;font-weight:600}
.callout.prefix_validated{background:#FDF0E6;color:#B4501C;border:1px solid #F6D3BE}
.callout.crc_mismatch{background:var(--pink);color:var(--red-dark);border:1px solid #F5C2C6}
.callout.unverified{background:#F3F1EF;color:#5F5A55;border:1px solid var(--line)}
.callout.structure_verified{background:#EEF3FA;color:#2F5A88;border:1px solid #D3E0F1}
.note{color:var(--muted);font-size:13px;margin:0 0 10px}
.mark{background:#FFF1C2;border-radius:3px;padding:0 4px;font-weight:700;font-size:.9em}
figure{margin:12px 0 18px;display:inline-block;border:1px solid var(--line);padding:10px;max-width:100%;background:#fff}
figure img{display:block;max-width:100%;max-height:520px}figcaption{font-size:12px;color:var(--sub);margin-top:6px}
.tablewrap{overflow-x:auto;margin:6px 0 14px}table{border-collapse:collapse;font-size:13px}
th,td{border:1px solid var(--line);padding:4px 8px;vertical-align:top;white-space:pre-wrap}
th{background:var(--beige);color:var(--sub);font-weight:700}td.rn{background:var(--beige);color:var(--muted);text-align:right}
td .f{display:block;font:11px Consolas,monospace;color:var(--muted)}
details{background:#fff;border:1px solid var(--line);margin:0 0 12px}
details summary{list-style:none;cursor:pointer;display:flex;align-items:center;gap:12px;padding:16px 20px;font-weight:700}
details summary::-webkit-details-marker{display:none}details summary::after{content:"+";margin-left:auto;color:var(--red);font-size:20px;font-weight:400}
details[open] summary::after{content:"\\2212"}details summary .cnt{color:var(--muted);font-weight:600;font-size:13px}
details ul{margin:0;padding:0 20px 16px;list-style:none}details li{padding:7px 0;border-top:1px solid var(--line);font-size:14px}
details li a{color:var(--red-dark);font-weight:600}details li .loc{color:var(--muted);font-size:12px;margin-left:6px}
footer{background:var(--dark);color:#B9B3AD;font-size:12px;padding:26px 0}footer b{color:#fff}
@media (max-width:820px){.stats,.cards,ol.steps{grid-template-columns:1fr 1fr}.inner{padding:0 18px}.props{grid-template-columns:1fr}}
@media print{header.top{position:static}.hero{background:#fff;color:#000}.hero .files a{color:#000}
section.block{padding:24px 0}article.part,figure{break-inside:avoid}}
"""


def esc(text):
    return html.escape(str(text), quote=True)


def pill(status, text=None, colour=None):
    colour = colour or STATUS_COLOURS.get(status or "not found", STATUS_COLOURS["not found"])
    return f'<span class="pill" style="background:{colour}">{esc(text or status or "not found")}</span>'


def with_marks(escaped):
    return re.sub(r"\[Image (\d+)\]", r'<span class="mark">Image \1</span>', escaped)


def grid(rows):
    """Sheet rows (["Row n", "A1=v", ...]) as a table with column letters and row numbers."""
    parsed, columns = [], set()
    for row in rows:
        cells = {}
        for item in row[1:]:
            reference, _, value = item.partition("=")
            letters = re.match(r"[A-Z]+", reference)
            letters = letters.group(0) if letters else reference
            formula = ""
            match = re.search(r" \(formula =(.*)\)$", value)
            if match:
                value, formula = value[:match.start()], match.group(1)
            cells[letters] = (value, formula)
            columns.add(letters)
        parsed.append((row[0].split()[-1], cells))
    ordered = sorted(columns, key=column_number)
    out = ['<div class="tablewrap"><table><tr><th></th>' + "".join(f"<th>{esc(c)}</th>" for c in ordered) + "</tr>"]
    for number, cells in parsed:
        out.append(f'<tr><td class="rn">{esc(number)}</td>' + "".join(
            f"<td>{esc(cells.get(c, ('', ''))[0])}"
            + (f'<span class="f">={esc(cells[c][1])}</span>' if c in cells and cells[c][1] else "") + "</td>"
            for c in ordered) + "</tr>")
    out.append("</table></div>")
    return "".join(out)


def render_html(content, blocks, manifest, outputs):
    m = manifest
    name = Path(m.get("input", "")).name
    part_status = Counter(i.get("content_validation_status") for p, i in content.entries.items() if p not in content.images)
    image_status = Counter(r.get("content_validation_status") for r in content.images.values())
    sections = [b for b in blocks if b[0] == "section"]
    o = ['<!doctype html><html lang="en"><head><meta charset="utf-8">',
         "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; img-src 'self' file:; style-src 'unsafe-inline'\">",
         '<meta name="viewport" content="width=device-width,initial-scale=1">',
         f"<title>Recovered · {esc(name)}</title><style>{CSS}</style></head><body>",
         '<header class="top"><div class="inner"><a class="brand" href="#overview">ORT<small>RECOVERED DOCUMENT</small></a>'
         '<nav><a href="#guide">Guide</a><a href="#summary">Overview</a><a href="#contents">Contents</a>'
         '<a href="#content">Recovered content</a><a href="#labels">Label index</a></nav>'
         + pill(None, f"overall: {m.get('overall_status')}", OVERALL_COLOURS.get(m.get("overall_status"), "#8A847D"))
         + "</div></header>",
         '<div class="hero" id="overview"><div class="inner"><div class="eyebrow">Rebuilt by ORT · not the original</div>',
         f"<h1>{esc(name)}</h1>",
         '<div class="warn-line">Formatting is not reproduced. Parts marked below may be incomplete or non-original.</div>',
         f'<div class="hash">Input SHA-256 {esc(m.get("input_sha256"))}<br>ORT {esc(m.get("tool_version"))} · code SHA-256 '
         f'{esc(m.get("tool_code_sha256"))}</div>']
    rebuilt, pdf = outputs.get("rebuilt_document") or {}, outputs.get("preview_pdf") or {}
    files = []
    if rebuilt.get("path"):
        files.append(f'Rebuilt document: <a href="{esc(rebuilt["path"])}">{esc(rebuilt["path"])}</a>')
    if pdf.get("path"):
        files.append(f'PDF preview: <a href="{esc(pdf["path"])}">{esc(pdf["path"])}</a> ({esc(pdf.get("converter"))})')
    elif pdf:
        files.append(f'PDF preview: {esc(pdf.get("status"))} — {esc(pdf.get("reason"))}'
                     + "".join(f'<br><span class="how">{esc(step)}</span>' for step in pdf.get("how_to") or []))
    if files:
        o.append('<div class="files">' + "<br>".join(files) + "</div>")
    o.append('<dl class="stats">'
             f'<div><dt>Overall</dt><dd>{esc(m.get("overall_status"))}</dd></div>'
             f'<div><dt>Core content</dt><dd>{esc(m.get("core_content_status"))}</dd></div>'
             f'<div><dt>Images preserved</dt><dd>{m.get("images_preserved", 0)} / {m.get("images_total", 0)}</dd></div>'
             f'<div><dt>Parts verified</dt><dd>{part_status.get("crc_verified", 0)} / {sum(part_status.values())}</dd></div>'
             "</dl></div></div>")
    # 01 overview
    # 01 guide: what the labels mean (right below the dark header)
    o.append('<section class="block" id="guide"><div class="inner"><div class="sec-head">01 &nbsp;How to read this report</div>'
             '<h2>What the <em>labels</em> mean</h2><p class="lead">Every recovered part and image carries one label. '
             'Labels are assigned conservatively: content is called original only when a CRC-32 value shows it.</p>'
             '<div class="tablewrap"><table class="guide"><tr><th>Label</th><th>What it means</th><th>Evidence</th>'
             '<th>In a report</th></tr>')
    for status in LABEL_ORDER:
        meaning, evidence, use = LABEL_GUIDE[status]
        o.append(f"<tr><td>{pill(status)}</td><td><b>{esc(meaning)}</b></td><td>{esc(evidence)}</td><td>{esc(use)}</td></tr>")
    o.append('</table></div><dl class="props legend">'
             '<dt><span class="mark">Image 3</span></dt><dd>where an image appears in the content; the image follows.</dd>'
             f'<dt><span class="pill" style="background:{STATUS_COLOURS["prefix_validated"]}">damage point</span></dt>'
             '<dd>the end of a partially recovered part; the text just above may be non-original.</dd>'
             '<dt>Overall status</dt><dd><b>complete</b>: every part listed in the central directory was found, and every '
             'part is crc_verified or structure_verified; <b>partial</b>: otherwise, with some text or images recovered; '
             '<b>failed</b>: nothing was recovered; <b>indeterminate</b>: nothing was recovered and there is no central '
             'directory to compare with.</dd>'
             "<dt>Rebuilt document</dt><dd>the recovered content written anew in the input's format; it is not the "
             "original and does not reproduce its formatting.</dd></dl></div></section>")
    # 02 overview
    o.append('<section class="block beige" id="summary"><div class="inner"><div class="sec-head">02 &nbsp;Overview</div>'
             f'<h2>Validation status of <em>{sum(part_status.values())} parts</em> and {len(content.images)} images</h2>'
             '<p class="lead">Each part and image carries one status. Only a CRC-32 match shows that content is '
             'byte-identical to the original.</p><div class="cards">')
    for status in LABEL_ORDER:
        count = part_status.get(status, 0) + image_status.get(status, 0)
        if count:
            o.append(f'<div class="card"><div class="k" style="color:{STATUS_COLOURS[status]}">{esc(status)}</div>'
                     f'<div class="n">{count}</div><p>{esc(NOTES[status])}</p></div>')
    o.append("</div>")
    properties = []
    for part, keys in (("docProps/core.xml", CORE_KEYS), ("docProps/app.xml", APP_KEYS)):
        record = (m.get("metadata") or {}).get(part)
        for key in keys if record else ():
            if record["values"].get(key):
                properties.append(f"<dt>{esc(key)}</dt><dd>{esc(record['values'][key])} "
                                  f"<span class='note'>({esc(part)}, {esc(record.get('content_validation_status'))})</span></dd>")
    if properties:
        o.append('<dl class="props">' + "".join(properties) + "</dl>")
    o.append("</div></section>")
    # 02 contents
    o.append('<section class="block" id="contents"><div class="inner"><div class="sec-head">03 &nbsp;Contents</div>'
             '<h2>Reading order</h2><p class="lead">Select a part to go to it.</p><ol class="steps">')
    for index, (_, heading, part, status, anchor) in enumerate(sections, 1):
        o.append(f'<li><a href="#{anchor}" title="{esc(part)}"><div class="num">{index:02d}</div><div class="t">{esc(heading)}</div>'
                 f'<div class="s"><span class="dot" style="background:{STATUS_COLOURS.get(status, "#B9B3AD")}"></span>'
                 f"{esc(status)}</div></a></li>")
    o.append("</ol></div></section>")
    # 03 recovered content
    o.append('<section class="block beige" id="content"><div class="inner"><div class="sec-head">04 &nbsp;Recovered content</div>')
    open_part, rows = False, []

    def flush():
        if rows:
            o.append(grid(rows))
            rows.clear()

    for block in blocks:
        kind = block[0]
        if kind != "row":
            flush()
        if kind in ("section", "heading"):
            if open_part:
                o.append("</div></article>")
            if kind == "section":
                _, heading, part, status, anchor = block
                o.append(f'<article class="part" id="{anchor}"><div class="ph"><h3>{esc(heading)}</h3>'
                         f'<span class="path">{esc(part)}</span>{pill(status)}</div><div class="pb">')
            else:
                o.append(f'<article class="part"><div class="ph"><h3>{esc(block[1])}</h3></div><div class="pb">')
            open_part = True
        elif kind == "warning":
            o.append(f'<div class="callout {esc(block[1])}">{esc(block[2])}</div>')
        elif kind == "note":
            o.append(f'<p class="note">{esc(block[1])}</p>')
        elif kind == "text":
            o.append(f'<p class="text">{with_marks(esc(block[1]))}</p>')
        elif kind == "table_row":
            row_html = "<tr>" + "".join(f"<td>{with_marks(esc(cell))}</td>" for cell in block[1]) + "</tr>"
            if o[-1].startswith('<div class="tablewrap"><table class="doc">'):
                o[-1] = o[-1][:-len("</table></div>")] + row_html + "</table></div>"  # the same table continues
            else:
                o.append('<div class="tablewrap"><table class="doc">' + row_html + "</table></div>")
        elif kind == "row":
            rows.append(block[1])
        elif kind == "image":
            _, number, media, status, source, image_format, problem = block
            caption = f"Image {number} · {status}" + (f" · {problem}" if problem else "") + f" · {media}"
            if not problem and image_format.lower() in BROWSER_IMAGES:
                o.append(f'<figure><img src="{esc(source)}" alt="Image {number}"><figcaption>{esc(caption)}</figcaption></figure>')
            else:
                why = ("does not decode completely" if problem
                       else f"format {image_format or 'unknown'} cannot be shown by a browser")
                o.append(f'<figure><figcaption>{esc(caption)}<br>The file {why}: '
                         f'<a href="{esc(source)}"><b>{esc(source)}</b></a></figcaption></figure>')
        elif kind == "image_ref":
            o.append(f'<p class="note">Image {block[1]} (shown above).</p>')
        elif kind == "image_note":
            o.append(f'<p class="note">Image {block[1]}: {esc(block[2])} — {esc(block[3] or "")}; {esc(block[4])}.</p>')
    flush()
    if open_part:
        o.append("</div></article>")
    o.append("</div></section>")
    # 04 labels
    o.append('<section class="block" id="labels"><div class="inner"><div class="sec-head">05 &nbsp;Label index</div>'
             '<h2>Index by validation status</h2><p class="lead">Statuses that do not show content to be original are expanded.</p>')
    for status in LABEL_ORDER:
        items = content.label_rows(status)
        if not items:
            continue
        opened = " open" if status != "crc_verified" else ""
        o.append(f'<details{opened}><summary>{pill(status)}<span>{esc(NOTES[status])}</span>'
                 f'<span class="cnt">{len(items)} item(s)</span></summary><ul>')
        for item, anchor, where in items:
            label = f'<a href="#{anchor}">{esc(item)}</a>' if anchor else esc(item)
            o.append(f'<li>{label}<span class="loc">({esc(where)})</span></li>')
        o.append("</ul></details>")
    o.append("</div></section>")
    o.append(f'<footer><div class="inner"><b>ORT {esc(m.get("tool_version"))}</b> · Rebuilt from recovered content; '
             "the exact text and parts are in text/ and parts/ of this folder.</div></footer></body></html>")
    return "\n".join(o)


def write_reading_outputs(entries, manifest, output_directory, make_pdf=True):
    """Write the rebuilt document, its PDF preview, and recovered_report.html; return the manifest fields.

    Each step is independent: a failure is recorded in its field and never affects recovery or the other steps."""
    from . import pdf, rebuild

    output_directory = Path(output_directory)
    fields = {}
    try:
        name = rebuild.write_rebuilt_document(Content(entries, manifest), manifest, output_directory)
        fields["rebuilt_document"] = {"path": name, "error": None}
    except Exception as error:  # noqa: BLE001
        fields["rebuilt_document"] = {"path": None, "error": f"{type(error).__name__}: {error}"}
    rebuilt_path = fields["rebuilt_document"]["path"]
    fields["preview_pdf"] = pdf.make_preview(output_directory / rebuilt_path if rebuilt_path else None,
                                             enabled=make_pdf)
    try:
        content = Content(entries, manifest)
        blocks = content.build()
        (output_directory / HTML_FILENAME).write_text(render_html(content, blocks, manifest, fields), encoding="utf-8")
        fields["recovered_report"] = {"path": HTML_FILENAME, "error": None}
    except Exception as error:  # noqa: BLE001
        fields["recovered_report"] = {"path": None, "error": f"{type(error).__name__}: {error}"}
    return fields
