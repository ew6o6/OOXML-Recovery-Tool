"""Shared XML-to-text and XML-to-table helpers for the format handlers."""

import posixpath
import re

from bs4 import BeautifulSoup
from docx import Document

from ..utils import local_name, parse_xml


def extract_xml_text_for_docx(xml_data: str) -> str:
    """Extract DOCX text without inventing spaces at Word run boundaries."""
    soup = BeautifulSoup(xml_data, "xml")
    lines = []
    for p in soup.find_all('w:p'):
        chunks = []
        for node in p.descendants:
            if getattr(node, 'name', None) == 't':
                chunks.append(node.get_text())
            elif getattr(node, 'name', None) in ('tab', 'br', 'cr'):
                chunks.append('\t' if node.name == 'tab' else '\n')
        lines.append(''.join(chunks))
    return '\n'.join(lines)


_XML_INVALID_CHARS = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


def save_text_to_docx(paragraphs, filename):
    """Write paragraphs to a plain .docx file (characters that XML cannot store become U+FFFD)."""
    doc = Document()
    for paragraph in paragraphs:
        doc.add_paragraph(_XML_INVALID_CHARS.sub("�", paragraph))
    doc.save(filename)


# ----------------------------------------------------------------------------- XLSX

# ECMA-376 Part 1, 18.8.30: built-in number formats (locale-independent ones).
BUILTIN_NUMBER_FORMATS = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00",
    12: "# ?/?", 13: "# ??/??", 14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss",
    46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}


def _rich_text(node):
    """Text of a string item: every <t> in order, except phonetic runs (<rPh>), whitespace kept."""
    pieces = []
    for element in node.iter():
        if local_name(element) == "t" and not any(local_name(parent) == "rPh" for parent in element.iterancestors()):
            pieces.append(element.text or "")
    return "".join(pieces)


def parse_shared_strings(xml):
    """Shared strings in index order: one entry per <si>, rich-text runs concatenated."""
    root = parse_xml(xml)
    if root is None:
        return []
    return [_rich_text(item) for item in root.iter() if local_name(item) == "si"]


def column_number(letters):
    number = 0
    for char in letters.upper():
        number = number * 26 + ord(char) - 64
    return number


def column_letters(number):
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def parse_sheet_cells(xml):
    """Cells of a worksheet in document order. Missing cell references are inferred from position."""
    root = parse_xml(xml)
    cells = []
    if root is None:
        return cells
    row_number = 0
    for row in root.iter():
        if local_name(row) != "row":
            continue
        declared_row = row.get("r")
        row_number = int(declared_row) if declared_row and declared_row.isdigit() else row_number + 1
        column = 0
        for cell in row:
            if local_name(cell) != "c":
                continue
            match = re.fullmatch(r"([A-Za-z]{1,3})(\d+)", cell.get("r") or "")
            if match:
                column, number = column_number(match.group(1)), int(match.group(2))
            else:
                column, number = column + 1, row_number
            value = formula = inline = None
            for child in cell:
                name = local_name(child)
                if name == "v":
                    value = child.text or ""
                elif name == "f":
                    formula = child.text or (f"[shared formula si={child.get('si')}]" if child.get("si") else "")
                elif name == "is":
                    inline = _rich_text(child)
            cells.append({"ref": f"{column_letters(column)}{number}", "row": number, "column": column,
                          "type": cell.get("t") or "n", "raw": value, "formula": formula, "inline": inline,
                          "style": cell.get("s")})
    return cells


def resolve_cell(cell, shared_strings):
    """(display value, shared-string index or None) for a parsed cell."""
    kind, raw = cell["type"], cell["raw"]
    if kind == "s":
        if raw is not None and raw.strip().isdigit():
            index = int(raw)
            if index < len(shared_strings):
                return shared_strings[index], index
            return f"[shared string #{index} not recovered]", index
        return raw or "", None
    if kind == "inlineStr":
        return cell["inline"] or "", None
    if kind == "b":
        return {"1": "TRUE", "0": "FALSE"}.get(raw, raw or ""), None
    return raw or "", None


def parse_number_formats(styles_xml):
    """Number format code per cell style index (cellXfs only; cellStyleXfs are not cell styles)."""
    root = parse_xml(styles_xml)
    if root is None:
        return {}
    custom = {}
    cell_xfs = None
    for element in root.iter():
        name = local_name(element)
        if name == "numFmt" and (element.get("numFmtId") or "").isdigit():
            custom[int(element.get("numFmtId"))] = element.get("formatCode")
        elif name == "cellXfs" and cell_xfs is None:
            cell_xfs = element
    formats = {}
    for index, xf in enumerate(child for child in (cell_xfs if cell_xfs is not None else ()) if local_name(child) == "xf"):
        identifier = xf.get("numFmtId") or "0"
        if identifier.isdigit():
            number = int(identifier)
            formats[str(index)] = custom.get(number) or BUILTIN_NUMBER_FORMATS.get(number) or f"numFmtId={number}"
    return formats


def sheet_display_names(workbook_xml, workbook_rels_xml):
    """Map worksheet part names to their display names via xl/workbook.xml and its relationships."""
    workbook, relationships = parse_xml(workbook_xml), parse_xml(workbook_rels_xml)
    if workbook is None or relationships is None:
        return {}
    targets = {}
    for relationship in relationships.iter():
        if local_name(relationship) == "Relationship" and relationship.get("Target"):
            target = relationship.get("Target")
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            targets[relationship.get("Id")] = part
    names = {}
    for sheet in workbook.iter():
        if local_name(sheet) != "sheet":
            continue
        relationship_id = next((value for key, value in sheet.attrib.items()
                                if key.endswith("}id") and "relationships" in key), None)
        if relationship_id in targets and sheet.get("name") is not None:
            names[targets[relationship_id]] = sheet.get("name")
    return names
