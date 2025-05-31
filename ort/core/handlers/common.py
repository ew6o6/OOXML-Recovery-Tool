import os
import re
import csv
from bs4 import BeautifulSoup
from docx import Document

def parse_styles_xml(styles_xml):
    """
    Parses the style definitions in a DOCX styles.xml file.
    Example return: {"Heading1": {"name": "heading 1", "based_on": "Normal"}}
    """
    soup = BeautifulSoup(styles_xml, "lxml")
    styles = {}

    for style in soup.find_all("w:style"):
        style_id = style.get("w:styleid")
        name_tag = style.find("w:name")
        based_on_tag = style.find("w:basedon")

        if style_id:
            styles[style_id] = {
                "name": name_tag.get("w:val") if name_tag else "",
                "based_on": based_on_tag.get("w:val") if based_on_tag else None
            }

    return styles

def extract_xml_text_for_docx(xml_data: str, styles: dict) -> str:
    """Extracts paragraphs from DOCX document XML and annotates them with style names."""
    soup = BeautifulSoup(xml_data, "lxml")
    lines = []

    for p in soup.find_all('w:p'):
        style_tag = p.find('w:pstyle')
        style_id = style_tag.get('w:val') if style_tag else ''
        style_info = styles.get(style_id, {})
        style_name = style_info.get('name', '')

        paragraph_text = ' '.join(p.stripped_strings)
        if style_name:
            lines.append(f"[{style_name}] {paragraph_text}")
        else:
            lines.append(paragraph_text)

    return '\n'.join(lines)

def save_text_to_docx(text, filename="output.docx"):
    """Writes text content to a .docx file."""
    doc = Document()
    doc.add_paragraph(text)
    doc.save(filename)

def extract_data_from_shared_strings(xml):
    """Extracts text values from sharedStrings.xml in XLSX files."""
    return re.findall(r'<t>([^<]+)</t>', xml)

def extract_data_from_sheet(sheet_xml, shared_strings, style_map):
    """Extracts structured cell data from an XLSX worksheet XML."""
    soup = BeautifulSoup(sheet_xml, "xml")
    mapped_data = {}
    unmapped_data = set(shared_strings)

    for cell in soup.select("sheetData c"):
        ref = cell.get("r")
        if not ref:
            continue

        row_match = re.search(r"(\d+)", ref)
        col_match = re.search(r"([A-Z]+)", ref)
        if not row_match or not col_match:
            continue

        row = int(row_match.group())
        col = col_match.group()

        dtype = cell.get("t")
        s_idx = cell.get("s")
        fmt = style_map.get(s_idx, "")

        val_tag = cell.find("v")
        val = val_tag.text if val_tag else ""

        is_tag = cell.find("is")
        f_tag = cell.find("f")

        if dtype == "inlineStr" and is_tag:
            t_tag = is_tag.find("t")
            text = t_tag.text if t_tag else ""
        elif f_tag:
            formula = f_tag.text
            text = f"= {formula}"
            if val:
                text += f" → {val}"
        elif dtype == "s" and val.isdigit():
            idx = int(val)
            if idx < len(shared_strings):
                text = shared_strings[idx]
                unmapped_data.discard(text)
            else:
                text = "INDEX_ERROR"
        else:
            text = val

        if fmt:
            text = f"{text} ({fmt})"

        mapped_data.setdefault(row, {})[col] = text

    return mapped_data, unmapped_data

def parse_xlsx_styles(styles_xml):
    """Parses styles.xml from an XLSX file and builds a style mapping dictionary."""
    soup = BeautifulSoup(styles_xml, "xml")
    style_map = {}
    numfmts = {}

    for fmt in soup.find_all("numFmt"):
        numfmts[fmt.get("numFmtId")] = fmt.get("formatCode")

    for i, xf in enumerate(soup.find_all("xf")):
        numfmt_id = xf.get("numFmtId")
        if numfmt_id in numfmts:
            style_map[str(i)] = numfmts[numfmt_id]
        elif numfmt_id:
            built_in_formats = {
                '14': 'mm-dd-yy',
                '22': 'm/d/yy h:mm',
                '165': 'yyyy/mm/dd',
                '44': '₩#,##0'
            }
            style_map[str(i)] = built_in_formats.get(numfmt_id, '')

    return style_map

def parse_rels_file(rels_xml):
    """Parses a .rels file and returns relationship entries as strings."""
    soup = BeautifulSoup(rels_xml, "xml")
    rels_info = []
    for rel in soup.find_all("Relationship"):
        r_id = rel.get("Id")
        target = rel.get("Target")
        r_type = rel.get("Type")
        rels_info.append(f"{r_id}: {target} ({r_type})")
    return rels_info

def parse_xlsx_rels_file(rels_xml):
    """Alias for parse_rels_file specific to XLSX context."""
    return parse_rels_file(rels_xml)

def save_unmapped_to_csv(unmapped, out_dir, filename="unmapped_data.csv"):
    """Saves a list of unmapped values to a CSV file."""
    out_path = os.path.join(out_dir, filename)
    with open(out_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["Unmapped Data"])
        writer.writerows([[val] for val in unmapped])

def display_and_save_table_to_csv(mapped_data, filename, out_dir):
    """Displays and saves structured table data to a CSV file."""
    if not mapped_data:
        return

    rows = sorted(mapped_data.keys())
    cols = sorted({c for r in mapped_data.values() for c in r})

    table = [[row] + [mapped_data[row].get(col, '') for col in cols] for row in rows]
    headers = ["Row"] + cols

    with open(os.path.join(out_dir, filename), 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(table)
