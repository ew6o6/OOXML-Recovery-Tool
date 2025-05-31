import os
from bs4 import BeautifulSoup
from docx import Document
from .common import parse_styles_xml, parse_rels_file, extract_xml_text_for_docx, save_text_to_docx


def process_extracted_docx_data(local_file_xml, file_path):
    out_dir = os.path.join(os.path.dirname(file_path), f"output_{os.path.basename(file_path).rsplit('.', 1)[0]}")
    os.makedirs(out_dir, exist_ok=True)

    doc_xml = ''
    styles_xml = ''
    rels_xml = ''

    for item in local_file_xml:
        if item['local_file_name'] == 'word/document.xml':
            doc_xml = item['local_file_data']
        elif item['local_file_name'] == 'word/styles.xml':
            styles_xml = item['local_file_data']
        elif item['local_file_name'] == 'word/_rels/document.xml.rels':
            rels_xml = item['local_file_data']

    if doc_xml:
        styles = parse_styles_xml(styles_xml) if styles_xml else {}
        text = extract_xml_text_for_docx(doc_xml, styles)

        if rels_xml:
            rels = parse_rels_file(rels_xml)
            text += '\n\n[Relationships]\n' + '\n'.join(rels)

        save_text_to_docx(text, os.path.join(out_dir, "output_damaged.docx"))
        print(f"\n[DOCX output path] {out_dir}")
