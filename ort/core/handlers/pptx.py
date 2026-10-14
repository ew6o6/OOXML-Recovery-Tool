import os
import re
import binascii
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw, ImageFont
from .common import parse_rels_file
from ..utils import extract_file_name, safe_output_stem, safe_print

def sanitize_filename(name):
    """Remove or replace invalid characters from filename."""
    if not name:
        return "unknown"
    # Keep only ASCII printable characters (32-126), replace others
    result = []
    for c in name:
        if 32 <= ord(c) <= 126:
            result.append(c)
    name = ''.join(result)
    # Remove invalid Windows filename characters
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    # Limit length
    if len(name) > 100:
        name = name[:100]
    return name.strip() if name.strip() else "unknown"

def process_extracted_pptx_data(local_file_xml, file_path, output_directory=None):
    out_dir = output_directory or os.path.join(os.path.dirname(file_path), f"output_{safe_output_stem(file_path)}")
    os.makedirs(out_dir, exist_ok=True)

    # 슬라이드 텍스트 추출
    # Only validated decoded XML is parsed; compressed payload hex is never treated as XML.
    slides = [(i['local_file_name'], i['decoded_xml']) for i in local_file_xml
              if re.match(r'ppt/slides/slide\d+\.xml$', i['local_file_name']) and i.get('decoded_xml')]
    for name, xml in slides:
        soup = BeautifulSoup(xml, "xml")
        slide_num = sanitize_filename(name.split("/")[-1].replace(".xml", ""))
        texts = []

        for sp in soup.find_all("p:sp"):
            placeholder = sp.find("p:ph")
            shape_type = placeholder.get("type") if placeholder else "body"
            content_type = "title" if shape_type == "title" else "body"

            text_elements = sp.find_all("a:t")
            shape_text = " ".join(t.get_text() for t in text_elements)

            if shape_text.strip():
                texts.append(f"[{content_type.upper()}] {shape_text.strip()}")

        # 텍스트 파일로 저장
        out_path = os.path.join(out_dir, f"{slide_num}.txt")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(texts))
            
    # 노트 슬라이드 텍스트 추출 (document_content_status 범위, Q7)
    notes = [(i['local_file_name'], i['decoded_xml']) for i in local_file_xml
             if re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', i['local_file_name']) and i.get('decoded_xml')]
    for name, xml in notes:
        soup = BeautifulSoup(xml, "xml")
        paragraphs = []
        for paragraph in soup.find_all("a:p"):
            text = "".join(t.get_text() for t in paragraph.find_all("a:t"))
            if text.strip():
                paragraphs.append(text)
        note_name = sanitize_filename(name.split("/")[-1].replace(".xml", ""))
        with open(os.path.join(out_dir, f"{note_name}.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(paragraphs))

    # 관계 파일 추출
    rels_xml = next((i.get('decoded_xml', '') for i in local_file_xml if i['local_file_name'] == 'ppt/_rels/presentation.xml.rels'), "")
    if rels_xml:
        rels = parse_rels_file(rels_xml)
        with open(os.path.join(out_dir, "relationships_pptx.txt"), 'w', encoding='utf-8') as f:
            f.write("[Relationships]\n" + "\n".join(rels))

    # 메타데이터 추출
    metadata_path = os.path.join(out_dir, 'metadata', 'metadata.txt')
    os.makedirs(os.path.dirname(metadata_path), exist_ok=True)
    core_xml = next((i.get('decoded_xml') for i in local_file_xml if i['local_file_name'] == 'docProps/core.xml'), None)
    if core_xml:
        tags = {
            'dc:creator': 'creator',
            'cp:lastModifiedBy': 'lastModifiedBy',
            'cp:revision': 'revision',
            'dcterms:created': 'created',
            'dcterms:modified': 'modified'
        }
        results = ["[Document Metadata]"]
        for tag, label in tags.items():
            match = BeautifulSoup(core_xml, "xml").find(tag)
            if match:
                results.append(f"{label} : {match.text} (<{tag}>{match.text}</{tag}>)")
        with open(metadata_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(results))

    safe_print(f"\n[PPTX output path] {out_dir}")
