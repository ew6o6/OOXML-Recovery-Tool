import os
import binascii
from bs4 import BeautifulSoup
from PIL import Image, ImageDraw, ImageFont
from .common import parse_rels_file
from ..utils import extract_file_name

def process_extracted_pptx_data(local_file_xml, file_path):
    out_dir = os.path.join(os.path.dirname(file_path), f"output_{os.path.splitext(os.path.basename(file_path))[0]}")
    os.makedirs(out_dir, exist_ok=True)

    # 슬라이드 텍스트 추출
    slides = [(i['local_file_name'], i['local_file_data']) for i in local_file_xml if i['local_file_name'].startswith('ppt/slides/slide')]
    for name, xml in slides:
        soup = BeautifulSoup(xml, "xml")
        slide_num = name.split("/")[-1].replace(".xml", "")
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
            
    # 관계 파일 추출
    rels_xml = next((i['local_file_data'] for i in local_file_xml if i['local_file_name'] == 'ppt/_rels/presentation.xml.rels'), "")
    if rels_xml:
        rels = parse_rels_file(rels_xml)
        with open(os.path.join(out_dir, "relationships_pptx.txt"), 'w', encoding='utf-8') as f:
            f.write("[Relationships]\n" + "\n".join(rels))

    # 이미지 추출
    img_dir = os.path.join(out_dir, 'img')
    os.makedirs(img_dir, exist_ok=True)
    for item in local_file_xml:
        if item['local_file_name'].startswith('ppt/media/'):
            img_ext = os.path.splitext(item['local_file_name'])[1]
            img_path = os.path.join(img_dir, os.path.basename(item['local_file_name']))
            try:
                with open(img_path, 'wb') as f:
                    f.write(binascii.unhexlify(item['local_file_data']))
            except Exception as e:
                print(f"[ERROR] Failed to save image {img_path}: {e}")

    # 메타데이터 추출
    metadata_path = os.path.join(out_dir, 'metadata', 'metadata.txt')
    os.makedirs(os.path.dirname(metadata_path), exist_ok=True)
    core_xml = next((i['local_file_data'] for i in local_file_xml if i['local_file_name'] == 'docProps/core.xml'), None)
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

    print(f"\n[PPTX output path] {out_dir}")
