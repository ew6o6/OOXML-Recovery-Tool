"""
core/utils.py - Utility functions for file name handling, image saving, and metadata extraction
role: Handles local file components extracted from OOXML files

author: Jiyoon Kim
date: 2025-05-06

description:
    extract_file_name() - Removes extra fields from local file names
    extract_img_file() - Extracts and saves images from the media/ directory
    extract_metadata() - Extracts document metadata from core.xml and writes to a text file
"""
import os
import re
import binascii

def extract_file_name(decoded_name):
    pattern = re.search(r'[\x00-\x1F\x7F]', decoded_name[::-1])
    if pattern:
        decoded_name = decoded_name[-pattern.end()+1:]
    return decoded_name

# def extract_img_file(local_file, output_path):
#     os.makedirs(output_path, exist_ok=True)
#     for item in local_file:
#         name = item.get('local_file_name', '')
#         if name.startswith(('word/media/', 'xl/media/', 'ppt/media/')):
#             img_name = os.path.basename(name)
#             img_path = os.path.join(output_path, img_name)
#             try:
#                 raw = item.get('decoded_data')
#                 if not raw:
#                     raw = binascii.unhexlify(item['local_file_data'])
#                 elif isinstance(raw, str):
#                     raw = raw.encode()  # Fallback in case
#                 with open(img_path, 'wb') as f:
#                     f.write(raw)
#             except Exception as e:
#                 print(f"[ERROR] Failed to save image: {img_name} ({e})")



def extract_img_file(local_file, output_path):
    os.makedirs(output_path, exist_ok=True)
    image_counter = 1

    for item in local_file:
        name = item.get('local_file_name', '')
        hex_data = item.get('local_file_data', '')

        if 'media/' not in name or not hex_data:
            continue

        # 확장자 추정 (기본 jpg, 필요시 서명 기반 개선 가능)
        ext = '.jpg'
        if hex_data.startswith('89504e47'):  # PNG signature
            ext = '.png'
        elif hex_data.startswith('47494638'):  # GIF signature
            ext = '.gif'
        elif hex_data.startswith('424d'):  # BMP signature
            ext = '.bmp'

        img_name = f"img{image_counter}{ext}"
        img_path = os.path.join(output_path, img_name)

        try:
            with open(img_path, 'wb') as f:
                f.write(binascii.unhexlify(hex_data))
            image_counter += 1
        except Exception as e:
            print(f"[ERROR] Failed to save image: {img_name} ({e})")





def extract_metadata(local_file, output_dir):
    tags = {
        'dc:creator': 'creator',
        'cp:lastModifiedBy': 'lastModifiedBy',
        'cp:revision': 'revision',
        'dcterms:created': 'created',
        'dcterms:modified': 'modified'
    }
    results = ["[Document Metadata]"]

    for item in local_file:
        if item['local_file_name'] == 'docProps/core.xml':
            try:
                xml_raw = item.get('decoded_data')
                if not xml_raw:
                    xml_raw = binascii.unhexlify(item['local_file_data']).decode('utf-8', errors='ignore')
                elif isinstance(xml_raw, bytes):
                    xml_raw = xml_raw.decode('utf-8', errors='ignore')

                all_missing = True
                for tag, name in tags.items():
                    match = re.search(f"<{tag}.*?>(.*?)</{tag}>", xml_raw)
                    if match and match.group(1):
                        all_missing = False
                        value = match.group(1)
                        results.append(f"{name} : {value} (<{tag}>{value}</{tag}>)")

                if all_missing:
                    results = ["Metadata is severely damaged or missing."]

            except Exception as e:
                results = [f"[ERROR] Metadata extraction failed: {e}"]

    metadata_path = os.path.join(output_dir, 'metadata', 'metadata.txt')
    os.makedirs(os.path.dirname(metadata_path), exist_ok=True)
    with open(metadata_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(results))
    print(f"[INFO] Metadata saved to {metadata_path}")
    return metadata_path
