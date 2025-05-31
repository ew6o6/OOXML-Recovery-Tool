import os
import re
import json
import binascii
import zipfile
import tempfile
from .decoder import decode_local_file_data, decode_utf8
from .utils import extract_file_name, extract_img_file, extract_metadata
from .handlers import docx, xlsx, pptx
from bs4 import XMLParsedAsHTMLWarning
import warnings
import xml.etree.ElementTree as ET

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

def extract_embedded_ooxml_if_needed(file_path):
    if not zipfile.is_zipfile(file_path):
        return file_path

    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        embedded = [f for f in zip_ref.namelist() if f.endswith(('.pptx', '.docx', '.xlsx'))]
        if not embedded:
            return file_path
        target = embedded[0]
        temp_dir = tempfile.mkdtemp()
        extracted_path = os.path.join(temp_dir, os.path.basename(target))
        with open(extracted_path, 'wb') as out_f:
            out_f.write(zip_ref.read(target))
        return extracted_path

def is_structurally_valid_ooxml(file_path):
    required_entries = {
        '.docx': ['[Content_Types].xml', '_rels/.rels', 'word/document.xml'],
        '.xlsx': ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml'],
        '.pptx': ['[Content_Types].xml', '_rels/.rels', 'ppt/slides/slide1.xml'],
    }

    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            names = z.namelist()
            for prefix, ext in [('word/', '.docx'), ('xl/', '.xlsx'), ('ppt/', '.pptx')]:
                if any(name.startswith(prefix) for name in names):
                    if all(req in names for req in required_entries[ext]):
                        return True, ext
                    else:
                        return False, ext
    except zipfile.BadZipFile:
        return False, None
    return False, None

def has_meaningful_content(xml_content):
    try:
        root = ET.fromstring(xml_content)
        return any("t" in elem.tag and elem.text and elem.text.strip() for elem in root.iter())
    except ET.ParseError:
        return False

def get_file_hex(file_path):
    file_path = extract_embedded_ooxml_if_needed(file_path)

    if not os.path.exists(file_path):
        return None, None

    with open(file_path, 'rb') as f:
        data = f.read()

    pk_matches = list(re.finditer(rb'\x50\x4B\x03\x04', data))
    if not pk_matches:
        print("[ERROR] No valid PK header found. Not an OOXML-based ZIP.")
        return None, None

    positions = [m.start() for m in pk_matches]
    parts = [data[start:end] for start, end in zip([0] + positions, positions + [None]) if data[start:end]]

    local_file_in_hex = []
    for part in parts:
        if len(part) < 30:
            continue

        file_name_len = int.from_bytes(part[26:28], 'little')
        extra_len = int.from_bytes(part[28:30], 'little')

        file_name_start = 30
        file_name_end = file_name_start + file_name_len
        file_data_start = file_name_end + extra_len

        file_name_bytes = part[file_name_start:file_name_end]
        file_data = part[file_data_start:]

        file_name_str = extract_file_name(decode_utf8(binascii.hexlify(file_name_bytes).decode()))
        ext_in_name = os.path.splitext(file_name_str)[1].lower()
        img_ext = ext_in_name if ext_in_name in ['.jpg', '.jpeg', '.png', '.bmp', '.gif'] else None

        local_file_in_hex.append({
            'local_file_name': file_name_str,
            'local_file_data': binascii.hexlify(file_data).decode(),
            'img_ext': img_ext
        })

    decode_local_file_data(local_file_in_hex)

    base_directory = os.path.dirname(file_path)
    file_name = os.path.basename(file_path)
    file_base_name = os.path.splitext(file_name)[0]
    output_directory = os.path.join(base_directory, f'output_{file_base_name}')
    os.makedirs(output_directory, exist_ok=True)

    json_path = os.path.join(output_directory, f'{file_base_name}_localFiles.json')
    i = 1
    while os.path.exists(json_path):
        json_path = os.path.join(output_directory, f'{file_base_name}({i}).json')
        i += 1

    img_dir = os.path.join(output_directory, 'img')
    os.makedirs(img_dir, exist_ok=True)

    with open(json_path, 'w', encoding='utf-8-sig') as f:
        json.dump(local_file_in_hex, f, ensure_ascii=False, indent=2)

    extract_img_file(local_file_in_hex, img_dir)
    extract_metadata(local_file_in_hex, output_directory)

    print(f"[INFO] Metadata saved to {output_directory}/metadata/metadata.txt")

    file_ext = None
    for entry in local_file_in_hex:
        path = entry['local_file_name']
        if path.startswith('word/'):
            file_ext = '.docx'
        elif path.startswith('xl/'):
            file_ext = '.xlsx'
        elif path.startswith('ppt/'):
            file_ext = '.pptx'

    print(f"Actual file extension: {file_ext[1:] if file_ext else 'unknown'}")

    if file_ext == '.docx':
        docx.process_extracted_docx_data(local_file_in_hex, file_path)
    elif file_ext == '.xlsx':
        xlsx.process_extracted_xlsx_data(local_file_in_hex, file_path)
    elif file_ext == '.pptx':
        pptx.process_extracted_pptx_data(local_file_in_hex, file_path)

    if file_ext:
        print(f"[INFO] Completed processing for {file_ext[1:]} file.")
    else:
        print("[INFO] Processing completed.")
    return local_file_in_hex, file_ext
