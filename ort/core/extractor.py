import os
import re
import json
import binascii
import struct
from .decoder import decode_local_file_data, decode_utf8
from .utils import extract_file_name, extract_img_file, extract_metadata
from .handlers import docx, xlsx, pptx

# PK ZIP signatures
LFH_SIG = b'\x50\x4B\x03\x04'  # Local file header
CD_SIG = b'\x50\x4B\x01\x02'   # Central directory header
EOCD_SIG = b'\x50\x4B\x05\x06' # End of central directory

def is_structurally_valid_ooxml_raw(file_path):
    """
    CD-independent check for OOXML structure.
    Parses local file headers directly without relying on Central Directory.
    """
    try:
        with open(file_path, 'rb') as f:
            data = f.read()

        # Find all local file headers
        local_file_names = extract_local_file_names_raw(data)

        if not local_file_names:
            return False, None

        # Check for OOXML indicators
        for prefix, ext in [('word/', '.docx'), ('xl/', '.xlsx'), ('ppt/', '.pptx')]:
            if any(name.startswith(prefix) for name in local_file_names):
                return True, ext

        return False, None
    except:
        return False, None

def extract_local_file_names_raw(data):
    """
    Extract local file names directly from binary data without using Central Directory.
    Based on PK ZIP local file header structure.
    """
    names = []
    pos = 0

    while pos < len(data) - 30:
        # Look for local file header signature
        idx = data.find(LFH_SIG, pos)
        if idx == -1:
            break

        # Parse local file header
        try:
            if idx + 30 > len(data):
                break

            # Get filename length and extra field length from header
            fn_len = struct.unpack_from('<H', data, idx + 26)[0]
            ex_len = struct.unpack_from('<H', data, idx + 28)[0]

            # Sanity check on lengths
            if fn_len == 0 or fn_len > 512 or ex_len > 65535:
                pos = idx + 4
                continue

            # Extract filename
            fn_start = idx + 30
            fn_end = fn_start + fn_len

            if fn_end > len(data):
                break

            filename = data[fn_start:fn_end].decode('utf-8', errors='ignore')
            if filename and not filename.startswith('\x00'):
                names.append(filename)

            pos = fn_end + ex_len
        except (struct.error, UnicodeDecodeError):
            pos = idx + 4
            continue

    return names

def get_file_hex(file_path):
    """
    Main extraction function implementing the paper's 3-stage process:
    Stage 1 (Preprocessing): PK signature check, Deflate verification, OOXML type inference
    Stage 2 (Damage Identification): Core Part presence check
    Stage 3 (Selective Extraction): Deflate decompression, XML parsing, output generation
    """
    if not os.path.exists(file_path):
        return None, None

    with open(file_path, 'rb') as f:
        data = f.read()

    # === Stage 1: Preprocessing ===
    # 1-1. PK signature check
    if data[:4] != LFH_SIG:
        print("[PREPROCESSING] No PK signature at file start. Not a ZIP-based file.")
        return None, None

    # 1-2. Compression method check (Deflate = 0x08)
    if len(data) >= 10:
        compression_method = struct.unpack_from('<H', data, 8)[0]
        if compression_method not in (0, 8):  # 0=stored, 8=deflate
            print(f"[PREPROCESSING] Unexpected compression method: {compression_method}. Not a valid OOXML file.")
            return None, None

    # 1-3. OOXML type inference via Core Part detection
    is_valid, inferred_ext = is_structurally_valid_ooxml_raw(file_path)
    if not is_valid:
        print("[PREPROCESSING] No OOXML core parts found. Filtering as non-OOXML file (Type ③).")
        return None, None

    print(f"[PREPROCESSING] Valid OOXML structure detected. Inferred type: {inferred_ext}")

    # === Stage 2: Damage Identification ===
    local_file_names = extract_local_file_names_raw(data)
    core_content_parts = {
        '.docx': ['word/document.xml'],
        '.xlsx': [r'xl/worksheets/sheet\d+\.xml'],
        '.pptx': [r'ppt/slides/slide\d+\.xml'],
    }

    has_content_part = False
    if inferred_ext in core_content_parts:
        for pattern in core_content_parts[inferred_ext]:
            if any(re.match(pattern, name) for name in local_file_names):
                has_content_part = True
                break

    if not has_content_part:
        print(f"[DAMAGE ID] No core content parts found for {inferred_ext}. File may be empty (Type ④) or unrecoverable.")
        return None, None

    print(f"[DAMAGE ID] Core content parts detected. Proceeding to extraction.")

    # === Stage 3: Selective Extraction ===
    pk_matches = list(re.finditer(rb'\x50\x4B\x03\x04', data))
    if not pk_matches:
        return None, None

    positions = [m.start() for m in pk_matches]
    parts = [data[start:end] for start, end in zip([0] + positions, positions + [None]) if data[start:end]]

    local_file_in_hex = []
    for part in parts:
        if len(part) < 30:
            continue

        file_name_len = int.from_bytes(part[26:28], 'little')
        extra_len = int.from_bytes(part[28:30], 'little')

        # Validate filename length - typical OOXML paths are under 200 chars
        # If corrupted (e.g. 0xFFFF), try to find the real filename
        if file_name_len > 500 or file_name_len == 0:
            # Try to extract filename by looking for null terminator or known patterns
            raw_data = part[30:min(30+300, len(part))]
            # Find first null byte or non-printable sequence
            real_len = 0
            for i, b in enumerate(raw_data):
                if b == 0 or b > 127:
                    real_len = i
                    break
            if real_len > 0:
                file_name_len = real_len
            else:
                continue  # Skip this entry entirely

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
