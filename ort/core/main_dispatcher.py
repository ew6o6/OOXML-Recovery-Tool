import os
import logging
import argparse
from core.handlers.docx import process_extracted_docx_data
from core.handlers.xlsx import process_extracted_xlsx_data
from core.handlers.pptx import process_extracted_pptx_data

def setup_logging():
    """Configure logging format and level"""
    logging.basicConfig(
        level=logging.INFO,
        format='[%(levelname)s] %(message)s'
    )

def process_file(file_path):
    """Process a single OOXML file and delegate to the corresponding handler based on file extension"""
    file_name = os.path.basename(file_path)
    ext = os.path.splitext(file_name)[-1].lower()
    logging.info(f"Input file: {file_name}")

    from core.extractor import get_file_hex
    local_file_in_hex, file_ext = get_file_hex(file_path)
    if not local_file_in_hex:
        logging.warning("Extraction failed or not an OOXML format file.")
        return

    if file_ext == '.docx':
        process_extracted_docx_data(local_file_in_hex, file_path)
    elif file_ext == '.xlsx':
        process_extracted_xlsx_data(local_file_in_hex, file_path)
    elif file_ext == '.pptx':
        process_extracted_pptx_data(local_file_in_hex, file_path)
    else:
        logging.warning("Unsupported file extension.")

def process_directory_or_file(path):
    """Process either a single file or all files in a directory"""
    if os.path.isdir(path):
        for filename in os.listdir(path):
            file_path = os.path.join(path, filename)
            if os.path.isfile(file_path):
                process_file(file_path)
    elif os.path.isfile(path):
        process_file(path)
    else:
        logging.error(f"Invalid path: '{path}'")

def parse_args():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description="OOXML damaged file recovery and extraction tool")
    parser.add_argument("path", help="Path to a .docx, .xlsx, .pptx file or a directory")
    return parser.parse_args()

if __name__ == '__main__':
    setup_logging()
    args = parse_args()
    process_directory_or_file(args.path)
