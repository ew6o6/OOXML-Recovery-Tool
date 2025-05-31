"""
main.py - Refactored standalone execution script (based on the original ooxml_code.py)
role: Command-line interface for processing OOXML files

author: Jiyoon Kim
last modified date: 2025-05-11
description: Delegates get_file_hex() call to core.extractor module, establishing a modular architecture
version: 1.0.0
history:
    2025-05-06: Initial version
    2025-05-07: Introduced logging and argparse
    2025-05-08: Added user input handling
    2025-05-09: Added path validation
    2025-05-10: Added exception handling
    2025-05-11: Code optimization and comment refinement
"""

import os
import logging
import argparse
from .core.extractor import get_file_hex

def setup_logging():
    """Configure log format and level"""
    logging.basicConfig(
        level=logging.INFO,
        format='[%(levelname)s] %(message)s'
    )

def parse_args():
    """Define command-line arguments"""
    parser = argparse.ArgumentParser(description="OOXML damaged file recovery and extraction tool")
    parser.add_argument("path", help="Path to a .docx, .xlsx file or a directory containing such files")
    return parser.parse_args()

def process_file(file_path):
    """Process an individual file"""
    file_name = os.path.basename(file_path)
    logging.info(f"Input file: {file_name}")
    local_file_in_hex, file_ext = get_file_hex(file_path)
    if local_file_in_hex:
        logging.info("Processing completed.\n")
    return local_file_in_hex

def process_directory_or_file(path):
    """Branch logic to handle either a single file or all files in a directory"""
    if os.path.isdir(path):
        for filename in os.listdir(path):
            file_path = os.path.join(path, filename)
            if os.path.isfile(file_path):
                process_file(file_path)
    elif os.path.isfile(path):
        process_file(path)
    else:
        logging.error(f"Invalid path: '{path}'")

def main():
    """Main entry point"""
    print("Starting OOXML damaged file recovery and extraction tool")
    setup_logging()
    args = parse_args()
    process_directory_or_file(args.path)

if __name__ == '__main__':
    main()
