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
    parser.add_argument("--output-root", help="Write each recovery folder below this directory")
    parser.add_argument(
        "--salvage-fragments",
        action="store_true",
        help="Experimental: save independently decodable XML suffix fragments separately",
    )
    return parser.parse_args()

def process_file(file_path, output_root=None, salvage_fragments=False):
    """Process an individual file"""
    file_name = os.path.basename(file_path)
    logging.info(f"Input file: {file_name}")
    local_file_in_hex, file_ext = get_file_hex(
        file_path, output_root=output_root, salvage_fragments=salvage_fragments
    )
    if local_file_in_hex:
        logging.info("Processing completed.\n")
    return local_file_in_hex

def process_directory_or_file(path, output_root=None, salvage_fragments=False):
    """Branch logic to handle either a single file or all files in a directory"""
    if os.path.isdir(path):
        for filename in os.listdir(path):
            file_path = os.path.join(path, filename)
            if os.path.isfile(file_path):
                process_file(file_path, output_root=output_root, salvage_fragments=salvage_fragments)
    elif os.path.isfile(path):
        process_file(path, output_root=output_root, salvage_fragments=salvage_fragments)
    else:
        logging.error(f"Invalid path: '{path}'")

def main():
    """Main entry point"""
    print("Starting OOXML damaged file recovery and extraction tool")
    setup_logging()
    args = parse_args()
    process_directory_or_file(
        args.path,
        output_root=args.output_root,
        salvage_fragments=args.salvage_fragments,
    )

if __name__ == '__main__':
    main()
