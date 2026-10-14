"""Command-line interface of the OOXML Recovery Toolkit (ORT).

Exit status: 0 when every input is complete, 1 when at least one is partial, and 2 when at
least one failed, is indeterminate, is not OOXML, or the path holds no file.
"""

import argparse
import logging
import os

from . import __version__
from .core.decoder import MAX_DECODED_BYTES
from .core.extractor import get_file_hex

MIB = 1024 * 1024
_EXIT_CODES = {"complete": 0, "partial": 1}


def exit_code(status):
    return _EXIT_CODES.get(status, 2)


def setup_logging():
    logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')


def _positive_int(value):
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return number


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        prog="ort",
        description="Recover document content and images from damaged OOXML files (.docx, .xlsx, .pptx).",
        epilog="Exit status: 0 = every file complete, 1 = at least one partial, "
               "2 = at least one failed, indeterminate, or not OOXML.",
    )
    parser.add_argument("path", help="A file, or a folder whose files are processed")
    parser.add_argument("--output-root", help="Write each output folder below this directory (default: next to the input)")
    parser.add_argument("-r", "--recursive", action="store_true",
                        help="Also process files in subfolders (folders named output_* are skipped)")
    parser.add_argument("--salvage-fragments", action="store_true",
                        help="Experimental: save independently decodable XML suffix fragments separately")
    parser.add_argument("--dump-entries", action="store_true",
                        help="Also write entries.json with every entry, including payload bytes as hex")
    parser.add_argument("--max-part-size", type=_positive_int, default=MAX_DECODED_BYTES // MIB, metavar="MIB",
                        help="Largest decoded size of one part in MiB (default: %(default)s)")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser.parse_args(argv)


def iter_input_files(path, recursive=False):
    """Files to process: the path itself, the files of a folder, or (recursive) of its subfolders."""
    if os.path.isfile(path):
        yield path
        return
    if not recursive:
        for name in sorted(os.listdir(path)):
            if os.path.isfile(os.path.join(path, name)):
                yield os.path.join(path, name)
        return
    for root, folders, files in os.walk(path):
        folders[:] = sorted(folder for folder in folders if not folder.startswith("output_"))
        for name in sorted(files):
            yield os.path.join(root, name)


def process_file(file_path, *, output_root=None, salvage_fragments=False, dump_entries=False,
                 max_part_size=MAX_DECODED_BYTES):
    """Recover one file and return its RecoveryResult."""
    logging.info("Input file: %s", file_path)
    result = get_file_hex(file_path, output_root=output_root, salvage_fragments=salvage_fragments,
                          dump_entries=dump_entries, max_part_size=max_part_size)
    if result.output_directory is None:
        logging.warning("Skipped (%s): %s", result.overall_status, file_path)
    return result


def main(argv=None):
    """Run the CLI and return the exit status."""
    setup_logging()
    args = parse_args(argv)
    if not os.path.exists(args.path):
        logging.error("Invalid path: '%s'", args.path)
        return 2
    logging.info("ORT %s", __version__)
    worst = None
    for file_path in iter_input_files(args.path, args.recursive):
        result = process_file(file_path, output_root=args.output_root, salvage_fragments=args.salvage_fragments,
                              dump_entries=args.dump_entries, max_part_size=args.max_part_size * MIB)
        worst = max(worst or 0, exit_code(result.overall_status))
    if worst is None:
        logging.error("No file found in '%s'", args.path)
        return 2
    return worst
