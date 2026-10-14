"""Compatibility wrapper around the package CLI."""

from ..main import parse_args, process_directory_or_file, process_file, setup_logging


if __name__ == "__main__":
    setup_logging()
    args = parse_args()
    process_directory_or_file(
        args.path,
        output_root=args.output_root,
        salvage_fragments=args.salvage_fragments,
    )
