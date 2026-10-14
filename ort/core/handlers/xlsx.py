"""XLSX cell values: one grid CSV and one per-cell CSV per recovered worksheet."""

import csv
import re
from pathlib import Path

from .common import parse_number_formats, parse_sheet_cells, parse_shared_strings, resolve_cell, column_letters, \
    sheet_display_names
from ..utils import first_with_content, safe_print, source_line


SHEET_PART = re.compile(r"^xl/worksheets/sheet(\d+)\.xml$")
CELL_COLUMNS = ["cell", "type", "raw_value", "value", "formula", "number_format"]


def _write_csv(path, header_lines, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as handle:
        for line in header_lines:
            handle.write(line + "\n")
        csv.writer(handle).writerows(rows)


def process_extracted_xlsx_data(entries, output_directory):
    """Write text/sheetN.csv (values on a grid) and text/sheetN_cells.csv (one row per cell).

    N is the worksheet part number, not the order of discovery. Values never carry the number
    format; it is a separate column of the per-cell file. Returns one record per worksheet."""
    out_dir = Path(output_directory)
    text_dir = out_dir / "text"
    shared_item = first_with_content(entries, "xl/sharedStrings.xml")
    shared_strings = parse_shared_strings(shared_item["decoded_xml"]) if shared_item else []
    styles_item = first_with_content(entries, "xl/styles.xml")
    formats = parse_number_formats(styles_item["decoded_xml"]) if styles_item else {}
    workbook, workbook_rels = (first_with_content(entries, name) for name in
                               ("xl/workbook.xml", "xl/_rels/workbook.xml.rels"))
    display_names = (sheet_display_names(workbook["decoded_xml"], workbook_rels["decoded_xml"])
                     if workbook and workbook_rels else {})

    sheets = {}
    for item in entries:
        match = SHEET_PART.match(item.get("local_file_name", ""))
        if match and item.get("decoded_xml") and item["local_file_name"] not in sheets:
            sheets[item["local_file_name"]] = (int(match.group(1)), item)

    records = []
    used_indices = set()
    for name, (number, item) in sorted(sheets.items(), key=lambda pair: pair[1][0]):
        cells = parse_sheet_cells(item["decoded_xml"])
        header = [source_line(item)]
        if name in display_names:
            header.append(f"# sheet name: {display_names[name]} (from xl/workbook.xml)")
        cell_rows = []
        grid = {}
        references_shared = False
        for cell in cells:
            value, index = resolve_cell(cell, shared_strings)
            if index is not None:
                used_indices.add(index)
                references_shared = True
            grid.setdefault(cell["row"], {})[cell["column"]] = value
            cell_rows.append([cell["ref"], cell["type"], cell["raw"] if cell["raw"] is not None else "", value,
                              cell["formula"] or "", formats.get(cell["style"] or "0", "")])
        if references_shared:
            header.append(source_line(shared_item, "shared strings") if shared_item
                          else "# shared strings: xl/sharedStrings.xml not recovered")
        text_dir.mkdir(parents=True, exist_ok=True)
        columns = sorted({column for row in grid.values() for column in row})
        grid_rows = [["Row"] + [column_letters(column) for column in columns]]
        grid_rows += [[row] + [grid[row].get(column, "") for column in columns] for row in sorted(grid)]
        grid_path, cells_path = text_dir / f"sheet{number}.csv", text_dir / f"sheet{number}_cells.csv"
        _write_csv(grid_path, header, grid_rows)
        _write_csv(cells_path, header, [CELL_COLUMNS] + cell_rows)
        records.append({"name": name, "sheet_name": display_names.get(name), "cells": len(cells),
                        "output": grid_path.relative_to(out_dir).as_posix(),
                        "cells_output": cells_path.relative_to(out_dir).as_posix()})

    unmapped = [(index, text) for index, text in enumerate(shared_strings) if index not in used_indices]
    if unmapped:
        text_dir.mkdir(parents=True, exist_ok=True)
        _write_csv(text_dir / "sharedStrings_unmapped.csv",
                   [source_line(shared_item), "# shared strings not referenced by any recovered cell"],
                   [["index", "string"]] + [[index, text] for index, text in unmapped])
    if not records and not unmapped:
        safe_print("[XLSX] No recoverable cell values found.")
    return records
