import os
import re
from .common import (
    extract_data_from_shared_strings,
    extract_data_from_sheet,
    parse_xlsx_styles,
    parse_xlsx_rels_file,
    save_unmapped_to_csv,
    display_and_save_table_to_csv
)

def process_extracted_xlsx_data(local_file_xml, file_path):
    out_dir = os.path.join(os.path.dirname(file_path), f"output_{os.path.basename(file_path).rsplit('.', 1)[0]}")
    os.makedirs(out_dir, exist_ok=True)

    shared_xml = next((i['local_file_data'] for i in local_file_xml if i['local_file_name'] == 'xl/sharedStrings.xml'), "")
    styles_xml = next((i['local_file_data'] for i in local_file_xml if i['local_file_name'] == 'xl/styles.xml'), "")
    rels_xml = next((i['local_file_data'] for i in local_file_xml if i['local_file_name'] == 'xl/_rels/workbook.xml.rels'), "")
    sheets = [i['local_file_data'] for i in local_file_xml if re.match(r'xl/worksheets/sheet\d+.xml', i['local_file_name'])]

    shared_strings = extract_data_from_shared_strings(shared_xml)
    style_map = parse_xlsx_styles(styles_xml) if styles_xml else {}
    all_values = []

    for i, sheet in enumerate(sheets, 1):
        mapped, unmapped = extract_data_from_sheet(sheet, shared_strings, style_map)
        all_values.extend([v for row in mapped.values() for v in row.values()])
        display_and_save_table_to_csv(mapped, f"output_damaged_sheet{i}.csv", out_dir)

    unmapped_final = set(shared_strings) - set(all_values)
    if unmapped_final:
        save_unmapped_to_csv(unmapped_final, out_dir)

    if rels_xml:
        rels_info = parse_xlsx_rels_file(rels_xml)
        with open(os.path.join(out_dir, "relationships_xlsx.txt"), 'w', encoding='utf-8') as f:
            f.write("[Relationships]\n" + "\n".join(rels_info))

    print(f"\n[XLSX output path] {out_dir}")