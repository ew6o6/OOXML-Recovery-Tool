# Changelog

## 0.3.0 (2026-10-14)

The recovery engine is unchanged: decoding, validation statuses, decoded XML, images, and the summary statuses are identical to 0.2.1 on every validation file. This release fixes and restructures the output.

### Fixed
- XLSX: cell values could be assigned to the wrong cell. Shared strings were read with a pattern that skipped `<t xml:space="preserve">` and split rich-text strings, which shifted every later index. Each `<si>` is now one string, with all runs joined and phonetic runs excluded. Inline rich strings are read in full.
- XLSX: number formats were looked up in all `<xf>` elements instead of `cellXfs`, and the format was appended to the cell value. Values are now stored without the format; the format is a separate column of `sheetN_cells.csv`.
- XLSX: sheet files were numbered in order of discovery, so `sheet3.xml` became `sheet1` when earlier sheets were lost. Files now use the part number, and the display name is taken from `xl/workbook.xml` when it survives. Grid columns are ordered A…Z, AA (not alphabetically).
- PPTX: runs were joined with spaces (`Forens ic`), paragraphs of a shape were merged into one line, and table text was missing. Runs are now joined as stored, each paragraph is a line, line breaks are kept, and tables and grouped shapes are included.
- Output folders were named without the extension and reused. `report.docx`, `report.xlsx`, and `report.pptx` wrote into the same folder, and a second run left old files behind.
- Metadata was written twice (the PPTX handler overwrote the general one).

### Changed (output layout 2, manifest schema 3)
- Output folder `output_<name>.<ext>/`; an existing folder is never reused (`_2`, `_3`, …).
- Text and tables are in `text/`; `output_damaged.docx` is now `text/body_text.docx`.
- Every recovered non-image part is written to `parts/<original path>`, including embedded files. The manifest entry records its `output` path and SHA-256.
- `<name>_localFiles.json` is no longer written. `--dump-entries` writes the same content as `entries.json`.
- `metadata.txt` (top level) reads more fields from `docProps/core.xml` and `docProps/app.xml`; the manifest has a `metadata` section.
- One `relationships.txt` lists every recovered `.rels` part, including `TargetMode`.
- Every text and CSV file starts with a provenance line: source part, validation status, and its meaning.
- Manifest: `schema_version` 3, `output_layout`, `tool_version`, `max_part_size_bytes`, `evaluation_definitions_v2_1`, entry field `extension` (was `img_ext`), and `tail_may_be_non_original` on `prefix_validated` entries.
- A rejected STORED image whose bytes are one repeated value (e.g. zero-filled) is no longer written to `img/rejected_stored/`; the image record gives the value in `uniform_fill_byte` (`null` otherwise).
- Python API: `get_file_hex` returns `RecoveryResult(entries, extension, output_directory, overall_status)`.

### Added
- Reading outputs for every input, which never affect recovery:
  - `recovered_report.html`: static page with the recovered content in reading order, image positions linked to `img/`, a warning at every part that may be non-original, and an index of parts and images by status.
  - `recovered_<name>.<ext>`: the recovered content rebuilt in the input's format (python-docx, openpyxl, python-pptx); not the original, and marked as such.
  - `recovered_<name>.pdf`: PDF preview through LibreOffice or Microsoft Office (Windows) when installed; `--no-pdf` skips it. Manifest fields `recovered_report`, `rebuilt_document`, and `preview_pdf` (with `how_to` steps when no PDF was written).
- An undamaged package is reported as `intact` and not processed (no output folder): every Central Directory record matches its Local File Header (name, flags, method, CRC-32, sizes), the entries fill the file up to the Central Directory, the EOCD record ends the file, and every entry decodes to its CRC-32. `--process-intact` processes it anyway.
- Dependencies `openpyxl` and `python-pptx`.
- `--version`, `-r/--recursive`, `--dump-entries`, `--max-part-size MIB`, `--process-intact`.
- Exit status: 0 complete or intact, 1 partial, 2 failed, indeterminate, or not OOXML.
- 32 unit tests (63 in total).

### Removed
- Unused code: `main_dispatcher.py`, `decode_utf8`, `decompress_deflate_hex`, `decompress_deflate_bytes`, `is_structurally_valid_ooxml_raw`, `extract_local_file_names_raw`, `parse_xlsx_rels_file`, `extract_file_name`, and the unused DOCX style lookup.

## 0.2.1 (2026-10-14)

This is the version evaluated in the benchmark of the accompanying manuscript (tool code SHA-256 `122ac3a9…`).

### Decoding and validation
- XML evidence for a prefix now requires a strict UTF-8 start, an OOXML namespace URI, and a complete start tag of an element that belongs in that part.
- A complete stream whose CRC-32 differs but carries XML evidence is kept as `crc_mismatch` instead of being rejected. Invalid UTF-8 sequences are replaced with U+FFFD and counted, and the SHA-256 of the decoded bytes and the first error offset are recorded.
- A validated prefix ends at the last complete UTF-8 character; the removed byte count is recorded.

### Images
- Rejected STORED image bytes are written to `img/rejected_stored/<name>.bin` for examination. They are never counted as recovered images.

### Packaging
- MIT license; author name corrected.

## 0.2.0 (2026-10-13)

### Discovery and boundaries
- Local headers are found anywhere in the file, and data before the first header is ignored.
- The part name is taken from the declared length (UTF-8 or CP437) with a safe-path check. Invalid lengths fall back to a name grammar. Non-ASCII, upper-case, and uncommon extensions are supported.
- Surviving central-directory records are used when their offset and name match.
- STORED boundaries use the central directory, then header sizes, then the data descriptor, then the next header, verified by CRC-32.
- A payload is skipped only when its end is verified. Skipped headers are merged by a re-scan, and headers inside CRC-verified payloads are excluded.

### Decoding and validation
- Prefix recovery is linear time and stops at the first DEFLATE error.
- The decode status and the content validation status (CRC-32, XML or image structure, validated prefix) are recorded separately.
- Rejected bytes are not written; only their offset and size are recorded.
- Images are stored in `img/verified`, `decodable`, `unverified`, or `incomplete` under their original name.

### Extraction
- A missing core part no longer stops processing; surviving parts (footer, glossary, images, and so on) are extracted.
- DOCX: text is extracted per part (document, header, footer, footnotes, endnotes, comments, glossary). Run boundaries no longer insert spaces.
- PPTX: speaker-note text is extracted. XLSX and PPTX now parse only validated XML.
- Experimental `--salvage-fragments` option.

### Reporting
- `recovery_manifest.json` schema v2 reports inventory, core content, document content, image, and overall (strict) statuses, failure counts, orphan spans, and the tool code SHA-256.
- New `--output-root` option.

### Packaging
- Removed the unused `tabulate` dependency.
