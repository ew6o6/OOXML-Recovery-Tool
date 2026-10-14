# OOXML Recovery Toolkit (ORT) 0.3.0

ORT is a command-line tool that recovers **document content** and **image files** from damaged OOXML files (`.docx`, `.xlsx`, `.pptx`). It reads ZIP local file headers directly and does not depend on the central directory. It extracts whatever parts survive, and it writes a manifest that states what was recovered and how strongly each result is supported.

ORT does not try to make the file open in Microsoft Office again, and it does not reproduce the original layout, fonts, or pagination.

Repository: https://github.com/ew6o6/OOXML-Recovery-Tool

> **Version used in the manuscript:** the benchmark was run with release `v0.2.1` (tool code SHA-256 `122ac3a9…`). Release 0.3.0 changes only the output layer; its recovery engine produces identical results (see [Validation](#validation)).

## Requirements

- Python 3.8+. Tested on Python 3.14.0, Windows 11, and zlib 1.3.1 (zlib-ng).
- Dependencies: `beautifulsoup4`, `lxml`, `python-docx`, `pillow`

## Installation

```bash
git clone https://github.com/ew6o6/OOXML-Recovery-Tool
cd OOXML-Recovery-Tool
python -m venv .venv            # optional
pip install .
```

## Usage

```bash
ort path/to/file.docx                         # one file; output next to the input
ort path/to/folder/                           # every file in a folder
ort path/to/folder/ -r                        # also subfolders (output_* folders are skipped)
ort file.pptx --output-root results/          # write output folders under results/
ort file.xlsx --max-part-size 256             # decode parts up to 256 MiB (default 16)
ort file.docx --dump-entries                  # also write entries.json (every entry, payload as hex)
ort file.docx --salvage-fragments             # experimental fragment salvage (see below)
ort --version
python -m ort path/to/file.xlsx               # same, without installing
```

ORT detects the format from the package content, not from the file extension.

**Exit status:**

| Code | Meaning |
|---|---|
| 0 | Every input is `complete` |
| 1 | At least one input is `partial` |
| 2 | At least one input failed, is indeterminate, or is not OOXML, or the path holds no file |

## Output

Each input `NAME.ext` produces a new folder `output_NAME.ext/`. An existing folder is never reused: a second run writes `output_NAME.ext_2/`, and so on.

```
output_report.docx/
├── recovery_manifest.json        status of the package and of every entry (schema v3)
├── metadata.txt                  docProps/core.xml and docProps/app.xml
├── relationships.txt             every recovered .rels part, grouped by source part
├── text/
│   ├── document_content.txt      all text parts in one file (DOCX, PPTX)
│   ├── word_document.xml.txt …   DOCX: document, header, footer, footnotes, endnotes, comments, glossary
│   ├── body_text.docx            DOCX: body text in a plain Word file
│   ├── slide3.txt, notesSlide3.txt               PPTX: slide and speaker-note text
│   ├── sheet2.csv, sheet2_cells.csv              XLSX: values on a grid, and one row per cell
│   └── sharedStrings_unmapped.csv                XLSX: shared strings no recovered cell refers to
├── parts/                        every recovered non-image part under its original path
├── img/
│   ├── verified/                 CRC-32 matches the stored value (byte-identical)
│   ├── decodable/                no CRC reference; decodes completely
│   ├── unverified/               no CRC reference; no validator for the format (e.g. EMF, WMF)
│   ├── incomplete/               recognized signature, but truncated or CRC mismatch
│   └── rejected_stored/*.bin     bytes of rejected uncompressed images, for examination only
└── fragments/                    only with --salvage-fragments
```

- **Provenance header.** Every text and CSV file starts with a line such as `# source: word/document.xml | prefix_validated | start of the part only; the end may contain non-original text`. The same header precedes each part in `document_content.txt` and each block in `metadata.txt`.
- **Part numbers.** `sheet2.csv` and `slide3.txt` come from `sheet2.xml` and `slide3.xml`. Numbers are never reassigned when other parts are missing. When `xl/workbook.xml` survives, the sheet's display name is added to the CSV header.
- **XLSX values.** `sheetN.csv` shows the stored value of each cell (shared strings resolved, cached formula results, `TRUE`/`FALSE`). Dates stay as Excel serial numbers. `sheetN_cells.csv` lists the cell type, raw value, resolved value, formula, and number format of every cell.
- **`parts/`.** Each file is the decoded part (XML or binary, e.g. embedded workbooks) with the original path. The manifest entry records its `output` path and SHA-256.
- **Rejected bytes.** Bytes with no evidence of being the claimed content are **not written**; the manifest records only their offset and size. The one exception is a rejected STORED image: its bytes are still original data after the damaged region, so they are written as a diagnostic `.bin` file, and the image stays `rejected`.

### Validation status of a part

| `content_validation_status` | Meaning | Where the content goes |
|---|---|---|
| `crc_verified` | CRC-32 matches the stored value: byte-identical to the original | text, `parts/`, `img/verified/` |
| `structure_verified` | No CRC reference; complete and well-formed (XML) or fully decodable (image) | text, `parts/`, `img/decodable/` |
| `prefix_validated` | The stream broke; the start carries OOXML evidence. **The end may contain non-original text** | text, `parts/` |
| `crc_mismatch` | Complete stream, but CRC-32 differs; content may be altered. Invalid UTF-8 is shown as U+FFFD | text, `parts/`, `img/incomplete/` |
| `unverified` | No CRC reference and no structural proof | text, `parts/`, `img/unverified/` or `incomplete/` |
| `rejected` | No evidence of being the claimed content | not written (except `img/rejected_stored/*.bin`) |

The manifest also reports five summary statuses for the package: `inventory_status`, `core_content_status`, `document_content_status`, `image_status`, and `overall_status`. `overall_status` is `complete` only when the inventory is complete and no entry failed.

## How recovery works

1. **Find local headers.** ORT scans for local file headers anywhere in the file. A surviving central-directory record is trusted only when its offset and name match the local header. A part name is taken from the declared length (UTF-8 or CP437, with a safe-path check). A name-grammar fallback is used only when the declared length is invalid.
2. **Determine boundaries.**
   - DEFLATE: the stream end, plus the data descriptor if present.
   - STORED: central directory, then header sizes, then data descriptor, then the next header. The first boundary whose CRC-32 matches is preferred.
   - ORT skips past a payload only when its end is verified. Afterwards it re-scans the whole file and merges headers it skipped, excluding headers inside CRC-verified payloads (embedded packages).
3. **Decode and validate.** Two things are recorded separately for each entry:
   - `decode_status`: complete, stored, partial prefix, or failed
   - `content_validation_status`: see the table above. XML evidence means a strict UTF-8 start, an OOXML namespace URI, and a complete start tag of an element that belongs in that part.
4. **Write and report.** Parts, text, images, and metadata are written from validated entries, and the manifest records the evidence for each.

A missing core part (for example `word/document.xml`) does not stop processing. The remaining parts (footer, glossary, images, and so on) are still extracted.

## What can and cannot be recovered

| Damage | Result |
|---|---|
| Header fields damaged (sizes 0 or `0xFFFFFFFF`, file-name length 0 or `0xFFFF`) | Recovered byte-identical |
| Central directory or EOCD missing | No effect on extraction |
| Data before the first header (prepended junk) | Ignored; the package is processed |
| Damage confined to one part | All other parts recovered byte-identical (CRC-verified) |
| Core part's header destroyed | Surviving parts are still extracted; core reported `missing` |
| STORED part with lost sizes | Boundary rebuilt from CD, data descriptor, or the next header and verified by CRC |
| Non-standard part names (upper-case extensions, non-ASCII, parentheses, `.emz`, `.mp4`) | Discovered and extracted |
| Damage inside one DEFLATE stream | Following parts are not lost. The part itself yields at most a validated prefix |
| Image damaged in the middle or at the end | Kept in `img/incomplete/`; a truncated tail usually leaves an exact prefix |
| **DEFLATE stream damaged at its start** (block header or Huffman table lost) | **That part's content is not recoverable** by normal decoding; see fragment salvage |
| **Image damaged at its start** | **Not recoverable**: the signature is gone, and the bytes are reported as lost or rejected |
| Part overwritten or deleted entirely | Not recoverable |

### Fragment salvage (experimental, `--salvage-fragments`)

When a DEFLATE stream is damaged at its start, ORT restarts decoding after `00 00 FF FF` flush markers inside rejected payloads and orphan spans. A fragment is accepted only if it reaches the stream end and carries OOXML evidence, and it is saved separately.

Such markers were observed almost only in Microsoft Word `document.xml`:
- About 57% of DOCX bodies had them.
- Median recovered size was about 10% of the bytes and 8% of the text.
- XLSX and PPTX parts almost never had them.

Fragments are never complete content and must not be counted as recovered body text.

## Code layout

```
OOXML-Recovery-Tool/
├── ort/
│   ├── __init__.py              package version
│   ├── __main__.py              `python -m ort` entry point
│   ├── main.py                  command line: options, file and folder traversal, exit status
│   └── core/
│       ├── extractor.py         recovery pipeline and manifest (entry point: get_file_hex)
│       ├── decoder.py           DEFLATE decoding and the two status axes
│       ├── utils.py             output folders, parts/, image validation, metadata, relationships
│       ├── fragment_salvage.py  --salvage-fragments
│       └── handlers/
│           ├── common.py        DOCX text and XLSX parsing helpers
│           ├── docx.py          DOCX text parts
│           ├── xlsx.py          XLSX sheets
│           └── pptx.py          PPTX slides and notes
├── tests/test_recovery.py       unit tests (`python -m pytest tests`)
├── setup.py, requirements.txt
├── CHANGELOG.md
└── LICENSE                      MIT
```

| Module | Main functions | Role |
|---|---|---|
| `main.py` | `main`, `iter_input_files`, `process_file` | Parses options, collects input files, calls the extractor, and turns the statuses into the exit status |
| `core/extractor.py` | `get_file_hex` | Runs the pipeline for one file and writes the output folder and manifest |
| | `parse_central_directory`, `_trusted_cd` | Reads surviving central-directory records and trusts one only when its offset and name match |
| | `recover_filename_length`, `decode_part_name` | Recovers the part name from the declared length, or from a name grammar when the length is damaged |
| | `discover_entries`, `stored_boundary` | Finds local headers (sequential scan with verified jumps, then re-scan) and bounds STORED payloads |
| | `_core_status`, `_document_content_status`, `_overall_status` | Summary statuses |
| `core/decoder.py` | `decompress_with_status`, `inflate_probe` | DEFLATE decoding: the full stream, or the prefix before the first error |
| | `assess_content`, `validate_xml_prefix` | Decide `content_validation_status` from CRC-32, XML evidence, or image structure |
| | `decode_local_file_data` | Applies both to every entry and keeps the decoded bytes |
| `core/utils.py` | `extract_img_file`, `validate_image_bytes`, `detect_image_format` | Image validation and the `img/` buckets |
| | `new_output_directory`, `write_part` | Output folder naming and safe `parts/` paths |
| | `extract_metadata`, `write_relationships`, `source_line` | `metadata.txt`, `relationships.txt`, and provenance headers |
| `core/handlers/*.py` | `process_extracted_*_data` | Text and tables per format, from validated XML only |

## Validation

| Set | Result |
|---|---|
| 278 intact OOXML controls (CRC-valid, SHA-unique) | 11,388/11,388 XML parts, 138/138 DOCX text parts, and 1,532/1,532 images byte-identical; all statuses complete |
| 2,496 damaged DOCX files | Body text exact in 2,240; in the 256 files with a destroyed body, the footer was recovered exactly in all 256 |
| Adversarial and derived image damage (700 cases) | Entry loss 0 after mid-stream damage; 1,125/1,125 STORED images recovered with sizes zeroed; junk never written; 0 images labelled verified or decodable that differ from the original |
| 0.3.0 vs 0.2.1 recovery engine (all 3,474 files above) | Identical on all 3,474 files: decode and validation statuses and decoded-XML SHA-256 of 143,810 entries, 9,379 image records, DOCX text, and the summary statuses |
| 0.3.0 output layer on the 278 controls | `parts/`: 11,599/11,599 files byte-identical to the original part. XLSX (96 workbooks): 5,269/5,269 text, 1,586/1,586 numeric, and 2,997/2,997 empty cells equal to openpyxl (538 date cells not compared); sheet names correct in 96/96. PPTX (89 decks, 2,047 slides): no slide misses text that python-pptx reads; 1,994 identical, 53 with additional text from alternate content (`mc:AlternateContent`), which python-pptx does not read. DOCX core properties: 271/271 values equal to python-docx (2 files without `docProps/core.xml` excluded) |
| Unit tests | 42 passed (`python -m pytest tests`) |

## Known limitations

- **A validated prefix can end with non-original text.** After damage inside a DEFLATE stream, the decoder may keep producing output (for example, wrong copies of earlier text) until it detects an error. `prefix_validated` means that the start of the part carries OOXML evidence, not that every byte is original. The manifest marks such entries with `tail_may_be_non_original`, and the text headers say so.
- **Completeness without a central directory.** If the central directory is missing and a content part's header is destroyed, `document_content_status` can still read `complete`, and `image_status` can read `all_preserved`, because the lost part is invisible. Treat `complete` as proven only when `inventory_status` is `complete`.
- Parts are decoded up to 16 MiB by default (`--max-part-size`); larger parts are reported as failed.
- PPTX text comes from shapes, groups, and tables. Text inside charts and SmartArt is not extracted.
- Not evaluated: encrypted (OLE-wrapped) OOXML, ZIP64, damaged CRC fields, and packages written by non-Microsoft applications.

## License

MIT. See [LICENSE](LICENSE).
