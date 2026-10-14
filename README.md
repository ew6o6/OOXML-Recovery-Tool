# OOXML Recovery Toolkit (ORT) 0.2.1

ORT is a command-line tool that recovers **document content** and **image files** from damaged OOXML files (`.docx`, `.xlsx`, `.pptx`). It reads ZIP local file headers directly and does not depend on the central directory. It extracts whatever parts survive, and it writes a manifest that states what was recovered and how strongly each result is supported.

ORT does not try to make the file open in Microsoft Office again, and it does not reproduce the original layout, fonts, or pagination.

Repository: https://github.com/ew6o6/OOXML-Recovery-Tool

## Requirements

- Python 3.8+. This release was tested on Python 3.14.0, Windows 11, and zlib 1.3.1 (zlib-ng).
- Tool code SHA-256 (recorded in every manifest): `122ac3a958e4b4a201dfb8c71e47a7ea6a1833a89b18be683c89b2e8eab8e833`
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
ort file.pptx --output-root results/          # write output folders under results/
ort file.docx --salvage-fragments             # experimental fragment salvage (see below)
python -m ort path/to/file.xlsx               # same, without installing
```

## Output

Each input `NAME.ext` produces a folder `output_NAME/`:

| Path | Content |
|---|---|
| `recovery_manifest.json` | Status of the package and of every entry, spans, images, fragments, and the tool code SHA-256 (schema v2) |
| `NAME_localFiles.json` | Every discovered entry, including decoded XML |
| `metadata/metadata.txt` | Author, revision, and dates from `docProps/core.xml` (validated XML only) |
| `text_parts/*.txt`, `document_content.txt` | DOCX text of document, header, footer, footnotes, endnotes, comments, and glossary, one file per part plus a combined file |
| `output_damaged.docx` | DOCX body text in a plain Word file (only when body text is recovered) |
| `output_damaged_sheetN.csv`, `unmapped_data.csv` | XLSX cell values and shared strings |
| `slideN.txt`, `notesSlideN.txt` | PPTX slide and speaker-note text |
| `relationships*.txt` | Relationship targets |
| `img/verified/` | Images whose CRC-32 matches the stored value (byte-identical to the original) |
| `img/decodable/` | Images with no CRC reference that decode completely |
| `img/unverified/` | Images with no CRC reference in formats without a validator (e.g. EMF, WMF) |
| `img/incomplete/` | Images with a recognized signature that are truncated or fail their CRC check, kept for examination |
| `img/rejected_stored/*.bin` | Bytes of uncompressed (STORED) images that were rejected, kept only for examination and never counted as recovered images |
| `fragments/` | Only with `--salvage-fragments`; never merged into the text above |

Images keep their original part name and extension. Bytes with no evidence of being the claimed content are **not written**; the manifest records only their offset and size. The one exception is a rejected STORED image: its bytes are still original data after the damaged region, so they are written as a diagnostic `.bin` file and the image stays `rejected`.

## How recovery works

1. **Find local headers.** ORT scans for local file headers anywhere in the file. A surviving central-directory record is trusted only when its offset and name match the local header. A part name is taken from the declared length (UTF-8 or CP437, with a safe-path check). A name-grammar fallback is used only when the declared length is invalid.
2. **Determine boundaries.**
   - DEFLATE: the stream end, plus the data descriptor if present.
   - STORED: central directory, then header sizes, then data descriptor, then the next header. The first boundary whose CRC-32 matches is preferred.
   - ORT skips past a payload only when its end is verified. Afterwards it re-scans the whole file and merges headers it skipped, excluding headers inside CRC-verified payloads (embedded packages).
3. **Decode and validate.** Two things are recorded separately for each entry:
   - `decode_status`: complete, stored, partial prefix, or failed
   - `content_validation_status`: CRC-verified, structure-verified, validated prefix, unverified, CRC mismatch (kept and flagged), or rejected (not written)
   - XML evidence means a strict UTF-8 start, an OOXML namespace URI, and a complete start tag of an element that belongs in that part.
   - A complete stream whose CRC-32 differs but carries XML evidence is kept as `crc_mismatch`. Invalid UTF-8 sequences are replaced with U+FFFD and counted, and the SHA-256 of the decoded bytes is recorded.
   - A validated prefix ends at the last complete UTF-8 character.
4. **Extract and report.** Text, images, and metadata are extracted from validated entries. The manifest reports the following statuses: inventory, core content, document content, image, and overall. Overall is `complete` only when the inventory is complete and no entry failed.

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

## Validation summary (this release)

| Set | Result |
|---|---|
| 2,496 damaged DOCX files | Body text exact in 2,240; in the 256 files with a destroyed body, the footer was recovered exactly in all 256 |
| 278 intact OOXML controls (CRC-valid, SHA-unique) | 11,388/11,388 XML parts, 138/138 DOCX text parts, and 1,532/1,532 images byte-identical; all statuses complete |
| Adversarial and derived image damage (700 cases) | Entry loss 0 after mid-stream damage; 1,125/1,125 STORED images recovered with sizes zeroed; junk never written; 0 images labelled verified or decodable that differ from the original |
| Unit tests | 31 passed (`python -m pytest tests`) |

## Known limitations

- **Completeness without a central directory.** If the central directory is missing and a content part's header is destroyed, `document_content_status` can still read `complete`, and `image_status` can read `all_preserved`, because the lost part is invisible. Treat `complete` as proven only when `inventory_status` is `complete`. A revised definition is planned.
- Each part is decoded up to 16 MiB; larger parts (for example video) are reported as failed.
- **A validated prefix can end with non-original text.** After damage inside a DEFLATE stream, the decoder may keep producing output (for example, wrong copies of earlier text) until it detects an error. `prefix_validated` means that the start of the part carries OOXML evidence, not that every byte is original. Treat the end of such a part with caution.
- Not evaluated: encrypted (OLE-wrapped) OOXML, ZIP64, damaged CRC fields, and packages written by non-Microsoft applications.
