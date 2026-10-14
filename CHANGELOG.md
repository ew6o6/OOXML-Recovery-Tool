# Changelog

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
