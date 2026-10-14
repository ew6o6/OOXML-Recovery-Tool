import contextlib
import io
import json
import struct
import sys
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from ort.core.decoder import decode_local_file_data, recover_deflate_prefix, validate_xml_prefix
from ort.core.extractor import get_file_hex, recover_filename_length
from ort.core.fragment_salvage import recover_deflate_fragments
from ort.core.handlers.common import extract_xml_text_for_docx
from ort.core.utils import detect_image_format, validate_image_bytes


BODY = (
    b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b'<w:body><w:p><w:r><w:t>Hel</w:t></w:r><w:r><w:t>lo</w:t></w:r>'
    b'<w:r><w:t xml:space="preserve"> world</w:t></w:r></w:p></w:body></w:document>'
)
FOOTER = (
    b'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    b'<w:p><w:r><w:t>Footer text</w:t></w:r></w:p></w:ftr>'
)


def package(parts, compression=zipfile.ZIP_DEFLATED):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", compression=compression) as archive:
        archive.writestr("[Content_Types].xml", b"<Types/>")
        for name, data in parts.items():
            archive.writestr(name, data)
    return stream.getvalue()


def png_bytes():
    stream = io.BytesIO()
    Image.new("RGB", (2, 2), (120, 30, 200)).save(stream, format="PNG")
    return stream.getvalue()


class ApprovedRecoveryTests(unittest.TestCase):
    def test_run_boundaries_do_not_insert_spaces(self):
        self.assertEqual(extract_xml_text_for_docx(BODY.decode(), {}), "Hello world")

    def test_stored_xml_uses_method_metadata(self):
        entries = [
            {
                "local_file_name": "word/document.xml",
                "local_file_data": BODY.hex(),
                "compression_method": 0,
            }
        ]
        decode_local_file_data(entries)
        self.assertEqual(entries[0]["decoded_xml"], BODY.decode())
        self.assertEqual(entries[0]["decode_status"], "stored")

    def test_prefix_validation_rejects_short_garbage(self):
        self.assertEqual(validate_xml_prefix(b"\r"), (False, "not_xml_start"))
        self.assertTrue(validate_xml_prefix(BODY[:180])[0])

    def test_linear_deflate_prefix(self):
        raw = zlib.compress(BODY)[2:-4]
        self.assertEqual(recover_deflate_prefix(raw), BODY)
        prefix = recover_deflate_prefix(raw[:-6])
        self.assertTrue(prefix and BODY.startswith(prefix))
        self.assertEqual(recover_deflate_prefix(b"\xff" * 1000), b"")

    def test_invalid_filename_length_is_bounded(self):
        part = bytearray(30) + b"word/document.xml" + b"\xff" * 4
        for declared in (0, 65535):
            struct.pack_into("<H", part, 26, declared)
            self.assertEqual(recover_filename_length(part), 17)

    def test_filename_length_does_not_absorb_deflate_bytes(self):
        name = b"word/webSettings.xml"
        part = bytearray(30) + name + b"rc" + b"payload"
        struct.pack_into("<H", part, 26, len(name) + 2)
        self.assertEqual(recover_filename_length(part), len(name))

    def test_root_relationship_part_name_is_valid(self):
        name = b"_rels/.rels"
        part = bytearray(30) + name + b"payload"
        struct.pack_into("<H", part, 26, len(name))
        self.assertEqual(recover_filename_length(part), len(name))

    def test_extensionless_ppt_metadata_part_name_is_valid(self):
        name = b"ppt/metadata"
        part = bytearray(30) + name + b"payload"
        struct.pack_into("<H", part, 26, len(name))
        self.assertEqual(recover_filename_length(part), len(name))

    def test_prefixed_package_is_processed_without_pseudo_entry(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "prefixed.docx"
            output_root = Path(temp_dir) / "results"
            input_path.write_bytes(b"J" * 64 + package({"word/document.xml": BODY}))
            with contextlib.redirect_stdout(io.StringIO()):
                entries, extension = get_file_hex(str(input_path), output_root=output_root)
            self.assertEqual(extension, ".docx")
            self.assertNotIn("J", "".join(item["local_file_name"] for item in entries))
            manifest = json.loads((output_root / "output_prefixed" / "recovery_manifest.json").read_text())
            self.assertEqual(manifest["orphan_spans"][0]["kind"], "unowned_prefix")

    def test_footer_is_extracted_when_body_is_missing(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "footer_only.docx"
            output_root = Path(temp_dir) / "results"
            input_path.write_bytes(package({"word/footer1.xml": FOOTER}))
            with contextlib.redirect_stdout(io.StringIO()):
                _, extension = get_file_hex(str(input_path), output_root=output_root)
            self.assertEqual(extension, ".docx")
            out = output_root / "output_footer_only"
            manifest = json.loads((out / "recovery_manifest.json").read_text())
            self.assertEqual(manifest["inventory_status"], "complete")
            # v2: a missing core part is "missing" regardless of other recovered parts;
            # nothing was lost from this intact package, so overall is complete (Q2 strict rule).
            self.assertEqual(manifest["core_content_status"], "missing")
            self.assertEqual(manifest["document_content_status"], "complete")
            self.assertEqual(manifest["overall_status"], "complete")
            self.assertEqual(manifest["content_parts"][0]["name"], "word/footer1.xml")
            self.assertIn("Footer text", (out / "document_content.txt").read_text())
            self.assertFalse((out / "output_damaged.docx").exists())

    def test_complete_body_has_conservative_complete_status(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "clean.docx"
            output_root = Path(temp_dir) / "results"
            input_path.write_bytes(package({"word/document.xml": BODY}))
            with contextlib.redirect_stdout(io.StringIO()):
                get_file_hex(str(input_path), output_root=output_root)
            out = output_root / "output_clean"
            manifest = json.loads((out / "recovery_manifest.json").read_text())
            self.assertEqual(manifest["inventory_status"], "complete")
            self.assertEqual(manifest["core_content_status"], "complete")
            self.assertEqual(manifest["image_status"], "none")
            self.assertEqual(manifest["overall_status"], "complete")
            self.assertTrue((out / "output_damaged.docx").exists())
            self.assertTrue((out / "text_parts" / "word_document.xml.txt").exists())

    def test_png_crc_and_terminal_structure_are_checked(self):
        raw = png_bytes()
        self.assertEqual(detect_image_format(raw), "png")
        self.assertEqual(validate_image_bytes(raw, "png")[0], "decodable_structurally_complete")
        self.assertNotEqual(validate_image_bytes(raw[:-2], "png")[0], "decodable_structurally_complete")
        damaged = bytearray(raw)
        damaged[-16] ^= 1
        self.assertNotEqual(validate_image_bytes(bytes(damaged), "png")[0], "decodable_structurally_complete")

    def test_unverified_media_keeps_original_name_and_extension(self):
        emf = bytearray(64)
        emf[:4] = b"\x01\x00\x00\x00"
        emf[40:44] = b" EMF"
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir) / "media.docx"
            output_root = Path(temp_dir) / "results"
            input_path.write_bytes(
                package({"word/document.xml": BODY, "word/media/original.emf": bytes(emf)})
            )
            with contextlib.redirect_stdout(io.StringIO()):
                get_file_hex(str(input_path), output_root=output_root)
            out = output_root / "output_media"
            manifest = json.loads((out / "recovery_manifest.json").read_text())
            image = manifest["images"][0]
            # v2: integrity (CRC-32) and format validation are separate evidence fields.
            self.assertEqual(image["status"], "preserved_verified")
            self.assertEqual(image["integrity"], "crc_verified")
            self.assertEqual(image["format_validation"], "unsupported_format")
            self.assertTrue(image["output"].endswith("original.emf"))
            self.assertTrue((out / image["output"]).exists())
            self.assertEqual(manifest["image_status"], "all_preserved")
            self.assertEqual(manifest["overall_status"], "complete")

    def test_fragment_salvage_is_eof_and_evidence_gated(self):
        prefix = b'<w:document xmlns:w="x"><w:body><w:p><w:t>Lost</w:t></w:p>'
        suffix = b'<w:p><w:r><w:t>Recovered</w:t></w:r></w:p></w:body></w:document>'
        compressor = zlib.compressobj(level=6, wbits=-zlib.MAX_WBITS)
        stream = compressor.compress(prefix) + compressor.flush(zlib.Z_FULL_FLUSH)
        stream += compressor.compress(suffix) + compressor.flush(zlib.Z_FINISH)
        damaged = b"\x00" * 32 + stream[32:]
        recovered = list(recover_deflate_fragments(damaged, extension=".docx"))
        self.assertTrue(recovered)
        self.assertIn("Recovered", recovered[0]["text"])
        self.assertGreater(recovered[0]["non_whitespace_codepoints"], 0)
        self.assertEqual(list(recover_deflate_fragments(b"random" + b"\x00\x00\xff\xff" * 3, extension=".docx")), [])


# ----------------------------------------------------------------------------- v2 (F1–F6)


def local_entry(name, payload, method=0, crc=None, csize=None, usize=None, flags=0):
    raw_name = name.encode("utf-8")
    if not name.isascii():
        flags |= 0x800
    crc = zlib.crc32(payload) & 0xFFFFFFFF if crc is None else crc
    csize = len(payload) if csize is None else csize
    usize = len(payload) if usize is None else usize
    header = struct.pack("<4sHHHHHIIIHH", b"PK\x03\x04", 20, flags, method, 0, 0, crc, csize, usize,
                         len(raw_name), 0)
    return header + raw_name + payload


def entry_layout(data, name):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        info = archive.getinfo(name)
    name_length, extra_length = struct.unpack_from("<HH", data, info.header_offset + 26)
    return info.header_offset, info.header_offset + 30 + name_length + extra_length, info.compress_size


def run_tool(data, stem, salvage=False):
    temp_dir = tempfile.mkdtemp()
    input_path = Path(temp_dir) / f"{stem}.docx"
    input_path.write_bytes(data)
    with contextlib.redirect_stdout(io.StringIO()):
        get_file_hex(str(input_path), output_root=Path(temp_dir) / "results", salvage_fragments=salvage)
    out = Path(temp_dir) / "results" / f"output_{stem}"
    return json.loads((out / "recovery_manifest.json").read_text(encoding="utf-8")), out


class StatusV2Tests(unittest.TestCase):
    def test_nonstandard_part_names_are_discovered(self):
        names = ["word/media/IMAGE_X.PNG", "word/media/photo.JPG", "word/media/image_x.emz",
                 "word/media/\uadf8\ub9bc1.png", "word/media/image(1).png", "word/media/media1.mp4",
                 "word/media/%EA%B7%B8%EB%A6%BC2.png"]
        parts = {"word/document.xml": BODY}
        parts.update({name: png_bytes() for name in names})
        manifest, _ = run_tool(package(parts), "names")
        found = {item["local_file_name"] for item in manifest["entries"]}
        self.assertTrue(set(names) <= found, set(names) - found)
        self.assertEqual(manifest["inventory_status"], "complete")
        self.assertEqual({image["status"] for image in manifest["images"]}, {"preserved_verified"})

    def test_filename_cannot_escape_or_contain_controls(self):
        for bad in (b"../evil.xml", b"/abs.xml", b"C:/x.xml", b"word/a\x01b.xml", b"word\\x.xml"):
            part = bytearray(30) + bad + b"\x00" * 8
            struct.pack_into("<H", part, 26, len(bad))
            self.assertNotEqual(recover_filename_length(bytes(part)), len(bad), bad)

    def test_stored_media_with_zeroed_sizes_uses_cd_then_next_header(self):
        image = png_bytes()
        original = package({"word/media/image1.png": image, "word/document.xml": BODY},
                           compression=zipfile.ZIP_STORED)
        offset, _, _ = entry_layout(original, "word/media/image1.png")
        damaged = bytearray(original)
        struct.pack_into("<II", damaged, offset + 18, 0, 0)
        without_cd = bytes(damaged[: bytes(damaged).find(b"PK\x01\x02")])
        for label, data, expected in (("with_cd", bytes(damaged), "central_directory"),
                                      ("without_cd", without_cd, "next_header_fallback")):
            manifest, out = run_tool(data, f"stored_{label}")
            record = next(item for item in manifest["images"] if item["name"] == "word/media/image1.png")
            self.assertEqual(record["status"], "preserved_verified", label)
            self.assertEqual((out / record["output"]).read_bytes(), image, label)
            entry = next(item for item in manifest["entries"] if item["local_file_name"] == "word/media/image1.png")
            self.assertEqual(entry["boundary_status"], expected, label)

    def test_stored_data_descriptor_boundary(self):
        image = png_bytes()
        descriptor = struct.pack("<4sIII", b"PK\x07\x08", zlib.crc32(image) & 0xFFFFFFFF, len(image), len(image))
        data = (local_entry("word/media/image1.png", image, crc=0, csize=0, usize=0, flags=0x8) + descriptor
                + local_entry("word/document.xml", BODY))
        manifest, _ = run_tool(data, "descriptor")
        entry = next(item for item in manifest["entries"] if item["local_file_name"] == "word/media/image1.png")
        self.assertEqual(entry["boundary_status"], "data_descriptor")
        self.assertEqual(entry["content_validation_status"], "crc_verified")
        self.assertEqual(manifest["inventory_status"], "unknown")

    def test_crc_mismatch_image_is_kept_as_incomplete(self):
        image = png_bytes()
        original = package({"word/document.xml": BODY, "word/media/image1.png": image},
                           compression=zipfile.ZIP_STORED)
        _, start, size = entry_layout(original, "word/media/image1.png")
        damaged = bytearray(original)
        damaged[start + size // 2] ^= 0xFF
        manifest, _ = run_tool(bytes(damaged), "crcmismatch")
        record = next(item for item in manifest["images"] if item["name"] == "word/media/image1.png")
        self.assertEqual(record["integrity"], "crc_mismatch")
        self.assertEqual(record["status"], "preserved_incomplete")
        self.assertEqual(manifest["overall_status"], "partial")

    def test_junk_complete_stream_is_rejected_and_not_written(self):
        image = png_bytes() * 50
        original = package({"word/document.xml": BODY, "word/media/image1.png": image})
        _, start, size = entry_layout(original, "word/media/image1.png")
        junk = zlib.compressobj(wbits=-zlib.MAX_WBITS)
        junk_stream = junk.compress(b"\x01") + junk.flush()
        damaged = original[:start] + junk_stream + original[start + size:]
        manifest, out = run_tool(damaged, "junk")
        record = next(item for item in manifest["images"] if item["name"] == "word/media/image1.png")
        self.assertEqual(record["status"], "rejected")
        self.assertIsNone(record["output"])
        self.assertEqual(record["rejected_bytes"], 1)
        self.assertIsNotNone(record["payload_offset"])
        self.assertFalse(list((out / "img").rglob("image1.png")))

    def test_rejected_stored_image_is_kept_as_diagnostic_bin(self):
        image = png_bytes() * 20
        original = package({"word/document.xml": BODY, "word/media/image1.png": image},
                           compression=zipfile.ZIP_STORED)
        _, start, size = entry_layout(original, "word/media/image1.png")
        damaged = bytearray(original)
        damaged[start:start + 64] = b"\x00" * 64  # signature destroyed, CRC mismatch
        manifest, out = run_tool(bytes(damaged), "storedhead")
        record = next(item for item in manifest["images"] if item["name"] == "word/media/image1.png")
        self.assertEqual(record["status"], "rejected")
        self.assertIsNone(record["output"])
        self.assertEqual(manifest["image_status"], "failed")
        self.assertTrue(record["diagnostic_output"].endswith("rejected_stored/image1.png.bin"))
        self.assertEqual((out / record["diagnostic_output"]).read_bytes(), bytes(damaged[start:start + size]))
        entry = next(item for item in manifest["entries"] if item["local_file_name"] == "word/media/image1.png")
        self.assertNotIn("rejected_stored_image_hex", entry)

    def test_rejected_deflate_image_has_no_diagnostic_bin(self):
        image = png_bytes() * 50
        original = package({"word/document.xml": BODY, "word/media/image1.png": image})
        _, start, size = entry_layout(original, "word/media/image1.png")
        junk = zlib.compressobj(wbits=-zlib.MAX_WBITS)
        damaged = original[:start] + junk.compress(b"\x01") + junk.flush() + original[start + size:]
        manifest, out = run_tool(damaged, "deflatejunk")
        record = next(item for item in manifest["images"] if item["name"] == "word/media/image1.png")
        self.assertEqual(record["status"], "rejected")
        self.assertIsNone(record["diagnostic_output"])
        self.assertFalse((out / "img" / "rejected_stored").exists())

    # ------------------------------------------------------------------ F9 / v2.1

    def test_complete_stream_with_corrupt_middle_is_kept_as_crc_mismatch(self):
        tail = b"".join(b'<w:p><w:r><w:t>After %d</w:t></w:r></w:p>' % i for i in range(20))
        corrupt_body = BODY.replace(b"</w:body>", b"\xff\xfe\xfd" + tail + b"</w:body>")
        compressor = zlib.compressobj(wbits=-zlib.MAX_WBITS)
        stream = compressor.compress(corrupt_body) + compressor.flush()
        data = local_entry("word/document.xml", stream, method=8, crc=zlib.crc32(BODY), csize=len(stream),
                           usize=len(BODY))
        manifest, out = run_tool(data, "crcmiddle")
        entry = next(e for e in manifest["entries"] if e["local_file_name"] == "word/document.xml")
        self.assertEqual(entry["decode_status"], "complete")
        self.assertEqual(entry["content_validation_status"], "crc_mismatch")
        self.assertEqual(entry["utf8_replaced_sequences"], 3)
        self.assertEqual(entry["utf8_first_error_offset"], corrupt_body.index(b"\xff"))
        self.assertEqual(entry["decoded_bytes_sha256"], __import__("hashlib").sha256(corrupt_body).hexdigest())
        text = (out / "text_parts" / "word_document.xml.txt").read_text(encoding="utf-8")
        self.assertIn("Hello world", text)
        self.assertIn("After 19", text)
        self.assertEqual(manifest["core_content_status"], "partial")

    def test_partial_prefix_with_split_trailing_character_is_trimmed(self):
        body = BODY.replace(b"Hel", "가나다".encode())
        cut = body.index("다".encode()) + 2  # stop inside the 3-byte character
        compressor = zlib.compressobj(0, zlib.DEFLATED, -zlib.MAX_WBITS)  # stored blocks: output is exact
        stream = compressor.compress(body) + compressor.flush()
        truncated = stream[: 5 + cut]  # 5-byte stored-block header, then the first `cut` bytes
        item = {"local_file_name": "word/document.xml", "compressed_data_hex": truncated.hex(), "compression_method": 8}
        decode_local_file_data([item])
        self.assertEqual(item["decode_status"], "partial_prefix")
        self.assertEqual(item["content_validation_status"], "prefix_validated")
        self.assertEqual(item["utf8_trailing_bytes_removed"], 2)
        self.assertEqual(item["decoded_xml"].encode(), body[: cut - 2])

    def test_xml_evidence_requires_namespace_and_complete_start_tag(self):
        namespace = b'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        self.assertFalse(validate_xml_prefix(b"<w:document" + b"\xff" * 8, "word/document.xml")[0])
        self.assertFalse(validate_xml_prefix(b"<w:document " + namespace + b"\xff", "word/document.xml")[0])
        self.assertTrue(validate_xml_prefix(b"<w:document " + namespace + b">\xff\xfe", "word/document.xml")[0])
        self.assertFalse(validate_xml_prefix(b"<w:t>inserted marker</w:t>", "word/document.xml")[0])
        self.assertFalse(validate_xml_prefix(b"\x8a\x01<w:p>", "word/document.xml")[0])
        self.assertFalse(validate_xml_prefix(b"<a:p " + namespace + b">", "word/document.xml")[0])

    def test_garbage_complete_xml_without_evidence_stays_rejected(self):
        junk = b"<w:document" + bytes(range(128, 256))
        compressor = zlib.compressobj(wbits=-zlib.MAX_WBITS)
        stream = compressor.compress(junk) + compressor.flush()
        data = local_entry("word/document.xml", stream, method=8, crc=zlib.crc32(BODY), csize=len(stream))
        manifest, _ = run_tool(data, "junkxml")
        entry = next(e for e in manifest["entries"] if e["local_file_name"] == "word/document.xml")
        self.assertEqual(entry["content_validation_status"], "rejected")
        self.assertEqual(manifest["core_content_status"], "failed")

    def test_unverified_stream_end_does_not_skip_embedded_header(self):
        # A DEFLATE stream whose end lies beyond a real local header (stored blocks keep
        # that header literal). Its CRC does not verify and its size is unknown, so the
        # scanner must not jump over the embedded footer entry.
        footer_entry = local_entry("word/footer1.xml", FOOTER)
        compressor = zlib.compressobj(0, zlib.DEFLATED, -zlib.MAX_WBITS)
        stream = compressor.compress(b"<x/>" + footer_entry + b"<y/>") + compressor.flush()
        data = (local_entry("word/document.xml", stream, method=8, crc=zlib.crc32(BODY), csize=0, usize=len(BODY))
                + local_entry("word/endnotes.xml", FOOTER))
        manifest, out = run_tool(data, "falseeof")
        names = [item["local_file_name"] for item in manifest["entries"]]
        self.assertIn("word/footer1.xml", names)
        self.assertIn("Footer text", (out / "text_parts" / "word_footer1.xml.txt").read_text(encoding="utf-8"))

    def test_verified_stored_package_hides_inner_headers(self):
        inner = package({"xl/worksheets/sheet1.xml": b"<worksheet/>"})
        data = package({"word/document.xml": BODY, "word/embeddings/Book1.xlsx": inner},
                       compression=zipfile.ZIP_STORED)
        manifest, _ = run_tool(data, "embedded")
        names = [item["local_file_name"] for item in manifest["entries"]]
        self.assertNotIn("xl/worksheets/sheet1.xml", names)
        self.assertEqual(manifest["inventory_status"], "complete")

    def test_overall_is_strict_and_reports_auxiliary_failures(self):
        styles = b"<w:styles xmlns:w='x'>" + b"<a/>" * 200 + b"</w:styles>"
        original = package({"word/document.xml": BODY, "word/styles.xml": styles})
        _, start, _ = entry_layout(original, "word/styles.xml")
        damaged = bytearray(original)
        damaged[start:start + 8] = b"\x00" * 8
        manifest, _ = run_tool(bytes(damaged), "strict")
        self.assertEqual(manifest["core_content_status"], "complete")
        self.assertEqual(manifest["document_content_status"], "complete")
        self.assertEqual(manifest["overall_status"], "partial")
        self.assertEqual(manifest["auxiliary_failed_entries"], 1)
        self.assertEqual(manifest["content_failed_entries"], 0)

    def test_xlsx_and_pptx_document_content_status(self):
        sheet = (b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
                 b'<row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>')
        shared = b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Cell</t></si></sst>'
        manifest, _ = run_tool(package({"xl/worksheets/sheet1.xml": sheet, "xl/sharedStrings.xml": shared}), "book")
        self.assertEqual(manifest["inferred_type"], ".xlsx")
        self.assertEqual(manifest["document_content_status"], "complete")
        slide = (b'<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                 b'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody>'
                 b'<a:p><a:r><a:t>Slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>')
        notes = slide.replace(b"p:sld", b"p:notes").replace(b"Slide", b"Speaker note")
        manifest, out = run_tool(package({"ppt/slides/slide1.xml": slide, "ppt/notesSlides/notesSlide1.xml": notes}),
                                 "deck")
        self.assertEqual(manifest["document_content_status"], "complete")
        self.assertIn("Speaker note", (out / "notesSlide1.txt").read_text(encoding="utf-8"))

    def test_manifest_records_schema_and_code_hash(self):
        manifest, _ = run_tool(package({"word/document.xml": BODY}), "hash")
        self.assertEqual(manifest["schema_version"], 2)
        self.assertRegex(manifest["tool_code_sha256"], r"^[0-9a-f]{64}$")


if __name__ == "__main__":
    unittest.main()
