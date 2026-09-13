import importlib.util
from io import BytesIO
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest import mock
import zipfile


MODULE = Path(__file__).with_name("legacy-doc-table-style-probes.py")
spec = importlib.util.spec_from_file_location("legacy_doc_table_style_probes", MODULE)
probes = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = probes
spec.loader.exec_module(probes)


class TableStyleProbeTests(unittest.TestCase):
    def test_generator_has_full_table_matrix_seven_styles_and_one_direct_override(self):
        value = probes.build_docx()
        with zipfile.ZipFile(BytesIO(value)) as archive:
            document = archive.read("word/document.xml").decode()
            styles = archive.read("word/styles.xml").decode()
        self.assertEqual(document.count("<w:tbl>"), 8)
        self.assertEqual(document.count('<w:tblLayout w:type="fixed"/>'), 8)
        self.assertEqual(document.count('<w:spacing w:before="0" w:after="0"/>'), 8)
        self.assertEqual(document.count('<w:color w:val="0000FF"/>'), 1)
        self.assertEqual(document.count("<w:tblStyle "), 8)
        self.assertEqual(document.count(" Ω</w:t>"), 16)
        for marker in probes.MARKERS:
            self.assertIn(f"<w:t>{marker}</w:t>", document)
        for ordinal in range(1, 9):
            label = f'<w:t>T{ordinal:02d}</w:t></w:r></w:p><w:tbl>'
            self.assertIn(label, document)
        for style_id in probes.TABLE_STYLES:
            self.assertIn(f'<w:tblStyle w:val="{style_id}"/>', document)
        self.assertEqual(document.count('<w:tblW w:w="9000"'), 8)
        self.assertIn('<w:tblStyle w:val="Base"/><w:tblW', document)
        self.assertIn('<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>', document)
        self.assertIn('<w:sz w:val="24"/>', document)
        self.assertEqual(styles.count('<w:style w:type="table"'), 7)
        self.assertEqual(styles.count("<w:tab "), 7 * 40)
        self.assertIn('w:pos="100"', styles)
        self.assertIn('w:pos="3220"', styles)
        for name, based_on, color, alignment, size, ascii_font, hansi_font in probes.STYLE_DEFINITIONS:
            self.assertIn(f'w:styleId="{name}"', styles)
            if based_on:
                self.assertIn(f'<w:basedOn w:val="{based_on}"/>', styles)
            if color:
                self.assertIn(f'<w:color w:val="{color}"/>', styles)
            if alignment:
                self.assertIn(f'<w:jc w:val="{alignment}"/>', styles)

    def test_generator_is_byte_deterministic(self):
        self.assertEqual(probes.build_docx(), probes.build_docx())

    def test_story_limit_is_checked_before_character_materialization(self):
        fib = {"ccp_text": probes.MAX_PROBE_STORY_UNITS + 1}
        with self.assertRaisesRegex(probes.ProbeError, "character policy"):
            probes._main_characters({"WordDocument": b""}, fib, ())

    def test_probe_loader_requests_the_small_cfb_policy(self):
        papx = probes._papx_module()
        sentinel = object()
        with mock.patch.object(papx, "load_document", return_value=sentinel) as load:
            self.assertIs(probes._load_probe_document("source.doc"), sentinel)
        load.assert_called_once_with(
            "source.doc", max_bytes=probes.MAX_PROBE_CFB_BYTES,
            max_aggregate_bytes=probes.MAX_PROBE_CFB_BYTES,
        )

    def test_chpx_rewrite_removes_only_supported_properties_in_same_range(self):
        keep = bytes.fromhex("000801")
        color = probes.CI_CO.to_bytes(2, "little") + b"\x06"
        size = (0x4A43).to_bytes(2, "little") + b"\x18\x00"
        raw = color + keep + size
        streams = {"WordDocument": bytes((len(raw),)) + raw}
        descriptor = {
            "stream": "WordDocument", "header_offset": 0,
            "offset": 1, "end": len(raw) + 1, "encoding": "character",
        }
        stream, offset, before, after, removed = probes._rewrite_property_payload(
            streams, descriptor, probes.SUPPORTED_CHPX
        )
        self.assertEqual((stream, offset), ("WordDocument", 0))
        self.assertEqual(before, streams["WordDocument"])
        self.assertEqual(len(after), len(before))
        self.assertEqual(after[:1 + len(keep)], bytes((len(keep),)) + keep)
        self.assertEqual(after[1 + len(keep):], bytes(len(raw) - len(keep)))
        self.assertEqual(removed, (probes.CI_CO, 0x4A43))

    def test_papx_rewrite_can_switch_extended_to_short_without_moving_payload(self):
        keep = bytes.fromhex("000801")
        raw = b"\x0b\x00" + probes.P_JC.to_bytes(2, "little") + b"\x02" + keep
        streams = {"WordDocument": b"\x00\x04" + raw}
        descriptor = {
            "stream": "WordDocument", "header_offset": 0,
            "offset": 2, "end": 2 + len(raw), "encoding": "extended",
        }
        _stream, _offset, before, after, removed = probes._rewrite_property_payload(
            streams, descriptor, probes.SUPPORTED_PAPX
        )
        self.assertEqual(len(after), len(before))
        self.assertEqual(after[0], 3)
        self.assertEqual(after[1:6], b"\x0b\x00" + keep)
        self.assertEqual(after[6:], bytes(len(after) - 6))
        self.assertEqual(removed, (probes.P_JC,))

    def test_rewrite_rejects_truncated_property_framing(self):
        raw = probes.C_CV.to_bytes(2, "little") + b"\x01"
        streams = {"WordDocument": bytes((len(raw),)) + raw}
        descriptor = {
            "stream": "WordDocument", "header_offset": 0,
            "offset": 1, "end": len(raw) + 1, "encoding": "character",
        }
        with self.assertRaisesRegex(ValueError, "exceeds its property array"):
            probes._rewrite_property_payload(streams, descriptor, probes.SUPPORTED_CHPX)

    def test_shared_formatting_payload_is_allowed_only_when_every_owner_is_targeted(self):
        run = probes.FormattingRun(10, 20, 100, 101, 106, "character")
        duplicate = probes.FormattingRun(20, 30, 100, 101, 106, "character")
        key = (100, 101, 106, "character")
        probes._assert_unshared(
            {}, {key: {(10, 20), (20, 30)}}, (run, duplicate), "CHPX"
        )
        with self.assertRaisesRegex(probes.ProbeError, "shared"):
            probes._assert_unshared(
                {}, {key: {(20, 30)}}, (run, duplicate), "CHPX"
            )

    def test_t06_body_sharing_a_targeted_mark_is_rejected(self):
        body = probes.FormattingRun(10, 20, 100, 101, 106, "character")
        mark = probes.FormattingRun(20, 21, 100, 101, 106, "character")
        with self.assertRaisesRegex(probes.ProbeError, "untargeted"):
            probes._assert_unshared(
                {}, {(100, 101, 106, "character"): {(20, 21)}},
                (body, mark), "CHPX",
            )

    def test_normalization_rejects_t06_body_and_mark_in_one_run(self):
        owner = (10, 20, 100, 101, 106, "character")
        layout = {"markers": [{
            "preserve_direct": True, "mark_chpx_run": owner,
            "chpx_runs": (owner,),
        }]}
        with self.assertRaisesRegex(probes.ProbeError, "T06 visible body"):
            probes._normalization_edits(layout)

    def test_normalization_records_every_owned_physical_range(self):
        chpx_owner = (10, 20, 100, 101, 104, "character")
        papx = {"stream": "WordDocument", "header_offset": 200,
                "offset": 201, "end": 207, "encoding": "short"}
        layout = {
            "streams": {"WordDocument": bytes(300)},
            "markers": [{
                "preserve_direct": False, "mark_chpx_run": chpx_owner,
                "chpx_runs": (chpx_owner,), "papx": papx,
                "papx_run": (30, 40),
            }],
            "chpx_runs": (probes.FormattingRun(*chpx_owner),),
            "papx_runs": ({"fc_start": 30, "fc_end": 40, "papx": papx},),
        }

        def trace(_streams, _stream, start, _end):
            code = probes.CI_CO if start == 101 else probes.P_JC
            return [{"code": f"{code:04x}"}]

        def rewrite(_streams, descriptor, codes):
            removed = probes.CI_CO if codes == probes.SUPPORTED_CHPX else probes.P_JC
            return ("WordDocument", descriptor["header_offset"], b"x", b"y", (removed,))

        with mock.patch.object(probes, "_prl_segments", side_effect=trace), \
                mock.patch.object(probes, "_rewrite_property_payload", side_effect=rewrite):
            _edits, removed = probes._normalization_edits(layout)
        self.assertEqual(removed["chpx"][0]["owners"], [[10, 20]])
        self.assertEqual(removed["papx"][0]["owners"], [[30, 40]])

    def test_normalization_skips_runs_without_supported_direct_properties(self):
        owner = (10, 20, 100, 101, 104, "character")
        papx = {"stream": "WordDocument", "header_offset": 200,
                "offset": 201, "end": 207, "encoding": "short"}
        layout = {
            "streams": {"WordDocument": bytes(300)},
            "markers": [{
                "preserve_direct": False, "mark_chpx_run": owner,
                "chpx_runs": (owner,), "papx": papx, "papx_run": (30, 40),
            }],
            "chpx_runs": (probes.FormattingRun(*owner),),
            "papx_runs": ({"fc_start": 30, "fc_end": 40, "papx": papx},),
        }
        unsupported = [{"code": "0800"}]
        with mock.patch.object(probes, "_prl_segments", return_value=unsupported), \
                mock.patch.object(probes, "_rewrite_property_payload") as rewrite:
            edits, removed = probes._normalization_edits(layout)
        self.assertEqual(edits, [])
        self.assertEqual(removed, {"chpx": [], "papx": []})
        rewrite.assert_not_called()

    def test_fixed_variant_validation_rejects_replanned_out_of_scope_change(self):
        papx = probes._papx_module()
        source = papx.LoadedDocument({"WordDocument": b"abc"}, b"source")
        expected_streams = {"WordDocument": b"aBc"}
        canonical = {"schema": probes.SCHEMA, "mode": "negative", "edits": [{
            "stream": "WordDocument", "offset": 1, "before": "62", "after": "42",
        }]}
        candidate = {"WordDocument": b"aBD"}
        replanned = {"schema": probes.SCHEMA, "mode": "negative", "edits": [{
            "stream": "WordDocument", "offset": 1, "before": "6263", "after": "4244",
        }]}
        with mock.patch.object(
                probes, "prepare_variants",
                return_value={"negative": (expected_streams, canonical)}):
            with self.assertRaisesRegex(probes.ProbeError, "deterministic"):
                probes.validate_variant(source, candidate, replanned)

    def test_apply_rejects_overlapping_same_length_edits(self):
        streams = {"WordDocument": b"abcdef"}
        edits = [
            ("WordDocument", 1, b"bc", b"BC", None),
            ("WordDocument", 2, b"cd", b"CD", None),
        ]
        with self.assertRaisesRegex(probes.ProbeError, "overlap"):
            probes._apply(streams, edits)

    def _reserved_style(self, name_units=20):
        istd = 17
        metadata = bytearray(10)
        name = "Base" + "_" * (name_units - 4)
        xstz = name_units.to_bytes(2, "little") + name.encode("utf-16le") + b"\0\0"
        tapx = b"\0\0"
        papx_value = istd.to_bytes(2, "little")
        papx = len(papx_value).to_bytes(2, "little") + papx_value
        chpx = b"\0\0"
        record = bytes(metadata) + xstz + tapx + papx + chpx
        metadata[6:8] = len(record).to_bytes(2, "little")
        record = bytes(metadata) + xstz + tapx + papx + chpx
        start = 100
        name_end = start + 10 + len(xstz)
        style = probes.StyleRecord(
            istd, "Base", 0xFFF, 3, start, start + len(record), 10,
            (start + 10, name_end),
            ((name_end + 2, name_end + 2),
             (name_end + 4, name_end + 6),
             (name_end + 8, name_end + 8)),
        )
        return {"1Table": bytes(start) + record}, style

    def test_style_rewrite_preserves_std_size_metadata_and_three_aligned_upxes(self):
        streams, style = self._reserved_style()
        color = probes.CI_CO.to_bytes(2, "little") + b"\x06"
        stream, offset, before, after, metadata = probes._rewrite_style(
            streams, "1Table", style, {"chpx": color.hex()}
        )
        self.assertEqual((stream, offset), ("1Table", style.start))
        self.assertEqual(len(after), len(before))
        self.assertEqual(after[:10], before[:10])
        self.assertEqual(int.from_bytes(after[6:8], "little"), len(after))
        self.assertTrue(metadata["stored_name"].startswith("Base_"))
        self.assertTrue(metadata["stored_name"].isascii())
        self.assertNotEqual(after, before)

    def test_style_rewrite_rejects_wrong_sprm_family(self):
        streams, style = self._reserved_style()
        paragraph_property = probes.P_JC.to_bytes(2, "little") + b"\x01"
        with self.assertRaisesRegex(probes.ProbeError, "wrong SPRM class"):
            probes._rewrite_style(
                streams, "1Table", style, {"chpx": paragraph_property.hex()}
            )

    def test_style_rewrite_rejects_property_sets_larger_than_name_reserve(self):
        streams, style = self._reserved_style(name_units=4)
        color = probes.CI_CO.to_bytes(2, "little") + b"\x06"
        with self.assertRaisesRegex(probes.ProbeError, "insufficient name reserve"):
            probes._rewrite_style(streams, "1Table", style, {"chpx": color.hex()})

    def test_candidate_writer_does_not_delete_a_replacement_path(self):
        papx = probes._papx_module()
        source = papx.LoadedDocument({"WordDocument": b"old"}, b"serialized")
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory, "candidate.doc")

            class FakeHandle:
                def __init__(self, _file, write_mode=False):
                    self.write_mode = write_mode

                def write_stream(self, _name, _value):
                    pass

                def close(self):
                    output.unlink()
                    output.write_bytes(b"replacement")

            fake_ole = SimpleNamespace(OleFileIO=FakeHandle)
            with mock.patch.object(probes.importlib, "import_module", return_value=fake_ole):
                with self.assertRaisesRegex(probes.ProbeError, "path was replaced"):
                    probes._write_doc_candidate(
                        source, output, {"WordDocument": b"changed"}
                    )
            self.assertEqual(output.read_bytes(), b"replacement")


if __name__ == "__main__":
    unittest.main()
