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

    def test_depth_one_cell_uses_cell_mark_and_rejects_ttp_ownership(self):
        characters = [
            {"fc": 10, "character": "x"},
            {"fc": 11, "character": "\x07"},
            {"fc": 12, "character": "\x07"},
        ]
        run = {"fc_start": 10, "fc_end": 12}
        depth = [{"kind": "prl", "applied": True, "code": "6649",
                  "operand": "01000000"}]
        self.assertEqual(
            probes._top_level_cell_mark(characters, run, depth)["fc"], 11
        )
        ttp = depth + [{"kind": "prl", "applied": True, "code": "2417",
                        "operand": "01"}]
        with self.assertRaisesRegex(probes.ProbeError, "not a top-level table cell"):
            probes._top_level_cell_mark(characters, run, ttp)
        with self.assertRaisesRegex(probes.ProbeError, "exactly one cell mark"):
            probes._top_level_cell_mark(characters, {"fc_start": 10, "fc_end": 13}, depth)

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
            "preserve_direct": True, "cell_mark_chpx_run": owner,
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
                "preserve_direct": False, "cell_mark_chpx_run": chpx_owner,
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
                "preserve_direct": False, "cell_mark_chpx_run": owner,
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

    def test_positive_upx_accepts_exact_conditional_border_and_reports_unchecked_values(self):
        border = (0xD47F).to_bytes(2, "little") + b"\x08" + bytes(8)
        conditional = (0xD66A).to_bytes(2, "little") + bytes((2 + len(border),)) \
            + b"\x01\x00" + border
        validation = probes._RecipeValidation("strict-positive")
        probes._validate_style_upx("tapx", conditional, 17, validation, True)
        color = probes.CI_CO.to_bytes(2, "little") + b"\x06"
        probes._validate_style_upx("chpx", color, 17, validation, True)
        report = validation.report()
        self.assertEqual(report["status"], "framing-accepted-with-unchecked-properties")
        self.assertEqual(report["property_sets"], 2)
        self.assertIn(
            {"kind": "chpx", "code": "2a42"}, report["unchecked_properties"]
        )

    def test_diagonal_borders_are_valid_only_inside_tcnf(self):
        for code in probes.OTHER_CONDITIONAL_BORDERS:
            border = code.to_bytes(2, "little") + b"\x08" + bytes(8)
            conditional = (0xD66A).to_bytes(2, "little") \
                + bytes((2 + len(border),)) + b"\x01\x00" + border
            probes._validate_style_upx("tapx", conditional, 17)
            with self.assertRaisesRegex(probes.ProbeError, "prohibited.*UpxTapx"):
                probes._validate_style_upx("tapx", border, 17)

    def test_cell_no_wrap_style_is_not_valid_inside_tcnf(self):
        no_wrap = (0x347D).to_bytes(2, "little") + b"\x01"
        conditional = (0xD66A).to_bytes(2, "little") \
            + bytes((2 + len(no_wrap),)) + b"\x01\x00" + no_wrap
        with self.assertRaisesRegex(probes.ProbeError, "prohibited.*UpxTapx"):
            probes._validate_style_upx("tapx", conditional, 17)
        probes._validate_style_upx("tapx", no_wrap, 17)

    def test_default_table_style_requires_unconditional_zero_width_before(self):
        width_before = (0xF617).to_bytes(2, "little") + b"\x03\x00\x00"
        conditional = (0xD66A).to_bytes(2, "little") \
            + bytes((2 + len(width_before),)) + b"\x01\x00" + width_before
        with self.assertRaisesRegex(probes.ProbeError, "lacks required"):
            probes._validate_style_upx("tapx", b"", 0x000B)
        with self.assertRaisesRegex(probes.ProbeError, "lacks required"):
            probes._validate_style_upx("tapx", conditional, 0x000B)
        probes._validate_style_upx("tapx", conditional + width_before, 0x000B)

        negative = probes._RecipeValidation("spec-invalid-negative-control")
        probes._validate_style_upx("tapx", b"", 0x000B, negative, True)
        self.assertIn(
            {"kind": "tapx", "code": "f617"},
            negative.report()["spec_invalid"],
        )

        # Untouched Word-produced property sets are not recertified by this
        # replacement validator.
        probes._validate_style_upx("tapx", b"", 0x000B, deep=False)

    def test_positive_upx_rejects_bad_condition_and_nested_condition(self):
        wrapper = (0xCA85).to_bytes(2, "little")
        bad_condition = wrapper + b"\x02\x03\x00"
        with self.assertRaisesRegex(probes.ProbeError, "CNF condition"):
            probes._validate_style_upx("chpx", bad_condition, 17)
        nested = wrapper + b"\x07\x01\x00" + wrapper + b"\x02\x01\x00"
        with self.assertRaisesRegex(probes.ProbeError, "nested CNF"):
            probes._validate_style_upx("chpx", nested, 17)

    def test_all_twelve_cnf_conditions_use_exact_nested_category(self):
        color = probes.CI_CO.to_bytes(2, "little") + b"\x06"
        for condition in probes.CNF_CONDITIONS:
            wrapper = (0xCA85).to_bytes(2, "little") \
                + bytes((2 + len(color),)) + condition.to_bytes(2, "little") + color
            probes._validate_style_upx("chpx", wrapper, 17)
        wrong = (0xCA85).to_bytes(2, "little") + b"\x05\x01\x00" \
            + probes.P_JC.to_bytes(2, "little") + b"\x01"
        with self.assertRaisesRegex(probes.ProbeError, "wrong SPRM class"):
            probes._validate_style_upx("chpx", wrong, 17)

    def test_conditional_border_requires_exact_brc_operand(self):
        border = (0xD47F).to_bytes(2, "little") + b"\x07" + bytes(7)
        conditional = (0xD66A).to_bytes(2, "little") + bytes((2 + len(border),)) \
            + b"\x01\x00" + border
        with self.assertRaisesRegex(probes.ProbeError, "BrcOperand cb must be 8"):
            probes._validate_style_upx("tapx", conditional, 17)
        invalid_type = (0xD47F).to_bytes(2, "little") + b"\x08" \
            + bytes(5) + b"\x02" + bytes(2)
        wrapped = (0xD66A).to_bytes(2, "little") + bytes((2 + len(invalid_type),)) \
            + b"\x01\x00" + invalid_type
        with self.assertRaisesRegex(probes.ProbeError, "invalid table border type"):
            probes._validate_style_upx("tapx", wrapped, 17)
        nil_border = (0xD47F).to_bytes(2, "little") + b"\x08" \
            + bytes(4) + b"\xff" * 4
        nil_wrapped = (0xD66A).to_bytes(2, "little") \
            + bytes((2 + len(nil_border),)) + b"\x01\x00" + nil_border
        with self.assertRaisesRegex(probes.ProbeError, "not a normative BrcOperand"):
            probes._validate_style_upx("tapx", nil_wrapped, 17)
        negative = probes._RecipeValidation("spec-invalid-negative-control")
        probes._validate_style_upx("tapx", nil_wrapped, 17, negative, True)
        self.assertIn(
            {"kind": "tapx", "code": "d47f"},
            negative.report()["spec_invalid"],
        )

    def test_cssa_requires_exact_size_range_sides_and_units(self):
        valid = (0xD63E).to_bytes(2, "little") + b"\x06\x00\x01\x0f\x03\x6c\x00"
        probes._validate_style_upx("tapx", valid, 17)
        for code, operand, message in (
            (0xD63E, b"\x05" + bytes(5), "CSSAOperand cb must be 6"),
            (0xD63E, b"\x06\x02\x01\x0f\x03\x00\x00", "CSSA cell range"),
            (0xD63E, b"\x06\x00\x01\x10\x03\x00\x00", "CSSA side mask"),
            (0xD634, b"\x06\x00\x01\x0f\x00\x01\x00", "ftsNil width"),
        ):
            with self.subTest(message=message), self.assertRaisesRegex(probes.ProbeError, message):
                probes._validate_style_upx(
                    "tapx", code.to_bytes(2, "little") + operand, 17
                )
        at_limit = b"\x06\x00\x01\x0f\x03" + (31_680).to_bytes(2, "little")
        probes._validate_style_upx(
            "tapx", (0xD634).to_bytes(2, "little") + at_limit, 17
        )
        above = at_limit[:-2] + (31_681).to_bytes(2, "little")
        with self.assertRaisesRegex(probes.ProbeError, "width exceeds"):
            probes._validate_style_upx(
                "tapx", (0xD634).to_bytes(2, "little") + above, 17
            )

    def test_raw_shading_is_bounded_and_requires_explicit_negative_mode(self):
        raw = (0xD670).to_bytes(2, "little") + b"\x0a" + bytes(10)
        with self.assertRaisesRegex(probes.ProbeError, "prohibited.*UpxTapx"):
            probes._validate_style_upx("tapx", raw, 17)
        negative = probes._RecipeValidation("spec-invalid-negative-control")
        probes._validate_style_upx("tapx", raw, 17, negative, True)
        self.assertEqual(negative.report()["status"], "spec-invalid-negative-control")
        malformed = (0xD670).to_bytes(2, "little") + b"\x0b" + bytes(11)
        with self.assertRaisesRegex(probes.ProbeError, "RawShd.*multiple of 10"):
            probes._validate_style_upx("tapx", malformed, 17, negative, True)
        for code, size in ((0xD670, 220), (0xD672, 190)):
            probes._validate_style_upx(
                "tapx", code.to_bytes(2, "little") + bytes((size,)) + bytes(size),
                17, probes._RecipeValidation("spec-invalid-negative-control"), True,
            )
        too_many = (0xD672).to_bytes(2, "little") + b"\xc8" + bytes(200)
        with self.assertRaisesRegex(probes.ProbeError, "cell limit"):
            probes._validate_style_upx(
                "tapx", too_many, 17,
                probes._RecipeValidation("spec-invalid-negative-control"), True,
            )

    def test_recipe_validation_budget_is_shared_across_property_sets(self):
        color = probes.CI_CO.to_bytes(2, "little") + b"\x06"
        validation = probes._RecipeValidation("strict-positive", byte_limit=5)
        probes._validate_style_upx("chpx", color, 17, validation, True)
        with self.assertRaisesRegex(probes.ProbeError, "recipe validation byte budget"):
            probes._validate_style_upx("chpx", color, 17, validation, True)

    def test_recipe_validation_bounds_records_before_a_second_prl(self):
        color = probes.CI_CO.to_bytes(2, "little") + b"\x06"
        validation = probes._RecipeValidation("strict-positive", prl_limit=1)
        with self.assertRaisesRegex(probes.ProbeError, "recipe validation SPRM budget"):
            probes._validate_style_upx("chpx", color + color, 17, validation, True)

    def test_papx_owner_and_indirection_are_validated_for_replacements(self):
        with self.assertRaisesRegex(probes.ProbeError, "own istd"):
            probes._validate_style_upx("papx", b"\x12\x00", 17)
        indirect = b"\x11\x00" + (0x646B).to_bytes(2, "little") + bytes(4)
        with self.assertRaisesRegex(probes.ProbeError, "indirect SPRM"):
            probes._validate_style_upx("papx", indirect, 17)

    def test_negative_mode_does_not_admit_malformed_nested_operands(self):
        border = (0xD47F).to_bytes(2, "little") + b"\x07" + bytes(7)
        conditional = (0xD66A).to_bytes(2, "little") + bytes((2 + len(border),)) \
            + b"\x01\x00" + border
        validation = probes._RecipeValidation("spec-invalid-negative-control")
        with self.assertRaisesRegex(probes.ProbeError, "BrcOperand"):
            probes._validate_style_upx("tapx", conditional, 17, validation, True)

    def test_recipe_rejects_an_unnamed_validation_mode_before_inspection(self):
        with self.assertRaisesRegex(probes.ProbeError, "validation mode"):
            probes.build_recipe_variant(None, {
                "schema": probes.RECIPE_SCHEMA,
                "validation_mode": "relaxed",
            })

    def test_historical_recipe_plan_replays_with_new_validation_unrecorded(self):
        papx = probes._papx_module()
        source = papx.LoadedDocument({"WordDocument": b"source"}, b"source")
        candidate = {"WordDocument": b"result"}
        historical = {
            "schema": probes.SCHEMA,
            "mode": "recipe",
            "recipe": {"schema": probes.RECIPE_SCHEMA},
            "source_sha256": source.source_sha256,
            "edits": [],
            "targets": [],
        }
        rebuilt = dict(historical)
        rebuilt["recipe_validation"] = {
            "status": "framing-accepted-with-unchecked-properties"
        }
        with mock.patch.object(
                probes, "build_recipe_variant", return_value=(candidate, rebuilt)):
            result = probes.validate_recipe_variant(source, candidate, historical)
        self.assertTrue(result["valid"])
        self.assertFalse(result["recipe_validation_recorded"])
        self.assertEqual(result["recipe_validation"], rebuilt["recipe_validation"])

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
