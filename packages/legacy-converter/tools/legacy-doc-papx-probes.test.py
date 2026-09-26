import importlib.util
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch


MODULE = Path(__file__).with_name("legacy-doc-papx-probes.py")
spec = importlib.util.spec_from_file_location("legacy_doc_papx_probes", MODULE)
probes = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probes)


def prl(code, operand):
    return code.to_bytes(2, "little") + bytes(operand)


def prc_data(grpprl):
    assert 10 <= len(grpprl) <= probes.MAX_PRC_GRPPRL_BYTES
    return len(grpprl).to_bytes(2, "little") + grpprl


def prm0(isprm, value):
    return (isprm << 1) | (value << 8)


def top_level_row(grpprl):
    return (
        grpprl
        + prl(0x6649, (1).to_bytes(4, "little"))
        + prl(0x2417, b"\x01")
    )


def owned_top_level_row(grpprl):
    return prl(0x2416, b"\x01") + top_level_row(grpprl)


def fixture(papx, data=b"", prcs=(), prm=0):
    word = bytearray(1024)
    word[0:2] = probes.FIB_IDENT.to_bytes(2, "little")
    word[2:4] = probes.FIB_WORD_97.to_bytes(2, "little")
    word[probes.CCP_TEXT_OFFSET:probes.CCP_TEXT_OFFSET + 4] = (10).to_bytes(4, "little")
    word[
        probes.FC_PLCF_BTE_PAPX_OFFSET:probes.FC_PLCF_BTE_PAPX_OFFSET + 4
    ] = (0).to_bytes(4, "little")
    word[
        probes.LCB_PLCF_BTE_PAPX_OFFSET:probes.LCB_PLCF_BTE_PAPX_OFFSET + 4
    ] = (12).to_bytes(4, "little")
    word[105] = 0x07

    page = memoryview(word)[512:1024]
    page[0:4] = (100).to_bytes(4, "little")
    page[4:8] = (110).to_bytes(4, "little")
    if papx:
        page[8] = 32
        if len(papx) % 2:
            page[64] = (len(papx) + 1) // 2
            page[65:65 + len(papx)] = papx
        else:
            page[64] = 0
            page[65] = len(papx) // 2
            page[66:66 + len(papx)] = papx
    page[511] = 1

    clx = bytearray()
    for properties in prcs:
        clx.append(0x01)
        clx.extend(len(properties).to_bytes(2, "little"))
        clx.extend(properties)
    plc = bytearray()
    plc.extend((0).to_bytes(4, "little"))
    plc.extend((10).to_bytes(4, "little"))
    plc.extend(b"\0\0")
    plc.extend(((100 * 2) | 0x40000000).to_bytes(4, "little"))
    plc.extend(prm.to_bytes(2, "little"))
    clx.append(0x02)
    clx.extend(len(plc).to_bytes(4, "little"))
    clx.extend(plc)

    table = bytearray()
    table.extend((100).to_bytes(4, "little"))
    table.extend((110).to_bytes(4, "little"))
    table.extend((1).to_bytes(4, "little"))
    fc_clx = len(table)
    table.extend(clx)
    word[probes.FC_CLX_OFFSET:probes.FC_CLX_OFFSET + 4] = fc_clx.to_bytes(4, "little")
    word[probes.LCB_CLX_OFFSET:probes.LCB_CLX_OFFSET + 4] = len(clx).to_bytes(4, "little")
    return {"WordDocument": bytes(word), "0Table": bytes(table), "Data": bytes(data)}


def loaded(streams, source=b"serialized source"):
    return probes.LoadedDocument(streams, source)


def changed(streams, stream, offset, replacement):
    result = dict(streams)
    value = bytearray(result[stream])
    value[offset:offset + len(replacement)] = replacement
    result[stream] = bytes(value)
    return result


class PapxProbeTests(unittest.TestCase):
    def trace_manifest(self, document, targets):
        return {
            "schema": probes.TRACE_ASSERTION_SCHEMA,
            "source_sha256": document.source_sha256,
            "targets": targets,
        }

    def test_trace_assertions_require_exact_operands_absence_and_order(self):
        props = owned_top_level_row(
            prl(0xD635, b"\x02ab")
            + prl(0x7621, b"\x01\0\0\0")
            + prl(0xD635, b"\x02cd")
        )
        document = loaded(fixture(b"\0\0" + props))
        target = {
            "fc": 105,
            "owner": "ttp",
            "properties": {
                "d635": ["026162", "026364"],
                "7621": ["01000000"],
                "7623": [],
            },
            "order": ["d635", "7621", "d635"],
        }
        report = probes.validate_trace_assertions(
            document, self.trace_manifest(document, [target])
        )
        self.assertEqual(report["targets"], [target])
        false_bit_presence = {**target, "properties": {
            **target["properties"], "7621": ["00000000"],
        }}
        with self.assertRaisesRegex(probes.ProbeError, "operand assertion"):
            probes.validate_trace_assertions(
                document, self.trace_manifest(document, [false_bit_presence])
            )
        with self.assertRaisesRegex(probes.ProbeError, "order assertion"):
            probes.validate_trace_assertions(
                document, self.trace_manifest(document, [{**target, "order": [
                    "7621", "d635", "d635",
                ]}])
            )

    def test_direct_ptableprops_ignores_its_tail_and_the_piece_redirect(self):
        ignored = prl(0x7623, b"\x03\0\0\0")
        indirect = prl(probes.P_TABLE_PROPS, b"\0\0\0\0") + ignored
        data = prc_data(owned_top_level_row(prl(0x7621, b"\x01\0\0\0")))
        pcd = prl(probes.P_TABLE_PROPS, b"\0\0\0\0")
        document = loaded(fixture(b"\0\0" + indirect, data, [pcd], prm=1))
        target = {
            "fc": 105,
            "owner": "ttp",
            "properties": {
                "7621": ["01000000"],
                "7623": [],
            },
            "order": ["7621"],
        }
        probes.validate_trace_assertions(
            document, self.trace_manifest(document, [target])
        )
        acquired = probes.acquire_property_trace(document, 105, "ttp")
        ignored_items = [item for item in acquired["entries"] if item["kind"] == "ignored_tail"]
        self.assertEqual(len(ignored_items), 1)
        self.assertIn(ignored.hex(), ignored_items[0]["bytes"])

    def test_simple_prm0_and_inline_piece_follow_the_direct_redirect_chain(self):
        data = prc_data(owned_top_level_row(prl(0x7621, b"\x01\0\0\0")))
        direct = b"\0\0" + prl(probes.P_TABLE_PROPS, b"\0\0\0\0")
        for streams in (
            fixture(direct, data, prm=prm0(0x05, 2)),
            fixture(direct, data, [prl(0x2461, b"\x02")], prm=1),
        ):
            acquired = probes.acquire_property_trace(loaded(streams), 105, "ttp")["entries"]
            self.assertEqual(
                [entry["operand"] for entry in acquired
                 if entry.get("code") == "2461" and entry.get("applied")],
                ["02"],
            )

    def test_phuge_in_appended_piece_is_never_followed(self):
        data = prc_data(owned_top_level_row(prl(0x2461, b"\x02")))
        direct = b"\0\0" + owned_top_level_row(prl(0x2461, b"\x00"))
        piece = prl(probes.PHUGE_PAPX, b"\0\0\0\0")
        document = loaded(fixture(direct, data, [piece], prm=1))
        acquired = probes.acquire_property_trace(document, 105, "ttp")["entries"]
        self.assertEqual(
            [entry["operand"] for entry in acquired
             if entry.get("code") == "2461" and entry.get("applied")],
            ["00"],
        )
        huge = next(entry for entry in acquired if entry.get("code") == "6646")
        self.assertEqual((huge["applied"], huge["reason"]),
                         (False, "complex-PCD-redirect"))

    def test_complex_pcd_keeps_table_sprms_and_never_follows_redirects(self):
        data = prc_data(owned_top_level_row(prl(0x7621, b"\0\x01\xe8\x03")))
        table = prl(0x7623, b"\0\x01\xe8\x03")
        bold = prl(0x0835, b"\x01")
        tail = prl(0x2461, b"\x02")
        for redirect in (probes.PHUGE_PAPX, probes.P_TABLE_PROPS):
            piece = prl(redirect, b"\0\0\0\0") + table + bold + tail
            streams = fixture(b"\0\0", data, [piece], prm=1)
            record = probes.inspect_document(streams)["prcs"][0]
            trace = probes._trace_properties(
                streams, record["stream"], record["offset"], record["end"],
                probes._TraceBudget(), filter_initial_paragraph=True,
            )
            self.assertEqual(
                [(entry["code"], entry["applied"], entry.get("reason"))
                 for entry in trace],
                [(f"{redirect:04x}", False, "complex-PCD-redirect"),
                 ("7623", True, None),
                 ("0835", False, "non-paragraph-complex-PCD"),
                 ("2461", True, None)],
            )

    def test_raw_and_complex_piece_trace_cache_entries_do_not_alias(self):
        streams = fixture(b"\0\0", prcs=[prl(0x0835, b"\x01")], prm=1)
        session = probes.AcquiredPropertyTraceSession(loaded(streams))
        record = session.prcs[0]
        key = (record["stream"], record["offset"], record["end"])
        raw = session._trace(key)
        filtered = session._acquired_trace(None, key)
        self.assertTrue(raw[0]["applied"])
        self.assertFalse(filtered[0]["applied"])
        self.assertEqual(len(session.trace_cache), 2)

    def test_trace_owner_distinguishes_cell_mark_from_top_level_ttp(self):
        cell_props = (
            prl(0x2416, b"\x01")
            + prl(0x6649, (2).to_bytes(4, "little"))
        )
        cell = loaded(fixture(b"\0\0" + cell_props))
        probes.validate_trace_assertions(cell, self.trace_manifest(cell, [{
            "fc": 105, "owner": "paragraph", "properties": {"6649": ["02000000"]},
        }]))
        with self.assertRaisesRegex(probes.ProbeError, "top-level TTP ownership"):
            probes.validate_trace_assertions(cell, self.trace_manifest(cell, [{
                "fc": 105, "owner": "ttp", "properties": {"6649": ["02000000"]},
            }]))
        ttp = loaded(fixture(b"\0\0" + owned_top_level_row(b"")))
        with self.assertRaisesRegex(probes.ProbeError, "paragraph owner rejects"):
            probes.validate_trace_assertions(ttp, self.trace_manifest(ttp, [{
                "fc": 105, "owner": "paragraph", "properties": {"2417": ["01"]},
            }]))

    def test_trace_schema_fc_boundaries_duplicates_and_budgets_fail_closed(self):
        document = loaded(fixture(b"\0\0" + owned_top_level_row(b"")))
        target = {"fc": 105, "owner": "ttp", "properties": {"2417": ["01"]}}
        manifest = self.trace_manifest(document, [target])
        for bad, message in [
            ({**manifest, "extra": 1}, "unknown fields"),
            ({**manifest, "targets": []}, "requires targets"),
            ({**manifest, "targets": [{**target, "fc": True}]}, "nonnegative integer"),
            ({**manifest, "targets": [{**target, "fc": -1}]}, "nonnegative integer"),
            ({**manifest, "targets": [target, target]}, "duplicate trace target"),
            ({**manifest, "targets": [{**target, "properties": {"2417": ["1"]}}]},
             "whole-byte hex"),
            ({**manifest, "targets": [{**target, "properties": {"7621": ["01"]}}]},
             "SPRM framing"),
            ({**manifest, "targets": [{**target, "fc": 110}]}, "PAPX run"),
        ]:
            with self.assertRaisesRegex(probes.ProbeError, message):
                probes.validate_trace_assertions(document, bad)

        uncompressed = dict(document.streams)
        table = bytearray(uncompressed["0Table"])
        table[27:31] = (100).to_bytes(4, "little")
        uncompressed["0Table"] = bytes(table)
        wide = loaded(uncompressed)
        with self.assertRaisesRegex(probes.ProbeError, "character boundary"):
            probes.validate_trace_assertions(wide, self.trace_manifest(wide, [{
                **target, "fc": 105,
            }]))
        with patch.object(probes, "MAX_TARGETS", 1):
            with self.assertRaisesRegex(probes.ProbeError, "target count"):
                probes.validate_trace_assertions(document, {
                    **manifest, "targets": [target, {**target, "fc": 106}],
                })
        with patch.object(probes, "MAX_TRACE_BYTES", 1):
            with self.assertRaisesRegex(probes.ProbeError, "retained output"):
                probes.validate_trace_assertions(document, manifest)

        no_in_table = loaded(fixture(b"\0\0" + top_level_row(b"")))
        with self.assertRaisesRegex(probes.ProbeError, "top-level TTP ownership"):
            probes.validate_trace_assertions(no_in_table, self.trace_manifest(
                no_in_table, [target]
            ))
        with patch.object(probes, "MAX_TRACE_BYTES", 128):
            oversized = {**target, "properties": {"d635": ["00" * 129]}}
            with self.assertRaisesRegex(probes.ProbeError, "retained output"):
                probes.validate_trace_assertions(
                    document, self.trace_manifest(document, [oversized])
                )

    def test_repeated_cached_trace_acquisition_is_aggregate_work_bounded(self):
        streams = fixture(b"\0\0" + owned_top_level_row(b""))
        word = bytearray(streams["WordDocument"])
        word[100:110] = b"\x07" * 10
        document = loaded({**streams, "WordDocument": bytes(word)})
        targets = [
            {"fc": fc, "owner": "ttp", "properties": {"2417": ["01"]}}
            for fc in range(100, 110)
        ]
        with patch.object(probes, "MAX_INSPECT_WORK", 40):
            with self.assertRaisesRegex(probes.ProbeError, "work budget"):
                probes.validate_trace_assertions(
                    document, self.trace_manifest(document, targets)
                )

    def test_validator_can_reuse_a_document_bound_trace_session(self):
        document = loaded(fixture(b"\0\0" + owned_top_level_row(b"")))
        target = {"fc": 105, "owner": "ttp", "properties": {"2417": ["01"]}}
        session = probes.AcquiredPropertyTraceSession(document)
        probes.validate_trace_assertions(
            document, self.trace_manifest(document, [target]), session=session
        )
        other = loaded(dict(document.streams), source=b"other source")
        with self.assertRaisesRegex(probes.ProbeError, "not bound"):
            probes.validate_trace_assertions(
                other, self.trace_manifest(other, [target]), session=session
            )

    def test_paragraph_prm0_appends_canonical_pjc_with_exact_origin(self):
        direct = prl(0x2461, b"\x01")
        document = loaded(fixture(b"\0\0" + direct, prm=prm0(0x05, 2)))
        acquired = probes.acquire_property_trace(document, 105, "paragraph")
        pjc = [item for item in acquired["entries"] if item["code"] == "2461"]
        self.assertEqual([item["operand"] for item in pjc], ["01", "02"])
        synthetic = pjc[-1]
        piece = probes.inspect_document(document)["pieces"][0]
        self.assertEqual(synthetic, {
            "kind": "prl", "stream": "0Table",
            "offset": piece["offset"] + 6, "end": piece["offset"] + 8,
            "code": "2461", "operand": "02", "applied": True,
            "origin": "Pcd.Prm0",
            "framing": "synthetic-Prl-from-2-byte-Prm0",
            "prm": f"{prm0(0x05, 2):04x}", "isprm": "05",
        })
        probes.validate_trace_assertions(document, self.trace_manifest(document, [{
            "fc": 105, "owner": "paragraph",
            "properties": {"2461": ["01", "02"]},
            "order": ["2461", "2461"],
        }]))

    def test_paragraph_prm0_structural_flags_participate_in_ownership(self):
        base = prl(0x6649, (1).to_bytes(4, "little")) + prl(0x2416, b"\x01")
        ttp = loaded(fixture(b"\0\0" + base, prm=prm0(0x19, 1)))
        acquired = probes.acquire_property_trace(ttp, 105, "ttp")
        self.assertEqual(acquired["entries"][-1]["code"], "2417")

        inconsistent = loaded(fixture(
            b"\0\0" + base, prm=prm0(0x18, 0)
        ))
        with self.assertRaisesRegex(probes.ProbeError, "inconsistent acquired in-table"):
            probes.acquire_property_trace(inconsistent, 105, "paragraph")

    def test_nonparagraph_and_no_effect_prm0_do_not_become_paragraph_properties(self):
        character = loaded(fixture(b"\0\0", prm=prm0(0x55, 1)))
        entry = probes.acquire_property_trace(character, 105, "paragraph")["entries"][-1]
        self.assertEqual(
            (entry["code"], entry["operand"], entry["applied"], entry["reason"]),
            ("0835", "01", False, "non-paragraph-Prm0"),
        )
        probes.validate_trace_assertions(character, self.trace_manifest(character, [{
            "fc": 105, "owner": "paragraph", "properties": {"0835": []},
        }]))
        line_break = loaded(fixture(b"\0\0", prm=prm0(0x00, 1)))
        entry = probes.acquire_property_trace(line_break, 105, "paragraph")["entries"][-1]
        self.assertEqual((entry["code"], entry["applied"]), ("2879", False))

        no_effect = loaded(fixture(b"\0\0", prm=prm0(0x00, 0)))
        self.assertEqual(
            probes.acquire_property_trace(no_effect, 105, "paragraph")["entries"], []
        )

    def test_prm0_tables_match_the_documented_closed_mapping(self):
        self.assertEqual(probes.PRM0_PARAGRAPH_SPRMS, {
            0x04: 0x2602, 0x05: 0x2461, 0x07: 0x2405, 0x08: 0x2406,
            0x09: 0x2407, 0x0C: 0x260A, 0x0D: 0x2470, 0x0E: 0x240C,
            0x0F: 0x2471, 0x18: 0x2416, 0x19: 0x2417, 0x1D: 0x261B,
            0x25: 0x2423, 0x2C: 0x242A, 0x32: 0x2430, 0x33: 0x2431,
            0x35: 0x2433, 0x36: 0x2434, 0x37: 0x2435, 0x38: 0x2436,
            0x39: 0x2437, 0x3A: 0x2438, 0x78: 0x2640, 0x7E: 0x2443,
        })
        self.assertEqual(probes.PRM0_CHARACTER_SPRMS, {
            0x00: 0x2879, 0x41: 0x0800, 0x42: 0x0801, 0x43: 0x0802,
            0x47: 0x0806, 0x4B: 0x080A, 0x4D: 0x2A0C, 0x4E: 0x0858,
            0x4F: 0x2859, 0x50: 0x0811, 0x51: 0x0818, 0x53: 0x2A33,
            0x55: 0x0835, 0x56: 0x0836, 0x57: 0x0837, 0x58: 0x0838,
            0x59: 0x0839, 0x5A: 0x083A, 0x5B: 0x083B, 0x5C: 0x083C,
            0x5E: 0x2A3E, 0x62: 0x2A42, 0x68: 0x2A48, 0x73: 0x2A53,
            0x74: 0x0854, 0x75: 0x0855, 0x76: 0x0856, 0x7B: 0x2A90,
            0x7C: 0x2A86,
        })

    def test_prm0_reserved_indices_and_budgets_fail_closed(self):
        reserved = loaded(fixture(b"\0\0", prm=prm0(0x01, 1)))
        with self.assertRaisesRegex(probes.ProbeError, "reserved or unknown Prm0"):
            probes.acquire_property_trace(reserved, 105, "paragraph")

        paragraph = loaded(fixture(b"\0\0", prm=prm0(0x05, 1)))
        with patch.object(probes, "MAX_TRACE_BYTES", 1):
            with self.assertRaisesRegex(probes.ProbeError, "byte budget"):
                probes.acquire_property_trace(paragraph, 105, "paragraph")

        with self.assertRaisesRegex(probes.ProbeError, "complex PCD property index"):
            probes.acquire_property_trace(
                loaded(fixture(b"\0\0", prcs=[b""], prm=3)), 105, "paragraph"
            )

    def test_physical_interval_lookup_accepts_reordered_and_detects_overlap(self):
        session = probes.AcquiredPropertyTraceSession.__new__(
            probes.AcquiredPropertyTraceSession
        )
        session.lookup_budget = probes._WorkBudget(probes.MAX_TARGET_LOOKUP_WORK)
        reordered = [
            {"fc_start": 200, "fc_end": 210, "name": "later"},
            {"fc_start": 100, "fc_end": 110, "name": "earlier"},
        ]
        index = session._interval_index(reordered)
        self.assertEqual(session._containing(index, 105, "piece")["name"], "earlier")
        overlapping = [
            {"fc_start": 100, "fc_end": 300},
            {"fc_start": 200, "fc_end": 210},
        ]
        with self.assertRaisesRegex(probes.ProbeError, "exactly one piece"):
            session._containing(session._interval_index(overlapping), 205, "piece")

    def test_trace_json_loader_rejects_duplicate_object_fields(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory, "assertions.json")
            path.write_text('{"schema":"a","schema":"b"}', encoding="utf-8")
            with self.assertRaisesRegex(probes.ProbeError, "duplicate JSON field schema"):
                probes._load_plan(path)

    def test_inspector_records_both_papx_encodings_and_ignored_tail_provenance(self):
        indirect = prl(probes.P_TABLE_PROPS, (0).to_bytes(4, "little"))
        compatibility = prl(0xD608, bytes([6, 0, 1, 0, 0, 0xD0, 7]))
        data = prc_data(top_level_row(
            prl(0x563A, (11).to_bytes(2, "little")) + prl(0x740A, b"\0\0 \0")
        ))
        short = fixture(b"\0\0" + indirect + compatibility, data)
        inspected = probes.inspect_document(short)
        run = inspected["papx_runs"][0]
        self.assertNotIn("pcds", run)
        self.assertEqual(len(inspected["pieces"]), 1)
        self.assertEqual(run["papx"]["encoding"], "short")
        self.assertEqual(run["papx"]["istd"], 0)
        trace = inspected["property_traces"][run["direct_trace_id"]]["entries"]
        self.assertEqual(
            [(item["kind"], item["stream"]) for item in trace],
            [("prl", "WordDocument"), ("ignored_tail", "WordDocument"),
             ("prc_data", "Data"), ("prl", "Data"), ("prl", "Data"),
             ("prl", "Data"), ("prl", "Data")],
        )
        ignored = trace[1]
        self.assertEqual(ignored["bytes"], compatibility.hex())
        self.assertEqual(ignored["offset"], run["papx"]["offset"] + 2 + len(indirect))

        extended = probes.inspect_document(fixture(b"\0\0" + indirect, data))["papx_runs"][0]
        self.assertEqual(extended["papx"]["encoding"], "extended")
        self.assertEqual(extended["papx"]["end"] - extended["papx"]["offset"], 8)

    def test_active_data_replacement_with_absent_before_target_validates(self):
        rsid = prl(0x6467, b"\x11\x22\x33\x44")
        ttlp = prl(0x740A, b"\0\0 \0")
        before_data = prc_data(top_level_row(prl(0x563A, (11).to_bytes(2, "little")) + rsid))
        before = fixture(b"\0\0" + prl(probes.P_TABLE_PROPS, b"\0\0\0\0"), before_data)
        offset = 2 + len(prl(0x563A, b"\0\0"))
        after = changed(before, "Data", offset, ttlp)
        source = loaded(before)
        plan = {
            "source_sha256": source.source_sha256,
            "edits": [{"stream": "Data", "offset": offset,
                       "before": rsid.hex(), "after": ttlp.hex()}],
            "targets": [{"fc": 105, "code": "TTlp", "before": None,
                         "after": "00002000"}],
        }
        result = probes.validate_plan(source, after, plan)
        self.assertTrue(result["valid"])
        self.assertIsNone(result["targets"][0]["before_source"])
        self.assertEqual(result["targets"][0]["after_source"]["stream"], "Data")

    def test_fkp_compatibility_tail_only_edit_is_not_an_effective_target(self):
        rsid = prl(0x6467, b"\x11\x22\x33\x44")
        ttlp = prl(0x740A, b"\0\0 \0")
        data = prc_data(top_level_row(prl(0x563A, (11).to_bytes(2, "little")) + rsid))
        before = fixture(
            b"\0\0" + prl(probes.P_TABLE_PROPS, b"\0\0\0\0") + rsid,
            data,
        )
        papx = probes.inspect_document(before)["papx_runs"][0]["papx"]
        offset = papx["offset"] + 2 + len(prl(probes.P_TABLE_PROPS, b"\0\0\0\0"))
        after = changed(before, "WordDocument", offset, ttlp)
        source = loaded(before)
        plan = {
            "source_sha256": source.source_sha256,
            "edits": [{"stream": "WordDocument", "offset": offset,
                       "before": rsid.hex(), "after": ttlp.hex()}],
            "targets": [{"fc": 105, "code": "TTlp", "before": None,
                         "after": "00002000"}],
        }
        with self.assertRaisesRegex(probes.ProbeError, "effective target"):
            probes.validate_plan(source, after, plan)

    def test_complex_pcd_table_sprm_shadows_direct_data_target(self):
        before_value = b"\0\0 \0"
        after_value = b"\0\0@\0"
        pcd_value = b"\0\0\x80\0"
        before_data = prc_data(top_level_row(
            prl(0x563A, (11).to_bytes(2, "little")) + prl(0x740A, before_value)
        ))
        complex_properties = prl(0x740A, pcd_value)
        papx = b"\0\0" + prl(probes.P_TABLE_PROPS, b"\0\0\0\0")
        before = fixture(papx, before_data, [complex_properties], prm=1)
        offset = 2 + len(prl(0x563A, b"\0\0")) + 2
        after = changed(before, "Data", offset, after_value)
        source = loaded(before)
        plan = {
            "source_sha256": source.source_sha256,
            "edits": [{"stream": "Data", "offset": offset,
                       "before": before_value.hex(), "after": after_value.hex()}],
            "targets": [{"fc": 105, "code": 0x740A,
                         "before": before_value.hex(), "after": after_value.hex()}],
        }
        # The piece TTlp applies after the direct PrcData, so the Data edit is
        # not the effective value and the plan must be rejected.
        with self.assertRaisesRegex(probes.ProbeError, "does not match"):
            probes.validate_plan(source, after, plan)
        inspected = probes.inspect_document(before)
        record = inspected["prcs"][0]
        raw = probes._trace_properties(
            before, record["stream"], record["offset"], record["end"],
            probes._TraceBudget(), filter_initial_paragraph=True,
        )[-1]
        self.assertEqual(
            (raw["code"], raw["applied"], raw.get("reason")),
            ("740a", True, None),
        )

    def test_wrapped_complex_pcd_redirect_does_not_shadow_direct_data(self):
        before_value = b"\0\0 \0"
        after_value = b"\0\0@\0"
        pcd_value = b"\0\0\x80\0"
        direct = prc_data(owned_top_level_row(prl(0x740A, before_value)))
        pcd_offset = len(direct)
        pcd_data = prc_data(owned_top_level_row(prl(0x740A, pcd_value)))
        papx = b"\0\0" + prl(probes.P_TABLE_PROPS, b"\0\0\0\0")
        pcd = prl(probes.P_TABLE_PROPS, pcd_offset.to_bytes(4, "little"))
        before = fixture(papx, direct + pcd_data, [pcd], prm=1)
        operand_offset = 2 + len(prl(0x2416, b"\x01")) + 2
        after = changed(before, "Data", operand_offset, after_value)
        source = loaded(before)
        plan = {
            "source_sha256": source.source_sha256,
            "edits": [{"stream": "Data", "offset": operand_offset,
                       "before": before_value.hex(), "after": after_value.hex()}],
            "targets": [{"fc": 105, "code": 0x740A,
                         "before": before_value.hex(), "after": after_value.hex()}],
        }
        self.assertTrue(probes.validate_plan(source, after, plan)["valid"])

    def test_exact_diff_source_hash_and_target_schema_fail_closed(self):
        rsid = prl(0x6467, b"\x11\x22\x33\x44")
        ttlp = prl(0x740A, b"\0\0 \0")
        before_data = prc_data(top_level_row(prl(0x563A, (11).to_bytes(2, "little")) + rsid))
        before = fixture(b"\0\0" + prl(probes.P_TABLE_PROPS, b"\0\0\0\0"), before_data)
        offset = 2 + len(prl(0x563A, b"\0\0"))
        after = changed(before, "Data", offset, ttlp)
        source = loaded(before)
        plan = {
            "source_sha256": source.source_sha256,
            "edits": [{"stream": "Data", "offset": offset,
                       "before": rsid.hex(), "after": ttlp.hex()}],
            "targets": [{"fc": 105, "code": "TTlp", "before": None,
                         "after": "00002000"}],
        }
        with self.assertRaisesRegex(probes.ProbeError, "serialized source"):
            probes.validate_plan(source, after, {**plan, "source_sha256": "0" * 64})
        with self.assertRaisesRegex(probes.ProbeError, "source bytes"):
            probes.validate_plan(before, after, plan)
        wrong = {**plan, "edits": [{**plan["edits"][0], "before": "00" * 6}]}
        with self.assertRaisesRegex(probes.ProbeError, "before bytes"):
            probes.validate_plan(source, after, wrong)
        extra = changed(after, "WordDocument", 20, b"\x01")
        with self.assertRaisesRegex(probes.ProbeError, "undeclared"):
            probes.validate_plan(source, extra, plan)
        unsupported = {**plan, "targets": [{"fc": 105, "code": "PJc",
                                             "before": None, "after": "00"}]}
        with self.assertRaisesRegex(probes.ProbeError, "TIstd or TTlp"):
            probes.validate_plan(source, after, unsupported)
        missing_before = {
            **plan,
            "targets": [{"fc": 105, "code": "TTlp", "after": "00002000"}],
        }
        with self.assertRaisesRegex(probes.ProbeError, "integer physical FC"):
            probes.validate_plan(source, after, missing_before)
        with patch.object(probes, "MAX_TARGET_LOOKUP_WORK", 1):
            with self.assertRaisesRegex(probes.ProbeError, "work budget"):
                probes.validate_plan(source, after, plan)
        with patch.object(probes, "MAX_INSPECT_PAYLOAD_BYTES", 1):
            with self.assertRaisesRegex(probes.ProbeError, "payload budget"):
                probes.validate_plan(source, after, plan)

    def test_validator_rejects_non_ttp_and_historical_ttlp_only_targets(self):
        rsid = prl(0x6467, b"\x11\x22\x33\x44")
        ttlp = prl(0x740A, b"\0\0 \0")
        data = prc_data(top_level_row(prl(0x563A, (11).to_bytes(2, "little")) + rsid))
        ordinary = fixture(b"\0\0" + prl(probes.P_TABLE_PROPS, b"\0\0\0\0"), data)
        ordinary = changed(ordinary, "WordDocument", 105, b"\x0d")
        offset = 2 + len(prl(0x563A, b"\0\0"))
        ordinary_after = changed(ordinary, "Data", offset, ttlp)
        ordinary_source = loaded(ordinary)
        ordinary_plan = {
            "source_sha256": ordinary_source.source_sha256,
            "edits": [{"stream": "Data", "offset": offset,
                       "before": rsid.hex(), "after": ttlp.hex()}],
            "targets": [{"fc": 105, "code": "TTlp", "before": None,
                         "after": "00002000"}],
        }
        with self.assertRaisesRegex(probes.ProbeError, "top-level TTP character"):
            probes.validate_plan(ordinary_source, ordinary_after, ordinary_plan)

        before_value = b"\x01\0 \0"
        after_value = b"\x02\0 \0"
        before_data = prc_data(top_level_row(prl(0x740A, before_value)))
        before = fixture(b"\0\0" + prl(probes.P_TABLE_PROPS, b"\0\0\0\0"), before_data)
        value_offset = 2 + 2
        after = changed(before, "Data", value_offset, after_value[:2])
        source = loaded(before)
        plan = {
            "source_sha256": source.source_sha256,
            "edits": [{"stream": "Data", "offset": value_offset,
                       "before": before_value[:2].hex(), "after": after_value[:2].hex()}],
            "targets": [{"fc": 105, "code": "TTlp",
                         "before": before_value.hex(), "after": after_value.hex()}],
        }
        with self.assertRaisesRegex(probes.ProbeError, "historical itl"):
            probes.validate_plan(source, after, plan)

    def test_trace_rejects_cycles_ranges_and_invalid_prc_lengths(self):
        reference = prl(probes.P_TABLE_PROPS, b"\0\0\0\0")
        cycle = prc_data(reference + prl(0x563A, b"\0\0"))
        with self.assertRaisesRegex(probes.ProbeError, "cyclic"):
            probes.trace_properties({"WordDocument": reference, "Data": cycle},
                                    "WordDocument", 0, len(reference))
        with self.assertRaisesRegex(probes.ProbeError, "property range"):
            probes.trace_properties({"WordDocument": reference}, "WordDocument", 2, 1)
        for size, message in [(9, "short"), (0x3FA3, "oversized")]:
            data = size.to_bytes(2, "little") + bytes(size)
            with self.assertRaisesRegex(probes.ProbeError, message):
                probes.trace_properties({"WordDocument": reference, "Data": data},
                                        "WordDocument", 0, len(reference))
        truncated = (10).to_bytes(2, "little") + bytes(9)
        with self.assertRaisesRegex(probes.ProbeError, "outside Data"):
            probes.trace_properties({"WordDocument": reference, "Data": truncated},
                                    "WordDocument", 0, len(reference))

    def test_inspector_rejects_fkp_clx_and_diagnostic_budget_boundaries(self):
        streams = fixture(b"\0\0")
        shared = bytearray(streams["WordDocument"])
        page = memoryview(shared)[512:1024]
        page[0:4] = (100).to_bytes(4, "little")
        page[4:8] = (105).to_bytes(4, "little")
        page[8:12] = (110).to_bytes(4, "little")
        page[12] = 32
        page[25] = 32
        page[511] = 2
        shared_inspection = probes.inspect_document({
            **streams, "WordDocument": bytes(shared),
        })
        self.assertEqual(len(shared_inspection["papx_runs"]), 2)
        self.assertEqual(len(shared_inspection["property_traces"]), 1)
        self.assertEqual(
            shared_inspection["papx_runs"][0]["direct_trace_id"],
            shared_inspection["papx_runs"][1]["direct_trace_id"],
        )

        word = bytearray(streams["WordDocument"])
        page = memoryview(word)[512:1024]
        page[8] = 250
        page[500] = 6
        page[511] = 1
        malformed = {**streams, "WordDocument": bytes(word)}
        with self.assertRaisesRegex(probes.ProbeError, "FKP payload area"):
            probes.inspect_document(malformed)

        oversized_clx = fixture(b"\0\0", prcs=[bytes(0x3FA3)], prm=1)
        with self.assertRaisesRegex(probes.ProbeError, "CLX property"):
            probes.inspect_document(oversized_clx)

        table = bytearray(streams["0Table"])
        table[27:31] = ((5000 * 2) | 0x40000000).to_bytes(4, "little")
        outside = {**streams, "0Table": bytes(table)}
        with self.assertRaisesRegex(probes.ProbeError, "physical range"):
            probes.inspect_document(outside)

        with patch.object(probes, "MAX_DIAGNOSTIC_RECORDS", 1):
            with self.assertRaisesRegex(probes.ProbeError, "record budget"):
                probes.inspect_document(streams)
        with patch.object(probes, "MAX_INSPECT_PAYLOAD_BYTES", 1):
            with self.assertRaisesRegex(probes.ProbeError, "payload budget"):
                probes.inspect_document(streams)

    def test_loader_imports_olefile_lazily_and_hashes_exact_source_bytes(self):
        source = b"synthetic cfb"
        streams = {"WordDocument": b"word", "0Table": b"table"}

        class FakeOle:
            def __init__(self, value):
                self.source = value.read()
                self.closed = False

            def listdir(self, streams=True, storages=False):
                self.assert_options = (streams, storages)
                return [[name] for name in self_streams]

            def get_size(self, parts):
                return len(self_streams[parts[0]])

            def openstream(self, parts):
                return BytesIO(self_streams[parts[0]])

            def close(self):
                self.closed = True

        self_streams = streams
        fake_module = SimpleNamespace(OleFileIO=FakeOle)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory, "source.doc")
            path.write_bytes(source)
            with patch.object(
                probes.importlib, "import_module", return_value=fake_module
            ) as imported:
                result = probes.load_document(path)
                with self.assertRaisesRegex(probes.ProbeError, "aggregate streams"):
                    probes.load_document(path, max_aggregate_bytes=6)
            self.assertEqual(imported.call_count, 2)
            self.assertTrue(all(call.args == ("olefile",) for call in imported.call_args_list))
            with patch.object(probes, "MAX_PLAN_BYTES", 1):
                with self.assertRaisesRegex(probes.ProbeError, "plan exceeds"):
                    probes._load_plan(path)
        self.assertEqual(dict(result.streams), streams)
        self.assertEqual(result.source_sha256, sha256(source).hexdigest())


if __name__ == "__main__":
    unittest.main()
