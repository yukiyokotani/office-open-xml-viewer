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


def top_level_row(grpprl):
    return (
        grpprl
        + prl(0x6649, (1).to_bytes(4, "little"))
        + prl(0x2417, b"\x01")
    )


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

    def test_later_complex_pcd_shadow_prevents_a_false_data_target(self):
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
        with self.assertRaisesRegex(probes.ProbeError, "effective target"):
            probes.validate_plan(source, after, plan)

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
            imported.assert_called_once_with("olefile")
            with patch.object(probes, "MAX_PLAN_BYTES", 1):
                with self.assertRaisesRegex(probes.ProbeError, "plan exceeds"):
                    probes._load_plan(path)
        self.assertEqual(dict(result.streams), streams)
        self.assertEqual(result.source_sha256, sha256(source).hexdigest())


if __name__ == "__main__":
    unittest.main()
