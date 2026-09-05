import importlib.util
from io import BytesIO
from pathlib import Path
import unittest
from zipfile import ZipFile

from lxml import etree as E

spec = importlib.util.spec_from_file_location("probes", Path(__file__).with_name("legacy-doc-list-probes.py"))
probes = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probes)
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"


class ProbeTests(unittest.TestCase):
    def test_matrix_changes_one_parameter_and_has_unchanged_repeats(self):
        cases = probes.matrix()
        self.assertEqual(len(cases), 64)
        by_id = {c["id"]: c for c in cases}
        self.assertEqual(len(by_id), len(cases))
        for case in cases:
            if case["parent"] is None:
                continue
            original = by_id[case["parent"]]["parameters"]
            changed = [k for k, v in case["parameters"].items() if original[k] != v]
            self.assertEqual(changed, case["changed"])
            self.assertLessEqual(len(changed), 1)
        self.assertEqual(sum(c["parent"] is not None and not c["changed"] for c in cases), 2)

    def test_sources_are_deterministic_passive_and_have_unique_numbering(self):
        payload = probes.build(probes.matrix())
        self.assertEqual(payload, probes.build(probes.matrix()))
        with ZipFile(BytesIO(payload)) as z:
            self.assertFalse(any("vba" in name.lower() for name in z.namelist()))
            for name in z.namelist():
                if name.endswith(".rels"):
                    self.assertFalse(any(r.get("TargetMode") == "External" for r in E.fromstring(z.read(name))))
            document = E.fromstring(z.read("word/document.xml"))
            numbering = E.fromstring(z.read("word/numbering.xml"))
            self.assertEqual(len(document.findall(".//w:p", NS)), 256)
            self.assertFalse(document.findall(".//w:fldChar", NS))
            definitions = numbering.findall("w:abstractNum", NS)
            instances = numbering.findall("w:num", NS)
            self.assertEqual(len(definitions), 64)
            self.assertEqual(len(instances), 64)
            self.assertEqual(len({n.get(W + "numId") for n in instances}), 64)
            self.assertEqual([E.QName(x).localname for x in numbering], ["abstractNum"] * 64 + ["num"] * 64)

    def test_xml_preserves_absent_zero_style_and_tab_conditions(self):
        cases = probes.matrix()
        with ZipFile(BytesIO(probes.build(cases))) as z:
            doc = E.fromstring(z.read("word/document.xml"))
            styles = E.fromstring(z.read("word/styles.xml"))
            numbering = E.fromstring(z.read("word/numbering.xml"))
            paragraphs = doc.findall("w:body/w:p", NS)
            for i, case in enumerate(cases):
                expected = case["parameters"]
                props = paragraphs[4 * i + 2].find("w:pPr", NS)
                self.assertEqual(props.find("w:bidi", NS).get(W + "val"), str(int(expected["rtl"])))
                self.assertEqual(props.find("w:numPr", NS) is None, expected["numbering_in_style"])
                indent = props.find("w:ind", NS)
                exists = expected["empty_direct_ind"] or any(expected[k] is not None for k in ["direct_left", "direct_right", "direct_first"])
                self.assertEqual(indent is not None, exists)
                if expected["direct_left"] is not None:
                    self.assertEqual(indent.get(W + "left"), str(expected["direct_left"]))
                if expected["direct_first"] is not None:
                    value = expected["direct_first"]
                    self.assertEqual(indent.get(W + ("hanging" if value < 0 else "firstLine")), str(abs(value)))
                style_id = props.find("w:pStyle", NS).get(W + "val")
                style = styles.xpath("w:style[@w:styleId=$id]", namespaces=NS, id=style_id)[0]
                self.assertEqual(style.find("w:pPr/w:numPr", NS) is not None, expected["numbering_in_style"])
                level = numbering.findall("w:abstractNum", NS)[i].find("w:lvl", NS)
                self.assertEqual(level.find("w:suff", NS).get(W + "val"), expected["suffix"])
                tab = props.find("w:tabs/w:tab", NS)
                if expected["direct_tab"] == "clear":
                    self.assertEqual((tab.get(W + "val"), tab.get(W + "pos")), ("clear", "720"))


if __name__ == "__main__":
    unittest.main()
