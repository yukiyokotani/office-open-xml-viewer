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
    def test_style_association_keeps_one_list_across_distinct_paragraph_styles(self):
        cases = probes.matrix("style-association")
        self.assertEqual(len(cases), 48)
        by_id = {c["id"]: c for c in cases}
        self.assertEqual(len(by_id), 48)
        for case in cases:
            if case["parent"]:
                original = by_id[case["parent"]]["parameters"]
                changed = [k for k, v in case["parameters"].items() if original[k] != v]
                self.assertEqual(changed, case["changed"])
                self.assertLessEqual(len(changed), 1)
        self.assertEqual(sum(c["parent"] is not None and not c["changed"] for c in cases), 8)
        with ZipFile(BytesIO(probes.build(cases))) as z:
            doc = E.fromstring(z.read("word/document.xml"))
            style_defs = E.fromstring(z.read("word/styles.xml"))
            paragraphs = doc.findall("w:body/w:p", NS)
            self.assertEqual(len(paragraphs), 192)
            for i, case in enumerate(cases):
                pair = paragraphs[4*i+2:4*i+4]
                self.assertEqual([p.find("w:pPr/w:numPr/w:numId", NS).get(W+"val") for p in pair], [str(i+1)]*2)
                styles = [p.find("w:pPr/w:pStyle", NS) for p in pair]
                styles = [s.get(W+"val") if s is not None else None for s in styles]
                main = "Probe" + case["id"]
                expected = {"uniform": [main, main], "normal": [None, None],
                            "mixed-normal": [main, None], "alternating": [main, main+"Alternate"]}
                self.assertEqual(styles, expected[case["parameters"]["paragraph_styles"]])
                if case["parameters"]["paragraph_styles"] == "alternating":
                    definitions = [style_defs.xpath("w:style[@w:styleId=$id]", namespaces=NS, id=s)[0] for s in styles]
                    self.assertEqual([d.find("w:basedOn", NS).get(W+"val") for d in definitions], ["Normal"]*2)
                    self.assertEqual(E.tostring(definitions[0].find("w:pPr", NS)),
                                     E.tostring(definitions[1].find("w:pPr", NS)))

    def test_all_phases_remain_deterministic_and_passive(self):
        for phase in ["baseline", "interactions", "style-association"]:
            with self.subTest(phase=phase):
                cases = probes.matrix(phase)
                payload = probes.build(cases)
                self.assertEqual(payload, probes.build(cases))
                with ZipFile(BytesIO(payload)) as z:
                    self.assertFalse(any(any(s in name.lower() for s in ["vba", "activex", "embeddings"]) for name in z.namelist()))
                    for name in z.namelist():
                        if name.endswith(".rels"):
                            self.assertFalse(any(r.get("TargetMode") == "External" for r in E.fromstring(z.read(name))))
                    doc = E.fromstring(z.read("word/document.xml"))
                    self.assertFalse(doc.findall(".//w:fldChar", NS))
                    self.assertFalse(doc.findall(".//w:fldSimple", NS))
                    numbering = E.fromstring(z.read("word/numbering.xml"))
                    self.assertEqual(len(numbering.findall("w:abstractNum", NS)), len(cases))
                    self.assertEqual(len(numbering.findall("w:num", NS)), len(cases))
                    # ECMA-376 CT_Numbering is a sequence, not a choice:
                    # all abstractNum definitions precede every num instance.
                    # Word repair of malformed sources must not become evidence
                    # for a list-formatting precedence or suffix rule.
                    self.assertEqual([E.QName(child).localname for child in numbering],
                                     ["abstractNum"] * len(cases) + ["num"] * len(cases))

    def test_interactions_cover_conflicting_right_indent_and_twip_boundaries(self):
        cases = probes.matrix("interactions")
        self.assertEqual(len(cases), 64)
        by_id = {c["id"]: c for c in cases}
        self.assertEqual(len(by_id), 64)
        for case in cases:
            self.assertEqual(case["parameters"]["list_right"], 720)
            if case["parent"]:
                original = by_id[case["parent"]]["parameters"]
                changed = [k for k, v in case["parameters"].items() if original[k] != v]
                self.assertEqual(changed, case["changed"])
                self.assertLessEqual(len(changed), 1)
        for rtl in [False, True]:
            subset = [c for c in cases if c["parameters"]["rtl"] == rtl]
            for value in [-1, 0, 1, 719, 720, 721]:
                self.assertTrue(any(c["parameters"]["direct_right"] == value for c in subset))
            self.assertTrue(any(all(c["parameters"][k] == 0 for k in
                ["direct_left", "direct_right", "direct_first"]) for c in subset))
        with ZipFile(BytesIO(probes.build(cases))) as z:
            doc = E.fromstring(z.read("word/document.xml"))
            self.assertEqual(len(doc.findall(".//w:p", NS)), 256)
            paragraphs = doc.findall("w:body/w:p", NS)
            levels = E.fromstring(z.read("word/numbering.xml")).findall("w:abstractNum/w:lvl", NS)
            for i, case in enumerate(cases):
                self.assertEqual(levels[i].find("w:pPr/w:ind", NS).get(W + "right"), "720")
                expected = case["parameters"]["direct_right"]
                if expected is not None:
                    ind = paragraphs[4*i + 2].find("w:pPr/w:ind", NS)
                    self.assertEqual(ind.get(W + "right"), str(expected))

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
