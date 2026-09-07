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
    def test_alignment_precedence_matrix_is_bounded_and_one_dimensional(self):
        cases = probes.matrix("alignment-precedence")
        self.assertEqual(len(cases), 18)
        by_id = {case["id"]: case for case in cases}
        self.assertEqual(len(by_id), 18)
        self.assertEqual(set(by_id), {f"T{i:03}" for i in range(1, 19)})
        for case in cases:
            if case["parent"]:
                parent = by_id[case["parent"]]["parameters"]
                actual = [key for key, value in case["parameters"].items() if parent[key] != value]
                self.assertEqual(actual, case["changed"])
                self.assertLessEqual(len(actual), 1)
        for rtl in [False, True]:
            subset = [case for case in cases if case["parameters"]["rtl"] == rtl]
            self.assertEqual(len(subset), 9)
            self.assertEqual(
                {case["parameters"]["direct_alignment"] for case in subset},
                {None, "left", "right", "center", "both", "distribute"},
            )
            self.assertEqual(sum(case["parameters"]["direct_alignment"] is not None for case in subset), 5)
            self.assertEqual(sum(case["parameters"]["direct_alignment"] is None for case in subset), 4)
            self.assertEqual(sum(case["parent"] is not None and not case["changed"] for case in subset), 2)
            baseline = next(case for case in subset if case["parent"] is None)
            list_aligned = next(case for case in subset if case["changed"] == ["list_alignment"])
            self.assertIsNone(baseline["parameters"]["list_alignment"])
            self.assertEqual(list_aligned["parameters"]["list_alignment"], "left" if rtl else "right")

    def test_alignment_precedence_xml_owns_exact_jc_and_schema_order(self):
        cases = probes.matrix("alignment-precedence")
        payload = probes.build(cases)
        self.assertEqual(payload, probes.build(cases))
        with ZipFile(BytesIO(payload)) as z:
            document = E.fromstring(z.read("word/document.xml"))
            numbering = E.fromstring(z.read("word/numbering.xml"))
            styles = E.fromstring(z.read("word/styles.xml"))
        paragraphs = document.findall("w:body/w:p", NS)
        self.assertEqual(len(paragraphs), 72)
        levels = numbering.findall("w:abstractNum/w:lvl", NS)
        self.assertEqual(len(levels), 18)
        self.assertEqual(len(numbering.findall("w:num", NS)), 18)
        self.assertEqual(
            [E.QName(child).localname for child in numbering],
            ["abstractNum"] * 18 + ["num"] * 18,
        )

        order = {name: index for index, name in enumerate(["tabs", "bidi", "ind", "jc"])}
        for index, case in enumerate(cases):
            params = case["parameters"]
            level_ppr = levels[index].find("w:pPr", NS)
            level_names = [E.QName(child).localname for child in level_ppr]
            selected = [order[name] for name in level_names if name in order]
            self.assertEqual(selected, sorted(selected), (case["id"], level_names))
            level_jc = level_ppr.find("w:jc", NS)
            if params["list_alignment"] is None:
                self.assertIsNone(level_jc)
            else:
                self.assertEqual(level_jc.get(W + "val"), params["list_alignment"])
            self.assertEqual(level_ppr.find("w:bidi", NS).get(W + "val"), str(int(not params["rtl"])))

            for paragraph in paragraphs[index * 4 + 2:index * 4 + 4]:
                ppr = paragraph.find("w:pPr", NS)
                names = [E.QName(child).localname for child in ppr]
                selected = [order[name] for name in names if name in order]
                self.assertEqual(selected, sorted(selected), (case["id"], names))
                direct_jc = ppr.find("w:jc", NS)
                if params["direct_alignment"] is None:
                    self.assertIsNone(direct_jc)
                else:
                    self.assertEqual(direct_jc.get(W + "val"), params["direct_alignment"])
                self.assertEqual(ppr.find("w:bidi", NS).get(W + "val"), str(int(params["rtl"])))

                style_id = ppr.find("w:pStyle", NS).get(W + "val")
                visited = set()
                while style_id:
                    self.assertNotIn(style_id, visited)
                    visited.add(style_id)
                    style = styles.xpath("w:style[@w:styleId=$id]", namespaces=NS, id=style_id)[0]
                    style_ppr = style.find("w:pPr", NS)
                    if style_ppr is not None:
                        self.assertIsNone(style_ppr.find("w:bidi", NS))
                        self.assertIsNone(style_ppr.find("w:jc", NS))
                    based_on = style.find("w:basedOn", NS)
                    style_id = based_on.get(W + "val") if based_on is not None else None
                self.assertIn("Normal", visited)

    def test_bidi_boundaries_change_one_parameter_with_plain_and_repeat_controls(self):
        cases = probes.matrix("bidi-boundaries")
        self.assertEqual(len(cases), 32)
        by_id = {case["id"]: case for case in cases}
        self.assertEqual(len(by_id), 32)
        for case in cases:
            if case["parent"]:
                previous = by_id[case["parent"]]["parameters"]
                actual = [key for key, value in case["parameters"].items() if previous[key] != value]
                self.assertEqual(actual, case["changed"])
                self.assertLessEqual(len(actual), 1)
        for rtl in [False, True]:
            subset = [c for c in cases if c["parameters"]["rtl"] == rtl]
            self.assertEqual(len(subset), 16)
            self.assertEqual({c["parameters"]["terminal_punctuation"] for c in subset}, {"", ".", "!", ":", "?"})
            self.assertEqual({c["parameters"]["text_runs"] for c in subset}, {"whole", "punctuation", "words"})
            self.assertEqual({c["parameters"]["run_rtl"] for c in subset}, {None, False})
            self.assertEqual({c["parameters"]["numbered"] for c in subset}, {False, True})
            self.assertEqual(sum(c["parent"] is not None and not c["changed"] for c in subset), 2)

    def test_bidi_source_preserves_run_boundaries_without_changing_logical_text(self):
        cases = probes.matrix("bidi-boundaries")
        with ZipFile(BytesIO(probes.build(cases))) as z:
            doc = E.fromstring(z.read("word/document.xml"))
            paragraphs = doc.findall("w:body/w:p", NS)
            self.assertEqual(len(paragraphs), 128)
            for index, case in enumerate(cases):
                params = case["parameters"]
                for label, paragraph in zip(["First", "Second"], paragraphs[index*4+2:index*4+4]):
                    self.assertEqual(paragraph.find("w:pPr/w:bidi", NS).get(W+"val"), str(int(params["rtl"])))
                    self.assertEqual(paragraph.find("w:pPr/w:numPr", NS) is not None, params["numbered"])
                    text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
                    self.assertEqual(text, label + " marker line alpha bravo charlie delta" + params["terminal_punctuation"]
                        + "Continuation line alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november oscar papa"
                        + params["terminal_punctuation"])
                    self.assertEqual(len(paragraph.findall(".//w:br", NS)), 1)
                    runs = paragraph.findall("w:r", NS)
                    text_runs = [run for run in runs if run.find("w:t", NS) is not None]
                    for run in text_runs:
                        mark = run.find("w:rPr/w:rtl", NS)
                        if params["run_rtl"] is None:
                            self.assertIsNone(mark)
                        else:
                            self.assertEqual(mark.get(W+"val"), "0")
                    pieces = ["".join(run.xpath("w:t/text()", namespaces=NS)) for run in text_runs]
                    if params["text_runs"] == "punctuation" and params["terminal_punctuation"]:
                        self.assertEqual(pieces.count(params["terminal_punctuation"]), 2)
                    elif params["text_runs"] == "words":
                        self.assertGreater(len(pieces), 20)
                    else:
                        self.assertEqual(len(text_runs), 2)

    def test_bidi_controls_reject_undefined_strong_latin_rtl_true(self):
        cases = probes.matrix("bidi-boundaries")
        cases[0]["parameters"]["run_rtl"] = True
        with self.assertRaisesRegex(ValueError, "not a defined-behavior control"):
            probes.build(cases)

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
        for phase in ["baseline", "interactions", "style-association", "bidi-boundaries", "alignment-precedence"]:
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
