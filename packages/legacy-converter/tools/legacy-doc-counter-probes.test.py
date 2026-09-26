import importlib.util
from copy import deepcopy
from io import BytesIO
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree as E

MODULE = Path(__file__).with_name("legacy-doc-counter-probes.py")
spec = importlib.util.spec_from_file_location("counter_probes", MODULE)
probes = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probes)
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"


def xml_shape(node):
    return (node.tag, tuple(sorted(node.attrib.items())), node.text,
            tuple(xml_shape(child) for child in node))


class CounterProbeTests(unittest.TestCase):
    def package(self):
        return ZipFile(BytesIO(probes.build(probes.cases())))

    def test_matrix_has_nine_isolated_four_paragraph_cases(self):
        cases = probes.cases()
        self.assertEqual([case["id"] for case in cases], [f"C{i:02}" for i in range(1, 10)])
        self.assertEqual([case["numPattern"] for case in cases], [
            "1111", "1212", "1122", "1212", "1212", "1212", "1212", "1212", "1212",
        ])
        self.assertTrue(all(case["counterOutcome"] == "UNOBSERVED" for case in cases))
        self.assertEqual(cases[1]["settings"], {"alias": "bare"})
        self.assertEqual(cases[8]["settings"], {"alias": "bare", "repeatOf": "C02"})

    def test_numbering_uses_distinct_case_ids_schema_order_and_exact_variants(self):
        with self.package() as archive:
            numbering = E.fromstring(archive.read("word/numbering.xml"))
            children = list(numbering)
            abstracts = numbering.findall("w:abstractNum", NS)
            nums = numbering.findall("w:num", NS)
            self.assertEqual([E.QName(node).localname for node in children],
                             ["abstractNum"] * 10 + ["num"] * 18)
            self.assertEqual(len({node.get(W + "abstractNumId") for node in abstracts}), 10)
            self.assertEqual(len({node.get(W + "numId") for node in nums}), 18)
            by_num = {node.get(W + "numId"): node for node in nums}
            for index, case in enumerate(probes.cases(), 1):
                first, second = str(index * 2 - 1), str(index * 2)
                first_abs = by_num[first].find("w:abstractNumId", NS).get(W + "val")
                second_abs = by_num[second].find("w:abstractNumId", NS).get(W + "val")
                self.assertEqual(first_abs, str(index * 10))
                self.assertEqual(second_abs == first_abs, index != 8)
            self.assertFalse(by_num["4"].findall("w:lvlOverride", NS))  # C02 bare alias
            by_abstract = {node.get(W + "abstractNumId"): node for node in abstracts}
            base_c04 = by_abstract["40"].find("w:lvl", NS)
            replacement_c04 = by_num["8"].find("w:lvlOverride/w:lvl", NS)
            self.assertEqual(xml_shape(base_c04), xml_shape(replacement_c04))
            bold = by_num["10"].find("w:lvlOverride/w:lvl/w:rPr/w:b", NS)
            self.assertIsNotNone(bold)
            base_c05 = deepcopy(by_abstract["50"].find("w:lvl", NS))
            replacement_c05 = deepcopy(by_num["10"].find("w:lvlOverride/w:lvl", NS))
            replacement_c05.remove(replacement_c05.find("w:rPr", NS))
            self.assertEqual(xml_shape(base_c05), xml_shape(replacement_c05))
            self.assertEqual(by_num["12"].find("w:lvlOverride/w:startOverride", NS).get(W + "val"), "7")
            self.assertIsNone(by_num["12"].find("w:lvlOverride/w:lvl", NS))
            self.assertEqual(by_num["14"].find("w:lvlOverride/w:lvl/w:start", NS).get(W + "val"), "7")
            replacement_c07 = deepcopy(by_num["14"].find("w:lvlOverride/w:lvl", NS))
            replacement_c07.find("w:start", NS).set(W + "val", "1")
            self.assertEqual(xml_shape(by_abstract["70"].find("w:lvl", NS)), xml_shape(replacement_c07))
            self.assertEqual(xml_shape(by_abstract["80"].find("w:lvl", NS)),
                             xml_shape(by_abstract["81"].find("w:lvl", NS)))
            self.assertFalse(by_num["18"].findall("w:lvlOverride", NS))  # repeated bare alias

    def test_document_references_only_instances_and_uses_dynamic_markers(self):
        with self.package() as archive:
            document = E.fromstring(archive.read("word/document.xml"))
            paragraphs = document.findall("w:body/w:p", NS)
            titles = [p for p in paragraphs if p.find("w:pPr/w:pStyle", NS) is not None]
            numbered = [p for p in paragraphs if p.find("w:pPr/w:numPr", NS) is not None]
            self.assertEqual(len(numbered), 36)
            self.assertEqual(len(document.findall(".//w:br", NS)), 0)
            for index, case in enumerate(probes.cases(), 1):
                refs = [p.find("w:pPr/w:numPr/w:numId", NS).get(W + "val")
                        for p in numbered[(index - 1) * 4:index * 4]]
                expected = [str(index * 2 - 2 + int(value)) for value in case["numPattern"]]
                self.assertEqual(refs, expected)
            numbering = E.fromstring(archive.read("word/numbering.xml"))
            self.assertEqual({node.get(W + "val") for node in numbering.findall(".//w:lvlText", NS)}, {"%1."})
            self.assertEqual([
                "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
                for paragraph in numbered
            ], ["Item A", "Item B", "Item C", "Item D"] * 9)
            for paragraph in numbered:
                text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
                self.assertNotRegex(text, r"^\d+[.)]")
            self.assertEqual(len(titles), 9)

    def test_title_is_black_without_border_and_package_is_passive_deterministic(self):
        cases = probes.cases()
        before = deepcopy(cases)
        payload = probes.build(cases)
        self.assertEqual(cases, before)
        self.assertEqual(payload, probes.build(cases))
        with ZipFile(BytesIO(payload)) as archive:
            self.assertFalse(any(any(token in name.lower() for token in ("vba", "activex", "embeddings"))
                                 for name in archive.namelist()))
            for name in archive.namelist():
                if name.endswith(".rels"):
                    self.assertFalse(any(rel.get("TargetMode") == "External"
                                         for rel in E.fromstring(archive.read(name))))
            styles = E.fromstring(archive.read("word/styles.xml"))
            title = styles.xpath("w:style[@w:styleId='Title']", namespaces=NS)[0]
            self.assertEqual(title.find("w:rPr/w:color", NS).get(W + "val"), "000000")
            self.assertIsNone(title.find("w:pPr/w:pBdr", NS))
            document = E.fromstring(archive.read("word/document.xml"))
            self.assertFalse(document.findall(".//w:fldChar", NS))
            self.assertFalse(document.findall(".//w:fldSimple", NS))

    def test_manifest_is_explicit_and_cli_rejects_existing_output(self):
        manifest = probes.manifest(probes.build(probes.cases()), probes.cases())
        self.assertEqual(manifest["oracle"], "source-docx-pdf")
        self.assertTrue(all(case["counterOutcome"] == "UNOBSERVED" for case in manifest["cases"]))
        json.dumps(manifest)
        with tempfile.TemporaryDirectory() as directory:
            result = subprocess.run([sys.executable, str(MODULE), directory], capture_output=True, text=True)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("exists", result.stderr)

    def test_builder_rejects_unknown_or_malformed_case_settings(self):
        for mutate, message in [
            (lambda case: case["settings"].update(alias="typo"), "unsupported"),
            (lambda case: case["settings"].update(unexpected=True), "unsupported"),
            (lambda case: case.update(numPattern="1234"), "pattern"),
            (lambda case: case.update(counterOutcome="shared"), "case"),
        ]:
            matrix = probes.cases()
            mutate(matrix[1])
            with self.assertRaisesRegex(ValueError, message):
                probes.build(matrix)
        matrix = probes.cases()
        matrix[1]["id"] = matrix[0]["id"]
        with self.assertRaisesRegex(ValueError, "duplicate"):
            probes.build(matrix)


if __name__ == "__main__":
    unittest.main()
