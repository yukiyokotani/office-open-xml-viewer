"""Tests for the strict local PDF comparison gate (requires Poppler)."""
import importlib.util
from pathlib import Path
import tempfile
import unittest
from pypdf import PdfWriter

spec = importlib.util.spec_from_file_location("compare", Path(__file__).with_name("legacy-office-compare.py"))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class ComparisonTests(unittest.TestCase):
    def compare_pages(self, first, second):
        with tempfile.TemporaryDirectory() as directory:
            for label, sizes in [("source", first), ("converted", second)]:
                writer = PdfWriter()
                for width, height in sizes:
                    writer.add_blank_page(width, height)
                with open(Path(directory) / f"{label}.pdf", "wb") as output:
                    writer.write(output)
            return module.compare(directory)

    def test_equal_pixels(self):
        result = self.compare_pages([(72, 72)], [(72, 72)])
        self.assertTrue(result["equal"])
        self.assertEqual(result["pages"][0]["changedPixels"], 0)

    def test_missing_pages_fail_even_when_all_common_pages_match(self):
        result = self.compare_pages([(72, 72), (72, 72)], [(72, 72)])
        self.assertFalse(result["equal"])
        self.assertTrue(result["pages"][0]["equal"])
        self.assertEqual(result["pages"][1]["status"], "missing-converted-page")

    def test_size_difference_is_not_rescaled_away(self):
        self.assertFalse(self.compare_pages([(72, 72)], [(73, 72)])["equal"])


if __name__ == "__main__":
    unittest.main()
