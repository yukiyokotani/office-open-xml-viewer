#!/usr/bin/env python3
"""Regenerate part-level fuzz seeds from tracked, public demo packages only."""
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
CORPUS = ROOT / "fuzz" / "corpus"
SOURCES = {
    "docx": ("fuzz_docx_part", [
        "word/document.xml", "word/styles.xml", "word/numbering.xml",
        "word/header1.xml", "word/footer1.xml", "word/footnotes.xml",
    ]),
    "xlsx": ("fuzz_xlsx_part", [
        "xl/worksheets/sheet1.xml", "xl/sharedStrings.xml",
        "xl/styles.xml", "xl/workbook.xml",
    ]),
    "pptx": ("fuzz_pptx_part", [
        "ppt/slides/slide1.xml", "ppt/slideLayouts/slideLayout1.xml",
        "ppt/slideMasters/slideMaster1.xml", "ppt/theme/theme1.xml",
    ]),
}

for fmt, (target, parts) in SOURCES.items():
    source = ROOT / "packages" / fmt / "public" / "demo" / f"sample-1.{fmt}"
    destination = CORPUS / target
    destination.mkdir(parents=True, exist_ok=True)
    with ZipFile(source) as archive:
        for selector, part in enumerate(parts):
            if part in archive.namelist():
                (destination / f"public-{selector:02d}").write_bytes(
                    bytes([selector]) + archive.read(part)
                )

docx_dir = CORPUS / "fuzz_docx_part"
word_ns = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
for selector, name, root in ((3, "header", b"hdr"), (4, "footer", b"ftr")):
    # No tracked public DOCX demo contains these optional story parts.
    (docx_dir / f"synthetic-{name}").write_bytes(
        bytes([selector]) + b"<w:" + root + b' xmlns:w="' + word_ns
        + b'"><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:' + root + b">"
    )

chart_dir = CORPUS / "fuzz_chart_part"
chart_dir.mkdir(parents=True, exist_ok=True)
with ZipFile(ROOT / "packages" / "pptx" / "public" / "demo" / "sample-1.pptx") as archive:
    (chart_dir / "public-chart").write_bytes(b"\x00" + archive.read("ppt/charts/chart1.xml"))

# The tracked demos have no chartEx part. This small synthetic seed reaches
# the chartEx path; all extracted sample bytes above come from public fixtures.
(chart_dir / "synthetic-chartex").write_bytes(
    b"\x01<cx:chartSpace xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\">"
    b"<cx:chartData><cx:data id=\"0\"><cx:numDim type=\"val\"><cx:lvl ptCount=\"1\">"
    b"<cx:pt idx=\"0\">1</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>"
    b"<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId=\"boxWhisker\"/>"
    b"</cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"
)
