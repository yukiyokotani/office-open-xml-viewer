"""Local Office PDF fidelity gate. Page counts/sizes and raster pixels must match.

No image registration, resizing, blur, threshold relaxation or reference updates.
Only one page pair is retained in memory. Artifacts remain local to the run.
"""
import json
from pathlib import Path
import subprocess
import sys

from PIL import Image, ImageChops
from pypdf import PdfReader


def compare(directory):
    directory = Path(directory)
    readers = [PdfReader(directory / f"{label}.pdf") for label in ("source", "converted")]
    counts = [len(reader.pages) for reader in readers]
    if not all(0 < count <= 500 for count in counts):
        raise ValueError("PDF page count exceeds the local verification budget")
    report = {"sourcePages": counts[0], "convertedPages": counts[1], "equal": counts[0] == counts[1], "pages": []}
    for index in range(max(counts)):
        if index >= min(counts):
            report["pages"].append({"page": index + 1, "equal": False, "status": "missing-source-page" if index >= counts[0] else "missing-converted-page"})
            continue
        sizes = [(float(reader.pages[index].mediabox.width), float(reader.pages[index].mediabox.height)) for reader in readers]
        page = {"page": index + 1, "sourceSize": sizes[0], "convertedSize": sizes[1], "equal": False}
        if any(w <= 0 or h <= 0 or w * h * (96 / 72) ** 2 > 50_000_000 for w, h in sizes):
            raise ValueError("PDF page dimensions exceed the local verification budget")
        images = []
        for label in ("source", "converted"):
            prefix = directory / f"{label}-{index + 1:04d}"
            subprocess.run(["pdftoppm", "-f", str(index + 1), "-l", str(index + 1), "-singlefile", "-r", "96", "-png", str(directory / f"{label}.pdf"), str(prefix)], check=True, capture_output=True, timeout=60)
            with Image.open(prefix.with_suffix(".png")) as image:
                images.append(image.convert("RGB"))
        if sizes[0] == sizes[1] and images[0].size == images[1].size:
            diff = ImageChops.difference(*images)
            red, green, blue = diff.split()
            maximum = ImageChops.lighter(ImageChops.lighter(red, green), blue)
            pixels = images[0].width * images[0].height
            page["changedPixels"] = pixels - maximum.histogram()[0]
            page["totalPixels"] = pixels
            page["equal"] = page["changedPixels"] == 0
            if not page["equal"]:
                diff.save(directory / f"diff-{index + 1:04d}.png")
        report["pages"].append(page)
        report["equal"] = report["equal"] and page["equal"]
        for image in images:
            image.close()
    return report


if __name__ == "__main__":
    print(json.dumps(compare(sys.argv[1])))
