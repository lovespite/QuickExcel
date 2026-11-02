"""Independent MuPDF semantic/visual checks of the headless rendering smoke (test-only dependency)."""
import argparse
import json
import math
import sys
from pathlib import Path

parser = argparse.ArgumentParser()
parser.add_argument("folder", type=Path)
parser.add_argument("--dependencies", type=Path)
parser.add_argument("--compare", type=Path)
args = parser.parse_args()
if args.dependencies:
    sys.path.insert(0, str(args.dependencies.resolve()))
import pymupdf
import numpy as np
from PIL import Image, ImageFilter

folder = args.folder.resolve()
text_pdf = pymupdf.open(folder / "text.pdf")
semantic = text_pdf[0].get_text()
for required in ["QuickExcel 中文报表：收入、成本和利润", "e\u0301 A\u030a", "office ffi", "العربية", "עברית", "𝄞", "𠮷"]:
    assert required in semantic, f"Missing Unicode/cluster text: {required!r}"
(folder / "semantic-text.txt").write_text(semantic, encoding="utf-8")
layout = json.loads((folder / "report-layout.json").read_text(encoding="utf-8-sig"))
diagnostics = json.loads((folder / "report-diagnostics.json").read_text(encoding="utf-8-sig"))
assert not diagnostics["Warnings"]
assert not [d for d in diagnostics["Diagnostics"] if d["Code"] == "missing-glyph"]
document = pymupdf.open(folder / "report.pdf")
assert len(document) == len(layout["Pages"]) == 3
results = []
for index, page in enumerate(document):
    assert abs(page.rect.width - layout["PaperWidth"] * .75) < .2
    assert abs(page.rect.height - layout["PaperHeight"] * .75) < .2
    text = page.get_text()
    assert "QuickExcel 中文报表" in text and f"第 {index + 1} / 3 页" in text
    assert len(page.get_drawings()) > 30, "Expected vector grid/chart content"
    fonts = page.get_fonts()
    assert any(f[2] == "Type0" and document.extract_font(f[0])[3] for f in fonts), "Missing embedded subset font"
    assert len(page.get_images()) <= 2, "Report must not become a page screenshot"
    # Check every body record's page membership and position against the frozen Core page plan.
    y = layout["Margins"]["Top"] * 96
    for row in layout["Pages"][index]["Rows"]:
        height = {0: 42, 4: 42, 5: 42, 8: 72}.get(row, 28)
        if row >= 2:
            hits = page.search_for(f"Record {row:02}")
            assert hits, f"Missing record {row} on page {index + 1}"
            # The chart can repeat categories; the first record appears in the table.
            expected_y = y * .75
            assert abs(hits[0].y0 - expected_y) < height * layout["Scale"] * .75, (row, hits[0], expected_y)
        y += height * layout["Scale"]
    preview = Image.open(folder / f"report-{index + 1}.png").convert("RGB")
    pixmap = page.get_pixmap(dpi=96, alpha=False)
    rendered = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    rendered.save(folder / f"pdf-render-{index + 1}.png")
    assert preview.size == rendered.size
    # Allow one pixel of antialias/hinting displacement, then inspect each occupied 128px tile.
    a, b = np.array(preview).astype(int), np.array(rendered).astype(int)
    low = np.array(preview.filter(ImageFilter.MinFilter(3))).astype(int)
    high = np.array(preview.filter(ImageFilter.MaxFilter(3))).astype(int)
    # FreeType and the PDF engine differ in low-coverage edge intensity (especially 10pt CJK).
    # A 64/255 intensity tolerance is restricted to a one-pixel neighborhood, per occupied tile.
    mismatch = np.maximum(low - b, b - high).max(axis=2) > 64
    occupied = np.minimum(a.min(axis=2), b.min(axis=2)) < 220
    tiles = []
    for top in range(0, a.shape[0], 128):
        for left in range(0, a.shape[1], 128):
            mask = occupied[top:top+128, left:left+128]
            if mask.sum() < 20:
                continue
            fraction = float((mismatch[top:top+128, left:left+128] & mask).sum() / mask.sum())
            assert fraction < .10, f"PDF/preview local drift: page {index + 1}, tile {(left, top)}, {fraction}"
            tiles.append(fraction)
    results.append({"Page": index + 1, "Fonts": [{"Type": f[2], "Name": f[3]} for f in fonts], "MaxTileMismatch": max(tiles), "Records": layout["Pages"][index]["Rows"]})
if args.compare:
    reference = json.loads((args.compare / "report-layout.json").read_text(encoding="utf-8-sig"))
    assert layout == reference, "Cross-platform pagination differs"
    other = pymupdf.open(args.compare / "report.pdf")
    for index, page in enumerate(document):
        assert page.get_text() == other[index].get_text(), "Cross-platform semantic output differs"
        for label in ["QuickExcel", f"第 {index + 1} / 3 页", "Revenue"]:
            actual, expected = page.search_for(label), other[index].search_for(label)
            assert len(actual) == len(expected)
            for a, b in zip(actual, expected):
                assert max(abs(x-y) for x, y in zip(a, b)) < .5, f"Cross-platform text geometry: {label}"
(folder / "independent-validation.json").write_text(json.dumps({"Status": "PASS", "Engine": f"MuPDF {pymupdf.VersionBind}", "Pages": results, "Compared": str(args.compare)}, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"PASS: Unicode, embedded text fonts, page membership, physical units, vector content and local PDF/preview comparison: {folder}")
