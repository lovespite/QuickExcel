"""Validate actual desktop PDFs with MuPDF, including local PDF/preview raster comparison.
Requires Pillow, numpy, pypdf and PyMuPDF (test-only; optional third argument is its module directory).
"""
import hashlib
import json
import math
import re
import sys
from pathlib import Path

if len(sys.argv) > 3:
    sys.path.insert(0, str(Path(sys.argv[3]).resolve()))
import pymupdf
import numpy as np

from PIL import Image, ImageChops, ImageFilter
from pypdf import PdfReader


folder = Path(sys.argv[1]).resolve()
results = []
reports = [("quickexcel", ["123.5", "Merged report title", "中文报表", "E4 revenue"])]
if (folder / "native.pdf").exists():
    reports.append(("native", ["Report 51", "Report 4", "Report 5", "Revenue", "Cost"]))
if (folder / "paged.pdf").exists():
    reports.append(("paged", ["Merged title", "Record 19", "Record 20", "Record 35"]))
for prefix, required in reports:
    layout = json.loads((folder / f"{prefix}-layout.json").read_text(encoding="utf-8-sig"))
    pdf_path = folder / f"{prefix}.pdf"
    pdf = PdfReader(pdf_path)
    actual = pymupdf.open(pdf_path)
    assert len(pdf.pages) == len(layout["Pages"]), "PDF/preview page count differs"
    text = " ".join(page.get_text() for page in actual)
    excel_pagination = None
    for label in required:
        assert label.replace(" ", "") in text.replace(" ", ""), f"Missing PDF text: {label}"
    if prefix == "paged" and len(sys.argv) > 2 and sys.argv[2] != "-":
        reference = PdfReader(Path(sys.argv[2]) / "excel-output-paged.pdf")
        records = lambda document: [re.findall(r"Record\s*(\d+)", page.extract_text()) for page in document.pages]
        assert records(pdf) == records(reference), "Record membership on pages differs from Excel PDF"
        excel_pagination = {"Matches": True, "RecordsByPage": records(pdf), "Reference": str(Path(sys.argv[2]).resolve() / "excel-output-paged.pdf")}
    for page in pdf.pages:
        assert abs(float(page.mediabox.width) - layout["PaperWidth"] * .75) < .2
        assert abs(float(page.mediabox.height) - layout["PaperHeight"] * .75) < .2
    baseline = Image.open(folder / f"{prefix}-96.png").convert("RGB")
    differences = []
    comparisons = []
    for index, page in enumerate(actual):
        assert len(page.get_drawings()) > 0, "PDF must preserve vector objects"
        assert any(actual.extract_font(f[0])[3] for f in page.get_fonts()), "No embedded fonts"
        preview_path = folder / (f"{prefix}-96.png" if index == 0 else f"{prefix}-page{index + 1}.png")
        preview = Image.open(preview_path).convert("RGB")
        pixmap = page.get_pixmap(dpi=96, alpha=False)
        rendered = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
        rendered.save(folder / f"{prefix}-pdf-render-{index + 1}.png")
        # MediaBox rounds within 0.2 pt, so raster extents may differ by one trailing pixel.
        assert abs(preview.width-rendered.width) <= 1 and abs(preview.height-rendered.height) <= 1
        box = (0, 0, min(preview.width, rendered.width), min(preview.height, rendered.height))
        preview, rendered = preview.crop(box), rendered.crop(box)
        a, b = np.asarray(preview).astype(int), np.asarray(rendered).astype(int)
        low = np.asarray(preview.filter(ImageFilter.MinFilter(3))).astype(int)
        high = np.asarray(preview.filter(ImageFilter.MaxFilter(3))).astype(int)
        mismatch = np.maximum(low-b, b-high).max(axis=2) > 64
        ink = np.minimum(a.min(axis=2), b.min(axis=2)) < 220
        tiles = []
        for top in range(0, a.shape[0], 128):
            for left in range(0, a.shape[1], 128):
                mask = ink[top:top+128, left:left+128]
                if mask.sum() < 20: continue
                fraction = float((mismatch[top:top+128, left:left+128] & mask).sum() / mask.sum())
                assert fraction < .10, f"Independent PDF/preview drift: {prefix} page {index+1}, tile {(left,top)}: {fraction}"
                tiles.append(fraction)
        comparisons.append({"Page": index+1, "MaxOccupiedTileMismatch": max(tiles, default=0), "FontTypes": sorted(set(f[2] for f in page.get_fonts()))})
    for dpi in (96, 144, 192):
        picture = Image.open(folder / f"{prefix}-{dpi}.png").convert("RGB")
        assert picture.size == (math.ceil(layout["PaperWidth"] * dpi / 96), math.ceil(layout["PaperHeight"] * dpi / 96))
        normalized = picture.resize(baseline.size, Image.Resampling.LANCZOS)
        difference = ImageChops.difference(baseline, normalized)
        data = difference.get_flattened_data() if hasattr(difference, "get_flattened_data") else difference.getdata()
        count = sum(max(pixel) > 32 for pixel in data)
        fraction = count / (baseline.width * baseline.height)
        assert fraction < .05, f"DPI layout drift: {prefix}/{dpi}: {fraction}"
        differences.append({"Dpi": dpi, "Dimensions": picture.size, "FractionOver32": fraction})
    results.append({"Report": prefix, "Pages": len(pdf.pages), "Scale": layout["Scale"],
                    "PdfSha256": hashlib.sha256(pdf_path.read_bytes()).hexdigest(),
                    "RequiredText": required, "Dpi": differences, "IndependentPdfComparison": comparisons, "ExcelPagination": excel_pagination})
(folder / "visual-validation.json").write_text(json.dumps({"Status": "PASS", "Results": results}, indent=2), encoding="utf-8")
print(json.dumps(results, indent=2))
