"""Validate reporting-smoke artifacts. Requires Pillow and pypdf; no workbook mutation."""
import hashlib
import json
import math
import re
import sys
from pathlib import Path

from PIL import Image, ImageChops
from pypdf import PdfReader


folder = Path(sys.argv[1]).resolve()
results = []
reports = [("quickexcel", ["123.5", "Merged report title", "E4 revenue"]), ("native", ["Report 51", "Report 4", "Report 5", "Revenue", "Cost"])]
if (folder / "paged.pdf").exists():
    reports.append(("paged", ["Merged title", "Record 19", "Record 20", "Record 35"]))
for prefix, required in reports:
    layout = json.loads((folder / f"{prefix}-layout.json").read_text(encoding="utf-8-sig"))
    pdf_path = folder / f"{prefix}.pdf"
    pdf = PdfReader(pdf_path)
    assert len(pdf.pages) == len(layout["Pages"]), "PDF/preview page count differs"
    text = " ".join(" ".join(page.extract_text().split()) for page in pdf.pages)
    excel_pagination = None
    for label in required:
        assert label.replace(" ", "") in text.replace(" ", ""), f"Missing PDF text: {label}"
    if prefix == "paged" and len(sys.argv) > 2:
        reference = PdfReader(Path(sys.argv[2]) / "excel-output-paged.pdf")
        records = lambda document: [re.findall(r"Record\s*(\d+)", page.extract_text()) for page in document.pages]
        assert records(pdf) == records(reference), "Record membership on pages differs from Excel PDF"
        excel_pagination = {"Matches": True, "RecordsByPage": records(pdf), "Reference": str(Path(sys.argv[2]).resolve() / "excel-output-paged.pdf")}
    for page in pdf.pages:
        assert abs(float(page.mediabox.width) - layout["PaperWidth"] * .75) < .2
        assert abs(float(page.mediabox.height) - layout["PaperHeight"] * .75) < .2
    baseline = Image.open(folder / f"{prefix}-96.png").convert("RGB")
    differences = []
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
                    "RequiredText": required, "Dpi": differences, "ExcelPagination": excel_pagination})
(folder / "visual-validation.json").write_text(json.dumps({"Status": "PASS", "Results": results}, indent=2), encoding="utf-8")
print(json.dumps(results, indent=2))
