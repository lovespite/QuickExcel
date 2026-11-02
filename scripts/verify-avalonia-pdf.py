"""Compare actual PDF.js canvas captures with independent MuPDF renders of the same E4 PDFs.

Usage: python scripts/verify-avalonia-pdf.py CAPTURE_DIRECTORY PYTHON_DEPENDENCIES
The dependencies are test-only PyMuPDF, Pillow and numpy under repository artifacts.
"""
import json
import sys
from pathlib import Path

folder = Path(sys.argv[1]).resolve()
sys.path.insert(0, str(Path(sys.argv[2]).resolve()))
import pymupdf
import numpy as np
from PIL import Image, ImageFilter

results = []
for name, count in [("e4-quickexcel", 1), ("e4-native", 2), ("e4-paged", 2)]:
    with pymupdf.open(folder / (name + ".pdf")) as document:
        assert len(document) == count
        for number, page in enumerate(document, 1):
            actual = Image.open(folder / f"{name}-{number}.png").convert("RGB")
            # PDF.js rounds its canvas extent and applies a separate output transform.
            # Use that actual transform rather than stretching the PDF to the rounded extent.
            metadata = json.loads((folder / f"{name}-{number}-canvas.json").read_text())
            reference_pixmap = page.get_pixmap(matrix=pymupdf.Matrix(metadata["scaleX"], metadata["scaleY"]), alpha=False)
            reference = Image.frombytes("RGB", (reference_pixmap.width, reference_pixmap.height), reference_pixmap.samples)
            assert abs(reference.width - actual.width) <= 8 and abs(reference.height - actual.height) <= 8, (reference.size, actual.size)
            extent = (0, 0, min(reference.width, actual.width), min(reference.height, actual.height))
            actual, reference = actual.crop(extent), reference.crop(extent)
            reference.save(folder / f"{name}-{number}-mupdf.png")
            a = np.asarray(actual, dtype=np.int16)
            r = np.asarray(reference, dtype=np.int16)
            # Allow one pixel of rasterization displacement and 64 levels of antialiasing.
            def mismatch(source, target):
                lo = np.asarray(target.filter(ImageFilter.MinFilter(3)), dtype=np.int16)
                hi = np.asarray(target.filter(ImageFilter.MaxFilter(3)), dtype=np.int16)
                return np.any((source < lo - 64) | (source > hi + 64), axis=2)
            color_bad = mismatch(a, reference) | mismatch(r, actual)
            # Canvas and MuPDF differ in antialiasing where table strokes meet cell
            # clips. Compare achromatic ink geometry in the same one-pixel neighborhood;
            # retain channel comparisons for colored text, fills, pictures and charts.
            ink_a = np.min(a, axis=2) < 245
            ink_r = np.min(r, axis=2) < 245
            near_a = np.asarray(Image.fromarray(ink_a.astype(np.uint8) * 255).filter(ImageFilter.MaxFilter(3))) > 0
            near_r = np.asarray(Image.fromarray(ink_r.astype(np.uint8) * 255).filter(ImageFilter.MaxFilter(3))) > 0
            gray = (np.ptp(a, axis=2) < 3) & (np.ptp(r, axis=2) < 3)
            bad = color_bad & (~gray | (ink_a & ~near_r) | (ink_r & ~near_a))
            occupied = np.any((a < 245) | (r < 245), axis=2)
            tiles = []
            raw_tiles = []
            for top in range(0, actual.height, 128):
                for left in range(0, actual.width, 128):
                    mask = occupied[top:top+128, left:left+128]
                    if mask.sum() >= 64:
                        tiles.append(float((bad[top:top+128, left:left+128] & mask).sum() / mask.sum()))
                        raw_tiles.append(float((color_bad[top:top+128, left:left+128] & mask).sum() / mask.sum()))
            worst = max(tiles, default=0)
            result = {"fixture": name, "page": number, "width": actual.width, "height": actual.height, "maxTileMismatch": worst, "rawChannelMismatch": max(raw_tiles, default=0), "passed": worst < 0.10, "text": page.get_text()}
            results.append(result)
passed = all(result["passed"] for result in results) and len(results) == 5
(folder / "independent-visual.json").write_text(json.dumps({"passed": passed, "engine": pymupdf.VersionBind, "method": "One pixel geometry tolerance; 64 channel levels; grayscale clip/stroke antialiasing uses ink coverage, colored content uses channel neighborhoods; 128px occupied tiles < 10% mismatch", "pages": results}, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"passed": passed, "pages": [{k: v for k, v in r.items() if k != "text"} for r in results]}, indent=2))
sys.exit(0 if passed else 1)
