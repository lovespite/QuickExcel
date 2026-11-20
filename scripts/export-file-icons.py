"""Export the document SVG artwork to Windows ICOs (Python 3 + ImageMagick 7).

The ICO contains 32-bit PNG frames, matching the established XLSQ icon. Sizes
below 32 px use the simplified artwork to keep the document silhouette legible.
Run from any directory; temporary raster images stay under artifacts.
"""

import argparse
from pathlib import Path
import struct
import subprocess
import tempfile


ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "QuickExcel.Desktop.Avalonia" / "Assets"
SIZES = (256, 128, 96, 64, 48, 40, 32, 24, 20, 16)


def export(extension: str) -> None:
    scratch = ROOT / "artifacts" / "file-icons"
    scratch.mkdir(parents=True, exist_ok=True)
    frames = []
    with tempfile.TemporaryDirectory(prefix=extension.lower() + "-", dir=scratch) as temp:
        for size in SIZES:
            source = ASSETS / f"QuickExcel-{extension}{'-small' if size < 32 else ''}.svg"
            destination = Path(temp) / f"{size}.png"
            subprocess.run([
                "magick", "-background", "none", "-density", "384", str(source),
                "-resize", f"{size}x{size}", "-depth", "8", "-strip",
                "-define", "png:color-type=6", str(destination),
            ], check=True)
            frames.append(destination.read_bytes())
    entries = []
    offset = 6 + len(SIZES) * 16
    for size, png in zip(SIZES, frames):
        entries.append(struct.pack("<BBBBHHII", size % 256, size % 256, 0, 0,
                                   1, 32, len(png), offset))
        offset += len(png)
    output = ASSETS / f"QuickExcel-{extension}.ico"
    output.write_bytes(struct.pack("<HHH", 0, 1, len(SIZES)) + b"".join(entries) + b"".join(frames))
    print(f"Exported {output.relative_to(ROOT)}: {len(SIZES)} RGBA frames")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("extensions", nargs="*", metavar="XLSQ|XLSX|CSV")
    extensions = parser.parse_args().extensions or ["XLSX", "CSV"]
    if any(extension not in ("XLSQ", "XLSX", "CSV") for extension in extensions):
        parser.error("extensions must be XLSQ, XLSX or CSV")
    for extension in extensions:
        export(extension)
