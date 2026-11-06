"""Fetch the macOS WebKit files omitted from the GTK release archive.

The checked-in macos-sources.json pins every file, its SHA-256 and the upstream
blob id. Files are downloaded from the same WebKit commit as the engine archive
and verified before they enter the ignored source cache.
"""
import hashlib
import json
import pathlib
import subprocess

HERE = pathlib.Path(__file__).resolve().parent
REPO = HERE.parents[1]
ROOT = REPO / "artifacts/jsc/source/webkitgtk-2.54.2"
PIN = HERE / "macos-sources.json"
SOURCES = json.loads((HERE / "sources.json").read_text(encoding="utf-8"))
PINNED = json.loads(PIN.read_text(encoding="utf-8"))


def fetch(url):
    return subprocess.run(["curl", "-fsSL", "--http1.1", "--retry", "5", "--retry-all-errors", url],
                          check=True, capture_output=True).stdout


def main():
    if PINNED["commit"] != SOURCES["webkit"]["commit"]:
        raise ValueError("macOS supplement must match the WebKit release commit")
    for entry in PINNED["files"]:
        target = ROOT / entry["path"]
        if target.exists():
            if hashlib.sha256(target.read_bytes()).hexdigest() != entry["sha256"]:
                raise ValueError(f"Cached macOS source differs from its pin: {target}")
            continue
        data = fetch(f"https://raw.githubusercontent.com/WebKit/WebKit/{PINNED['commit']}/{entry['path']}")
        if hashlib.sha256(data).hexdigest() != entry["sha256"]:
            raise ValueError(f"Pinned macOS source mismatch: {entry['path']}")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
    print(f"Verified {len(PINNED['files'])} pinned macOS sources")


if __name__ == "__main__":
    main()
