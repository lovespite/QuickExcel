"""Fetch hash-pinned J0 inputs into this repository's ignored artifacts directory."""
import hashlib
import json
from pathlib import Path, PurePosixPath
import shutil
import subprocess
import tarfile
import urllib.request
import zipfile

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
ROOT = REPO / "artifacts" / "jsc"
PINS = json.loads((HERE / "sources.json").read_text(encoding="utf-8"))
ADJUSTMENTS = json.loads((HERE / 'source-adjustments.json').read_text(encoding='utf-8'))


def artifact_path(path):
    if not path.resolve().is_relative_to(ROOT.resolve()):
        raise ValueError('Artifact path escapes the JSC workspace')
    return path


def file_digest(stream, algorithm):
    digest = hashlib.new(algorithm)
    for chunk in iter(lambda: stream.read(1 << 20), b""):
        digest.update(chunk)
    return digest


def is_junction(path):
    return getattr(path, "is_junction", lambda: False)()


def download(pin):
    path = artifact_path(ROOT / "downloads" / pin["url"].rsplit("/", 1)[1])
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        partial = artifact_path(path.with_suffix(path.suffix + ".partial"))
        print(f"Downloading {path.name}", flush=True)
        with urllib.request.urlopen(pin["url"], timeout=300) as response, partial.open("wb") as out:
            shutil.copyfileobj(response, out)
        partial.replace(path)
    algorithm = "sha512" if "sha512" in pin else "sha256"
    with path.open("rb") as stream:
        actual = file_digest(stream, algorithm).hexdigest()
    if actual != pin[algorithm]:
        raise ValueError(f"Pinned {algorithm} mismatch: {path}. No extraction or build permitted.")
    return path


def archive_source(archive, destination, webkit=False):
    """Extract verified archives safely, and verify existing source cache against the archive."""
    if not destination.resolve().is_relative_to(REPO):
        raise ValueError('Source destination escapes this repository')
    destination.mkdir(parents=True, exist_ok=True)
    checked_directories = set()
    def target_for(name):
        relative = PurePosixPath(name)
        if relative.is_absolute() or '..' in relative.parts or '\\' in name or ':' in name:
            raise ValueError("Archive path escapes destination")
        target = destination / name
        if target.is_symlink() or is_junction(target):
            raise ValueError("Source cache contains a reparse point")
        parent = target.parent
        while parent != destination.parent and parent not in checked_directories:
            if parent.is_symlink() or is_junction(parent):
                raise ValueError('Source cache contains a directory reparse point')
            checked_directories.add(parent)
            parent = parent.parent
        return target

    if archive.suffix == ".zip":
        with zipfile.ZipFile(archive) as source:
            for entry in source.infolist():
                if entry.is_dir():
                    continue
                target = target_for(entry.filename)
                content = source.read(entry)
                if target.exists():
                    if target.read_bytes() != content:
                        raise ValueError(f"Source cache differs from pinned archive: {target}")
                else:
                    target.parent.mkdir(parents=True, exist_ok=True)
                    target.write_bytes(content)
    else:
        with tarfile.open(archive, mode='r|*') as source:
            for entry in source:
                if entry.isdir():
                    continue
                if webkit:
                    relative = entry.name.partition('/')[2]
                    if not (relative in ('CMakeLists.txt', 'Source/CMakeLists.txt', 'Tools/CMakeLists.txt') or
                            relative.startswith(('Source/cmake/', 'Source/bmalloc/', 'Source/WTF/',
                                'Source/JavaScriptCore/', 'Source/ThirdParty/unifdef/', 'Tools/Scripts/'))):
                        continue
                target = target_for(entry.name)
                if not entry.isfile():
                    raise ValueError(f"Unexpected archive path/type: {entry.name}")
                content = source.extractfile(entry).read()
                for adjustment in ADJUSTMENTS if webkit else ():
                    if adjustment['path'] == entry.name:
                        before, after = adjustment['before'].encode(), adjustment['after'].encode()
                        if content.count(before) != 1:
                            raise ValueError('Source adjustment no longer applies to the pinned source')
                        adjusted = content.replace(before, after)
                        if target.exists() and target.read_bytes() == content:
                            target.write_bytes(adjusted)
                        content = adjusted
                if target.exists():
                    if target.read_bytes() != content:
                        raise ValueError(f"Source cache differs from pinned archive: {target}")
                else:
                    target.parent.mkdir(parents=True, exist_ok=True)
                    target.write_bytes(content)


def main():
    if not ROOT.resolve().is_relative_to(REPO):
        raise ValueError('Artifact destination escapes this repository')
    archives = {name: download(PINS[name]) for name in ("webkit", "icu", "ruby", "llvm", "typescript")}
    for name in ("webkit", "icu"):
        print(f"Verifying extracted {name} sources", flush=True)
        archive_source(archives[name], ROOT / "source", webkit=name == "webkit")
    windows = json.loads((HERE / 'windows-sources.json').read_text(encoding='utf-8'))
    if windows['commit'] != PINS['webkit']['commit']:
        raise ValueError('Windows supplement must match the WebKit release commit')
    for entry in windows['files']:
        target = artifact_path(ROOT / 'source' / ('webkitgtk-' + PINS['webkit']['version']) / entry['path'])
        data = target.read_bytes() if target.exists() else urllib.request.urlopen(
            f"https://raw.githubusercontent.com/WebKit/WebKit/{windows['commit']}/{entry['path']}", timeout=60).read()
        if hashlib.sha256(data).hexdigest() != entry['sha256']:
            raise ValueError(f'Pinned Windows source mismatch: {target}')
        if not target.exists():
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data)
    archive_source(archives["typescript"], ROOT / "typescript")
    seven_zip = shutil.which("7z") or r"C:\Program Files\7-Zip\7z.exe"
    for name, executable in (("ruby", "rubyinstaller-4.0.9-1-x64/bin/ruby.exe"), ("llvm", "bin/clang-cl.exe")):
        destination = artifact_path(ROOT / "tools" / name)
        if not (destination / executable).exists():
            subprocess.run([seven_zip, "x", str(archives[name]), f"-o{destination}", "-y"], check=True,
                           stdout=subprocess.DEVNULL)
    print(f"Verified J0 source archives and source caches: {ROOT}")


if __name__ == "__main__":
    main()
