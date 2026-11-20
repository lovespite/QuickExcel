"""Retain notices and exact corresponding sources/build adjustments for redistribution."""
import hashlib
import json
import urllib.request
import zipfile
from pathlib import Path, PurePosixPath

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / 'artifacts/jsc'
HERE = Path(__file__).resolve().parent
OUT = ROOT / 'distribution'


def prepare_supplements():
    """A clean runner has only its own platform's supplemental WebKit files."""
    pins = json.loads((HERE / 'sources.json').read_text(encoding='utf-8'))
    source = ROOT / 'source' / ('webkitgtk-' + pins['webkit']['version'])
    files = []
    for name in ('windows-sources.json', 'macos-sources.json'):
        supplement = json.loads((HERE / name).read_text(encoding='utf-8'))
        if supplement['commit'] != pins['webkit']['commit']:
            raise ValueError(f'{name} must match the WebKit release commit')
        for entry in supplement['files']:
            relative = PurePosixPath(entry['path'])
            if relative.is_absolute() or '..' in relative.parts or '\\' in entry['path'] or ':' in entry['path']:
                raise ValueError(f'Invalid supplemental source path: {entry["path"]}')
            target = source / entry['path']
            if target.is_symlink() or source.resolve() not in target.resolve().parents:
                raise ValueError(f'Supplemental source escapes its cache: {target}')
            if target.exists():
                data = target.read_bytes()
            else:
                url = f"https://raw.githubusercontent.com/WebKit/WebKit/{supplement['commit']}/{entry['path']}"
                with urllib.request.urlopen(url, timeout=60) as response:
                    data = response.read()
            if hashlib.sha256(data).hexdigest() != entry['sha256']:
                raise ValueError(f'Pinned supplemental source mismatch: {target}')
            if not target.exists():
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(data)
            files.append(target)
    return files


# Fetch and verify both supplements before creating any distributable output.
supplemental_sources = prepare_supplements()
OUT.mkdir(exist_ok=True)
licenses = []
for origin in (ROOT / 'source/webkitgtk-2.54.2', ROOT / 'source/icu', ROOT / 'typescript/package'):
    for path in sorted(origin.rglob('*')):
        if path.is_file() and any(word in path.name.lower() for word in ('license', 'copying', 'copyright', 'thirdpartynotice')):
            if path.stat().st_size <= 1024 * 1024:
                licenses.append((path.relative_to(ROOT).as_posix(), path))
with (OUT / 'JSC-THIRD-PARTY-NOTICES.txt').open('w', encoding='utf-8') as target:
    target.write('QuickExcel JSC: JavaScriptCore/WebKit 2.54.2, ICU 77.1, mimalloc 3.2.8, TypeScript 5.9.3.\n'
                 'Copyrights and licenses retain their individual scope. This is not a single MIT-licensed dependency.\n'
                 'Windows also carries Microsoft Visual C++ 14.44 redistributables from the installed Visual Studio; those use Microsoft terms.\n'
                 'macOS arm64 links the pinned JavaScriptCore inside the worker payload; WebKit advises its Apple port to use the OS libicucore ICU.\n'
                 'Original source headers and additional notices are retained in the accompanying source archives.\n\n')
    for name, path in licenses:
        target.write('\n===== ' + name + ' =====\n' + path.read_text(encoding='utf-8', errors='replace') + '\n')
with zipfile.ZipFile(OUT / 'JSC-CORRESPONDING-SOURCE.zip', 'w', zipfile.ZIP_DEFLATED) as archive:
    for name in ('webkitgtk-2.54.2.tar.xz', 'icu4c-77_1-src.zip', 'typescript-5.9.3.tgz'):
        archive.write(ROOT / 'downloads' / name, 'archives/' + name, compress_type=zipfile.ZIP_STORED)
    for path in supplemental_sources:
        archive.write(path, path.relative_to(REPO).as_posix())
    for path in sorted(HERE.iterdir()):
        if path.suffix in ('.json', '.py', '.ps1', '.sh', '.cmake', '.txt', '.cpp', '.md'):
            archive.write(path, path.relative_to(REPO).as_posix())
    for path in (REPO / 'QuickExcel.Scripting/native/jsc').rglob('*'):
        if path.is_file(): archive.write(path, path.relative_to(REPO).as_posix())
    for path in (REPO / 'QuickExcel.Scripting/TypeScript/frontend.js', REPO / 'QuickExcel.Web/src/quickexcel.d.ts'):
        archive.write(path, path.relative_to(REPO).as_posix())
(OUT / 'JSC-REBUILD.txt').write_text(
    'Corresponding source and QuickExcel bridge/build adjustments are supplied in JSC-CORRESPONDING-SOURCE.zip.\n'
    'Extract into a QuickExcel checkout and put archives in artifacts/jsc/downloads. Run tools/QuickExcel.JscProbe/build.ps1\n'
    'on Windows, build-linux.sh on Linux or build-macos.sh on macOS arm64. Requirements and pinned hashes are in sources.json and the README.\n'
    'The bridge is dynamically linked. Modified libraries can replace the same-named libraries in a private installation.\n'
    'Rebuild the managed host/worker if the C ABI or identity is changed; update local package hashes/active pointer or launch that private executable directly.\n'
    'Do not mix host and worker protocol versions. No reverse-engineering restriction is imposed for debugging modifications to LGPL components.\n'
    'Keep an original complete package for rollback. Older packages reject the typescript-sources feature.\n'
    'OPEN ISSUE JSC-MEM-001: 64 MiB total engine allocation cap is not implemented and is deferred.\n'
    'Execution keeps the 256 MiB process cap, 256 KiB JS stack, 1-second native operation deadline and 5-second round deadline.\n'
    'Linux enforces the cap through cgroup v2; macOS arm64 uses a native polling footprint/thread watchdog because Darwin\n'
    'rejects RLIMIT_AS/DATA and jetsam limits require private entitlements. The macOS bound is not cgroup-equivalent.\n', encoding='utf-8')
print('Prepared corresponding sources, build instructions and notices: ' + str(OUT))
