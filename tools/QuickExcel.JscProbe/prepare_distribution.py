"""Retain notices and exact corresponding sources/build adjustments for redistribution."""
import json
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / 'artifacts/jsc'
HERE = Path(__file__).resolve().parent
OUT = ROOT / 'distribution'
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
                 'Original source headers and additional notices are retained in the accompanying source archives.\n\n')
    for name, path in licenses:
        target.write('\n===== ' + name + ' =====\n' + path.read_text(encoding='utf-8', errors='replace') + '\n')
with zipfile.ZipFile(OUT / 'JSC-CORRESPONDING-SOURCE.zip', 'w', zipfile.ZIP_DEFLATED) as archive:
    for name in ('webkitgtk-2.54.2.tar.xz', 'icu4c-77_1-src.zip', 'typescript-5.9.3.tgz'):
        archive.write(ROOT / 'downloads' / name, 'archives/' + name, compress_type=zipfile.ZIP_STORED)
    for entry in json.loads((HERE / 'windows-sources.json').read_text())['files']:
        path = ROOT / 'source/webkitgtk-2.54.2' / entry['path']
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
    'on Windows or build-linux.sh on Linux. Requirements and pinned hashes are in sources.json and the README.\n'
    'The bridge is dynamically linked. Modified libraries can replace the same-named libraries in a private installation.\n'
    'Rebuild the managed host/worker if the C ABI or identity is changed; update local package hashes/active pointer or launch that private executable directly.\n'
    'Do not mix host and worker protocol versions. No reverse-engineering restriction is imposed for debugging modifications to LGPL components.\n'
    'Keep an original complete package for rollback. Older packages reject the typescript-sources feature.\n'
    'OPEN ISSUE JSC-MEM-001: 64 MiB total engine allocation cap is not implemented and is deferred.\n'
    'Execution keeps the 256 MiB process cap, 256 KiB JS stack, 1-second native operation deadline and 5-second round deadline.\n', encoding='utf-8')
print('Prepared corresponding sources, build instructions and notices: ' + str(OUT))
