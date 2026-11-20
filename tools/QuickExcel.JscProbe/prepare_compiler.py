"""Package verified TypeScript inputs for the isolated compiler worker."""
import hashlib
import json
import subprocess
import sys
import uuid
from pathlib import Path
from prepare import archive_source, download, PINS

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / 'artifacts/jsc'
TS = ROOT / 'typescript/package/lib'
archive_source(download(PINS['typescript']), ROOT / 'typescript')
if '--sources-only' in sys.argv:
    raise SystemExit(0)
# Direct native/platform builds also prepare the same closed library catalog.
subprocess.run(['node', str(REPO / 'tools/QuickExcel.ScriptLibraries/build.mjs')], check=True)
payload = {
    'version': '5.9.3',
    'compiler': (TS / 'typescript.js').read_text(encoding='utf-8'),
    'frontend': (REPO / 'QuickExcel.Scripting/TypeScript/frontend.js').read_text(encoding='utf-8'),
    'libraries': {p.name: p.read_text(encoding='utf-8') for p in sorted(TS.glob('lib.*.d.ts'))
                  if p.name.startswith(('lib.es', 'lib.decorators'))},
    'declarations': (REPO / 'QuickExcel.Web/src/quickexcel.d.ts').read_text(encoding='utf-8'),
    'builtins': json.loads((REPO / 'artifacts/script-libraries/declarations.json').read_text(encoding='utf-8')),
    'builtinsIdentity': (REPO / 'artifacts/script-libraries/identity.txt').read_text(encoding='ascii').strip(),
}
target = ROOT / 'compiler'
target.mkdir(exist_ok=True)
data = json.dumps(payload, ensure_ascii=False, separators=(',', ':'), sort_keys=True).encode('utf-8')
for name, content in [('typescript.bundle.json', data), ('identity.txt', hashlib.sha256(data).hexdigest().upper().encode('ascii'))]:
    temporary = target / (name + '.' + uuid.uuid4().hex + '.tmp')
    temporary.write_bytes(content)
    temporary.replace(target / name)
print(f'TypeScript payload: {len(data)} bytes, SHA-256 {hashlib.sha256(data).hexdigest()}')
