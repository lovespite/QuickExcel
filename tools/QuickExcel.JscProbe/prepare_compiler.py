"""Package verified TypeScript inputs for the isolated compiler worker."""
import hashlib
import json
from pathlib import Path
from prepare import archive_source, download, PINS

REPO = Path(__file__).resolve().parents[2]
ROOT = REPO / 'artifacts/jsc'
TS = ROOT / 'typescript/package/lib'
archive_source(download(PINS['typescript']), ROOT / 'typescript')
payload = {
    'version': '5.9.3',
    'compiler': (TS / 'typescript.js').read_text(encoding='utf-8'),
    'frontend': (REPO / 'QuickExcel.Scripting/TypeScript/frontend.js').read_text(encoding='utf-8'),
    'libraries': {p.name: p.read_text(encoding='utf-8') for p in sorted(TS.glob('lib.*.d.ts'))
                  if p.name.startswith(('lib.es', 'lib.decorators'))},
    'declarations': (REPO / 'QuickExcel.Web/src/quickexcel.d.ts').read_text(encoding='utf-8'),
}
target = ROOT / 'compiler'
target.mkdir(exist_ok=True)
data = json.dumps(payload, ensure_ascii=False, separators=(',', ':'), sort_keys=True).encode('utf-8')
(target / 'typescript.bundle.json').write_bytes(data)
(target / 'identity.txt').write_text(hashlib.sha256(data).hexdigest().upper(), encoding='ascii')
print(f'TypeScript payload: {len(data)} bytes, SHA-256 {hashlib.sha256(data).hexdigest()}')
