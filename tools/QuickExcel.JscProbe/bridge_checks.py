"""J1 real C ABI checks, trusted repository fixtures only; no workbook loader."""
import argparse
import ctypes as C
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import textwrap
import time
import uuid

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
ROOT = REPO / 'artifacts/jsc'
BIN = ROOT / 'build/win-x64/bin'


def bootstrap():
    main = (REPO / 'QuickExcel.Scripting/JscSession.cs').read_text(encoding='utf-8').split('"""', 2)[1]
    files = (REPO / 'QuickExcel.Scripting/VirtualFileBootstrap.cs').read_text(encoding='utf-8').split('"""', 2)[1]
    return textwrap.dedent(main).replace('/*VIRTUAL_FILES*/', textwrap.dedent(files))


class Native:
    def __init__(self):
        self.dll_dir = os.add_dll_directory(str(BIN))
        self.lib = C.CDLL(str(BIN / 'JavaScriptCore.dll'))
        signatures = {
            'qx_create': (C.c_void_p, [C.c_char_p, C.POINTER(C.c_void_p)]),
            'qx_destroy': (None, [C.c_void_p]),
            'qx_call': (C.c_void_p, [C.c_void_p, C.c_char_p, C.c_char_p, C.POINTER(C.c_void_p)]),
            'qx_load_modules': (C.c_int, [C.c_void_p, C.c_char_p, C.POINTER(C.c_void_p)]),
            'qx_free_string': (None, [C.c_void_p]),
            'qx_was_interrupted': (C.c_int, [C.c_void_p]),
            'qx_is_usable': (C.c_int, [C.c_void_p]),
            'qx_engine_identity': (C.c_char_p, []),
        }
        for name, (result, args) in signatures.items():
            fn = getattr(self.lib, name)
            fn.restype, fn.argtypes = result, args
        error = C.c_void_p()
        self.handle = self.lib.qx_create(bootstrap().encode(), C.byref(error))
        assert self.handle, self.consume(error.value)
        self.calls = []

    def consume(self, pointer):
        if not pointer:
            return None
        try:
            return C.string_at(pointer).decode('utf-8')
        finally:
            self.lib.qx_free_string(pointer)

    def call(self, method, args):
        error = C.c_void_p()
        start = time.perf_counter()
        value = self.consume(self.lib.qx_call(self.handle, method.encode(), json.dumps(args).encode(), C.byref(error)))
        record = {'method': method, 'ms': (time.perf_counter()-start)*1000, 'error': self.consume(error.value),
                  'value': json.loads(value) if value is not None else None,
                  'interrupted': bool(self.lib.qx_was_interrupted(self.handle)),
                  'usable': bool(self.lib.qx_is_usable(self.handle))}
        self.calls.append(record)
        return record

    def load(self, modules, entries=None, actions=False):
        assert not self.call('prepareModules', [actions])['error']
        error = C.c_void_p()
        start = time.perf_counter()
        payload = {'modules': [{'id': k, 'source': v} for k, v in modules.items()], 'entries': entries or ['main']}
        ok = self.lib.qx_load_modules(self.handle, json.dumps(payload).encode(), C.byref(error))
        record = {'method': 'modules', 'ms': (time.perf_counter()-start)*1000, 'ok': bool(ok),
                  'error': self.consume(error.value), 'interrupted': bool(self.lib.qx_was_interrupted(self.handle)),
                  'usable': bool(self.lib.qx_is_usable(self.handle))}
        self.calls.append(record)
        if ok:
            assert not self.call('finishModules', [])['error']
        return record

    def close(self):
        self.lib.qx_destroy(self.handle)
        self.handle = None


def fixtures():
    return {
        'live-binding': ({'counter': 'export let n=1; export function inc(){n++}',
                          'main': "import {n,inc} from 'counter'; quickExcel.functions.register('LIVE',()=>{inc();return n});"}, [2, 3]),
        'cycle': ({'a': "import {b} from 'b'; export const a=()=>b()+1;",
                   'b': "import {a} from 'a'; export const b=()=>2; export const back=()=>a();",
                   'main': "import {a} from 'a'; quickExcel.functions.register('LIVE',a);"}, [3, 3]),
        'namespace-reexport': ({'a': 'export default 4; export const n=7;', 'b': "export {default as d,n} from 'a';",
                               'main': "import * as ns from 'b'; quickExcel.functions.register('LIVE',()=>ns.d+ns.n);"}, [11, 11]),
        'once-only': ({'shared': 'globalThis.count=(globalThis.count||0)+1; export const n=3;',
                       'a': "import {n} from 'shared'; export {n};", 'b': "import {n} from 'shared'; export {n};",
                       'unused': "throw new Error('unreachable');",
                       'main': "import 'a'; import 'b'; quickExcel.functions.register('LIVE',()=>globalThis.count);"}, [1, 1]),
        'tdz': ({'a': "import {b} from 'b'; export const a=b;", 'b': "import {a} from 'a'; export const b=a;", 'main': "import 'a';"}, None),
        'undeclared': ({'main': "import 'missing';"}, None),
        'relative': ({'main': "import './a';", 'a': 'export const n=1;'}, None),
        'empty-attributes': ({'main': "import 'a' with {};", 'a': 'export const n=1;'}, None),
        'attributes': ({'main': "import n from 'a' with {type:'json'};", 'a': '{}'}, None),
        'top-level-await': ({'main': 'await 0;'}, None),
        'user-microtask': ({'main': "Promise.resolve().then(()=>{while(true){}}); quickExcel.functions.register('LIVE',()=>1);"}, None),
        'async-function': ({'main': '(async()=>{await 0; while(true) {}})();'}, None),
        'dynamic-import': ({'main': "import('a');", 'a': 'while(true){}'}, None),
        'caught-dynamic-import': ({'main': "try { import('a'); } catch {}", 'a': 'export const n=1;'}, None),
        'load-timeout': ({'main': 'while(true){}'}, 'timeout'),
        'catch-timeout': ({'main': 'try {while(true){}} catch(e) { quickExcel.functions.register("LIVE",()=>99); }'}, 'timeout'),
    }


def child(name):
    native = Native()
    try:
        if name in fixtures():
            modules, expected = fixtures()[name]
            result = native.load(modules)
            if expected == 'timeout':
                assert not result['ok'] and result['interrupted'] and not result['usable'], result
                assert 800 <= result['ms'] <= 1600, result
            elif expected is None:
                assert not result['ok'] and result['error'] and not result['usable'], result
                if name in ('user-microtask', 'async-function'):
                    assert result['ms'] < 800 and not result['interrupted'], result
            else:
                assert result['ok'], result
                actual = [native.call('invoke', ['LIVE', []]) for _ in expected]
                assert [r['value'] for r in actual] == expected and all(not r['error'] for r in actual), actual
        elif name == 'entry-order':
            result = native.load({'a': "globalThis.order='a';", 'b': "globalThis.order+='b';",
                                  'main': "quickExcel.functions.register('ORDER',()=>order);"}, ['a','b','main'])
            assert result['ok'], result
            assert native.call('invoke', ['ORDER', []])['value'] == 'ab'
        elif name == 'later-entry-timeout':
            result = native.load({'a': 'globalThis.n=1;', 'b': 'while(true){}'}, ['a', 'b'])
            assert result['interrupted'] and not result['ok'] and 800 <= result['ms'] <= 1600, result
        elif name == 'invoke-timeout':
            assert native.load({'main': "quickExcel.functions.register('LOOP',()=>{while(true){}});"})['ok']
            result = native.call('invoke', ['LOOP', []])
            assert result['interrupted'] and not result['usable'] and result['error'] and 800 <= result['ms'] <= 1600, result
            assert native.call('logs', [])['error']
        elif name == 'invoke-microtask':
            assert native.load({'main': "quickExcel.functions.register('ASYNC',()=>{Promise.resolve().then(()=>{while(true){}});return 1;});"})['ok']
            result = native.call('invoke', ['ASYNC', []])
            assert result['error'] and not result['usable'] and not result['interrupted'] and result['ms'] < 800, result
        elif name == 'hostile-error':
            result = native.load({'main': 'throw {get stack(){while(true){}},get message(){while(true){}},toString(){while(true){}}};'})
            assert result['error'] and result['ms'] < 800 and not result['interrupted'], result
        elif name == 'error-location':
            assert native.load({'main': "quickExcel.functions.register('FAIL',()=>{\nthrow new Error('expected failure');\n});"})['ok']
            result = native.call('invoke', ['FAIL', []])
            assert result['error'] and 'expected failure' in result['error'] and 'main:2:' in result['error'], result
            assert not result['interrupted']
        elif name == 'v1':
            result = native.call('load', ["quickExcel.functions.register('ADD',(a,b)=>a+b);"])
            assert not result['error'], result
            assert native.call('invoke', ['ADD', [3,4]])['value'] == 7
        elif name == 'serialization-timeout':
            # The bridge bootstrap's logs() calls slice(), which may be changed by
            # user JS. This forces hostile toJSON during native serialization.
            assert native.load({'main': "Array.prototype.slice=function(){return {toJSON(){while(true){}}};};"})['ok']
            result = native.call('logs', [])
            assert result['interrupted'] and not result['usable'] and result['error'] and 800 <= result['ms'] <= 1600, result
        else:
            raise AssertionError(name)
        print(json.dumps({'name': name, 'passed': True, 'engine': native.lib.qx_engine_identity().decode(), 'calls': native.calls}))
    finally:
        native.close()  # Parent timeout also covers poisoned-VM cleanup.


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--child')
    args = parser.parse_args()
    if args.child:
        child(args.child)
        return
    output = ROOT / 'bridge-results' / (datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S-') + uuid.uuid4().hex[:8])
    output.mkdir(parents=True)
    identity = json.loads((BIN.parent/'identity.json').read_text(encoding='utf-8-sig'))
    for item in identity['binaries']:
        assert hashlib.sha256((BIN/item['name']).read_bytes()).hexdigest() == item['sha256'].lower(), item['name']
    names = list(fixtures()) + ['entry-order','later-entry-timeout','invoke-timeout','invoke-microtask','hostile-error','error-location','v1','serialization-timeout']
    failures = []
    for name in names:
        start = time.perf_counter()
        env = {k:v for k,v in os.environ.items() if not k.startswith(('JSC_', 'QX_PROBE_'))}
        try:
            process = subprocess.run([sys.executable, str(Path(__file__).resolve()), '--child', name], capture_output=True,
                                     env=env, timeout=5, encoding='utf-8', errors='replace')
            record = {'name':name,'exitCode':process.returncode,'stdout':process.stdout,'stderr':process.stderr[:16384]}
            if process.returncode: failures.append(name)
        except subprocess.TimeoutExpired:
            record = {'name':name,'exitCode':'host-timeout'}
            failures.append(name)
        record['wallMs'] = (time.perf_counter()-start)*1000
        (output/(name+'.json')).write_text(json.dumps(record,indent=2),encoding='utf-8')
        print(name,record['exitCode'],round(record['wallMs']),flush=True)
    summary = {'samples':len(names),'failures':failures,'identity':identity,
               'sourceSha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
               'bootstrapSha256':hashlib.sha256(bootstrap().encode()).hexdigest(),
               'memoryIssue':'JSC-MEM-001 open, deferred by user; not a progression gate'}
    (output/'summary.json').write_text(json.dumps(summary,indent=2),encoding='utf-8')
    print('Evidence:',output,'Failures:',failures)
    raise SystemExit(bool(failures))


if __name__ == '__main__':
    main()
