"""Trusted J0 fixtures. No workbook input. Functional checks and sampling are separate."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
import platform
from pathlib import Path
import subprocess
import time
import uuid

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
ROOT = REPO / "artifacts" / "jsc"
TS = ROOT / "typescript" / "package" / "lib"
FRONTEND = REPO / "QuickExcel.Scripting/TypeScript/frontend.js"


def module(source, name="main", language="typescript"):
    return {"id": name, "language": language, "source": source}


def cases():
    return [
        ("generic", [module("export const identity = <T>(x:T):T => x;\nquickExcel.functions.register('JS.ID', () => identity(3));")], True),
        ("cross-module", [module("export interface Box<T> { value:T }; export const box = <T>(value:T):Box<T> => ({value});", "shared"),
            module("import {box} from 'shared'; import type {Box} from 'shared'; const n:Box<number> = box(3); export const value = n.value;")], True),
        ("cross-module-error", [module("export function twice(n:number):number { return n*2; }", "shared"),
            module("import {twice} from 'shared'; export const value = twice('wrong');")], False),
        ("enum-class-namespace", [module("enum E { A=3,B }; namespace N { export const value=E.B; } class C { x=N.value; }; export const value = new C().x;")], True),
        ("mixed-js-ts", [module("export const twice = n => n * 2;", "shared", "javascript"),
            module("import {twice} from 'shared'; export const value:number = twice(3);")], True),
        ("no-execution", [module("throw new Error('MUST NOT EXECUTE'); export const value:number = 7;")], True),
        ("type-error", [module("export const value: number = 'wrong';")], False),
        ("syntax-error", [module("export const value: = ;")], False),
        ("undeclared", [module("import {x} from 'missing'; export const value = x;")], False),
        ("relative", [module("import {x} from './shared'; export const value = x;"), module("export const x=1;", "shared")], False),
        ("no-dom", [module("export const value = document.title;")], False),
        ("no-node", [module("export const value = process.env.PATH;")], False),
        ("type-import", [module("export type X = {n:number};", "shared"), module("import type {X} from 'shared'; export const value:X = {n:1};")], True),
        ("dynamic-import", [module("export const value = import('main');")], False),
        ("top-level-await", [module("export const value = await Promise.resolve(3);")], False),
        ("import-attributes", [module("import {x} from 'shared' with {type:'json'}; export {x};"), module("export const x=1;", "shared")], False),
        ("decorators", [module("function dec(x:any) {} @dec class X {} export {X};")], False),
        ("external-reference", [module('/// <reference path="C:/outside.d.ts" />\nexport const value = 1;')], False),
        ("unicode-budget", [module("//" + "\u4e2d" * 350000)], False),
        ("generated-budget", [module("export enum E {" + ",".join("Value" + str(i) for i in range(24000)) + "}")], False),
        ("diagnostic-budget", [module('\n'.join(f"export const v{i}:number = 'wrong';" for i in range(1000)))], False),
    ]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--performance", action="store_true")
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--compiler-seconds", type=float, default=5)
    parser.add_argument("--compiler-memory-mib", type=int, default=768)
    args = parser.parse_args()
    if not 1 <= args.repeats <= 20:
        parser.error("repeats must be between 1 and 20")
    if not 0 < args.compiler_seconds <= 30 or not 16 <= args.compiler_memory_mib <= 1024:
        parser.error('compiler ceiling must be > 0..30 seconds and 16..1024 MiB')
    output = ROOT / "results" / (datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:8])
    output.mkdir(parents=True)
    executable = ROOT / "build" / "win-x64" / "bin" / "quickexcel-jsc-probe.exe"
    identity_path = executable.parent.parent / 'identity.json'
    identity_bytes = identity_path.read_bytes()
    identity = json.loads(identity_bytes.decode('utf-8-sig'))
    for item in identity['binaries']:
        if hashlib.sha256((executable.parent / item['name']).read_bytes()).hexdigest() != item['sha256'].lower():
            raise ValueError('Build payload differs from recorded identity; rebuild before sampling')
    for item in identity['probeSources']:
        if hashlib.sha256((HERE / item['name']).read_bytes()).hexdigest() != item['sha256'].lower():
            raise ValueError('Probe sources differ from recorded build; rebuild before sampling')
    records = []

    def invoke(name, sources, seconds=1, environment=None):
        paths = []
        for index, source in enumerate(sources):
            if isinstance(source, Path):
                paths.append(source)
            else:
                path = output / f"{name}-{index}.js"
                path.write_text(source, encoding="utf-8")
                paths.append(path)
        env = {k: v for k, v in os.environ.items() if not k.startswith(("JSC_", "QX_PROBE_"))}
        candidate_options = {"JSC_jitMemoryReservationSize": "16777216",
                             "JSC_maxPerThreadStackUsage": "262144",
                             "QX_PROBE_PROCESS_LIMIT_MIB": "256", **(environment or {})}
        env.update(candidate_options)
        started = time.perf_counter()
        host_timeout = seconds + (10 if candidate_options.get('QX_PROBE_CPU_WATCHDOG') else 1)
        record = {"name": name, "seconds": seconds, "hostTimeoutSeconds": host_timeout, "environment": candidate_options}
        try:
            process = subprocess.run([str(executable), str(seconds), *map(str, paths)], env=env,
                capture_output=True, timeout=host_timeout, encoding="utf-8", errors="replace")
            record.update(exitCode=process.returncode, stderr=process.stderr[:16384],
                          native=json.loads(process.stdout) if process.stdout.startswith("{") else process.stdout[:16384])
        except subprocess.TimeoutExpired:
            record.update(exitCode="host-timeout", native=None)
        record["wallMilliseconds"] = (time.perf_counter() - started) * 1000
        records.append(record)
        (output / f"{name}.json").write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"{name}: {record['exitCode']}, {record['wallMilliseconds']:.1f} ms", flush=True)
        return record

    failures = []
    if not args.performance:
        r = invoke("scalar", ["[null, true, 42, 'text', Object.is(-0, -0)]"])
        if not isinstance(r.get("native"), dict) or r["native"].get("result") != [None, True, 42, "text", True]: failures.append("scalar")
        # Both outcomes are evidence. Setup failure is not a valid memory-budget observation.
        r = invoke("process-256m-observation", ["42;"], environment={"QX_PROBE_PROCESS_LIMIT_MIB": "256"})
        if r["exitCode"] in (68, "host-timeout") or "probe-job-limit-installed=268435456" not in r.get("stderr", ""):
            failures.append("process-256m-setup")
        r = invoke('process-256m-overflow', ['globalThis.keep = new Uint8Array(300 * 1024 * 1024); keep.fill(7); keep.byteLength;'],
                   environment={"QX_PROBE_PROCESS_LIMIT_MIB": "256"})
        if (r['exitCode'] in (0, 68, 'host-timeout') or
                'probe-job-limit-installed=268435456' not in r.get('stderr', '')):
            failures.append('process-memory-enforcement')
        r = invoke("interrupt", ["while(true) {}"])
        if not isinstance(r["native"], dict) or r["native"].get("interrupted") is not True: failures.append("interrupt")
        invoke("upstream-cpu-time-observation", ["while(true) {}"], environment={"QX_PROBE_CPU_WATCHDOG": "1"})
        r = invoke("hostile-exception", ["throw {get stack(){while(true){}}, toString(){while(true){}}};"])
        if r["exitCode"] != 65: failures.append("hostile-exception")
        # Observation, not an acceptance test for a product memory budget.
        r = invoke("gc-threshold-is-not-cap", ["globalThis.keep = new Uint8Array(96 * 1024 * 1024); keep.fill(7); keep.byteLength;"],
               environment={"JSC_gcMaxHeapSize": "67108864"})
        if not isinstance(r.get("native"), dict) or r["native"].get("gcMaxHeapSize") != 67108864:
            failures.append("gc-threshold-option")
        invoke("microtask-observation", ["globalThis.ran = false; Promise.resolve().then(() => {ran=true;}); true;", "ran;"])
        r = invoke("microtask-interrupt", ["Promise.resolve().then(function again(){Promise.resolve().then(again);}); true;"])
        if not isinstance(r["native"], dict) or r["native"].get("interrupted") is not True: failures.append("microtask-interrupt")
        r = invoke("stack-limit", ["function recurse(){return 1+recurse()}; recurse();"],
                   environment={"JSC_maxPerThreadStackUsage": "262144"})
        if (r["exitCode"] != 65 or not isinstance(r.get("native"), dict) or
                r["native"].get("maxPerThreadStackUsage") != 262144): failures.append("stack-limit")
    libraries = {p.name: p.read_text(encoding="utf-8") for p in TS.glob("lib.*.d.ts")
                 if p.name.startswith(("lib.es", "lib.decorators"))}
    declarations = (REPO / "QuickExcel.Web" / "src" / "quickexcel.d.ts").read_text(encoding="utf-8")
    def dense(size):
        lines = []; length = 0
        while True:
            line = f"export const v{len(lines)}:number={len(lines)};\n"
            if length + len(line) > size: break
            lines.append(line); length += len(line)
        return ''.join(lines)

    fixtures = cases() if not args.performance else [
        ("small", [module("export const value:number = 3;")], True),
        ("padded-128k", [module("export const value:number=3;\n//" + "x" * (128 * 1024 - 128))], True),
        ("padded-1m", [module("export const value:number=3;\n//" + "x" * (1024 * 1024 - 128))], True),
        ("64-modules", [module(f"export const value:number = {i};", "m" + str(i)) for i in range(64)], True),
        ("dense-128k", [module(dense(128 * 1024))], True),
        ("dense-1m", [module(dense(1024 * 1024))], True),
    ]
    for repeat in range(args.repeats if args.performance else 1):
        for name, modules, success in fixtures:
            for mode in ("transpile", "check") if args.performance else ("check",):
                request = {"mode": mode, "modules": modules, "entries": [modules[-1]["id"]],
                           "libraries": libraries, "declarations": declarations}
                r = invoke(f"ts-{name}-{mode}-{repeat}", [TS / "typescript.js", FRONTEND,
                    "qxCompile(" + json.dumps(request, ensure_ascii=True) + ");"], seconds=args.compiler_seconds,
                    environment={"QX_PROBE_PROCESS_LIMIT_MIB": str(args.compiler_memory_mib),
                                 "JSC_maxPerThreadStackUsage": "5242880"})
                native = r.get("native")
                result = native.get("result") if isinstance(native, dict) else None
                if (r["exitCode"] != 0 or not isinstance(result, dict) or result.get("success") != success or
                        native.get('processLimitBytes') != args.compiler_memory_mib * 1024 * 1024):
                    failures.append(r["name"])
                elif not success and result.get('outputs'):
                    failures.append(r['name'] + '-partial-output')
                elif result["success"]:
                    for emitted in result["outputs"]:
                        source_map = json.loads(emitted["sourceMap"])
                        if (source_map.get("version") != 3 or not source_map.get("sourcesContent") or
                                source_map.get("sources") != [emitted["id"]] or source_map.get("sourceRoot") != ''):
                            failures.append(r["name"] + "-map")
                    if name == "type-import" and any("from " in e["javascript"] for e in result["outputs"]):
                        failures.append(r["name"] + "-type-erasure")
                elif name == 'diagnostic-budget':
                    if (len(result['diagnostics']) > 128 or
                        len(json.dumps(result['diagnostics'], ensure_ascii=False, separators=(',', ':')).encode('utf-8')) > 16 * 1024 or
                        not any('Additional diagnostics omitted' in d['message'] for d in result['diagnostics'])):
                        failures.append(r['name'] + '-unbounded')
    if identity_path.read_bytes() != identity_bytes:
        failures.append('build-changed-during-sampling')
    summary = {"status": "development-evidence-not-migration-acceptance", "migrationReady": False, "performance": args.performance,
               "failures": failures, "samples": len(records),
               "compilerBudget": {"seconds": args.compiler_seconds, "processMiB": args.compiler_memory_mib,
                                  "stackBytes": 5242880, "hostTimeoutSeconds": args.compiler_seconds + 1},
               "machine": {"platform": platform.platform(), "processor": platform.processor(),
                           "logicalProcessors": os.cpu_count(), "python": platform.python_version()},
               "identity": identity,
               "standardLibrariesSha256": hashlib.sha256(json.dumps(libraries, sort_keys=True).encode('utf-8')).hexdigest(),
               "inputs": {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in
                   (HERE / "run.py", HERE / "probe.cpp", FRONTEND, TS / "typescript.js", REPO / "QuickExcel.Web/src/quickexcel.d.ts")}}
    (output / "summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(f"Evidence: {output}; failures: {failures}")
    raise SystemExit(bool(failures))


if __name__ == "__main__":
    main()
