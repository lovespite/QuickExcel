# JSC / TypeScript build and development probes

These tools implement the pinned source build and development evidence for
[plan 70](../../QuickExcel.Desktop/docs/70-jsc-typescript-migration-plan.md).
The `quickexcel-jsc-probe` executable is development-only: it reads trusted local
fixtures and is never a workbook execution path. The product uses the
bridge in `QuickExcel.Scripting/native/jsc/bridge.cpp` and isolated workers.
Current results are in [record 71](../../QuickExcel.Desktop/docs/71-jsc-typescript-implementation-log.md).

From the repository root on Windows x64:

```powershell
.\tools\QuickExcel.JscProbe\build.ps1
python .\tools\QuickExcel.JscProbe\bridge_checks.py
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build-jsc -p:QuickExcelScriptEngine=Jsc --filter 'FullyQualifiedName~JscBridgeTests|FullyQualifiedName~TypeScriptCompilerTests|FullyQualifiedName~TypeScriptStorageTests'
.\scripts\build-avalonia-release.ps1 -ScriptEngine Jsc -RunSmoke
```

Windows prerequisites: Python 3.13, CMake 4.3, Visual Studio C++ tools
v143 14.44.35207 and Windows SDK, Git for Windows Perl, and 7-Zip.
Pinned LLVM 20.1.8 and Ruby 4.0.9 are extracted inside this repository.
ICU 77.1 and JSC are source-built. `-PrepareOnly` verifies inputs.
The build copies required Microsoft VC Runtime 14.44 redistributables;
the build identity records their actual hashes. No global tool installation is done.

Linux x64, from the repository root:

```sh
QX_BUILD_PARALLEL=2 bash tools/QuickExcel.JscProbe/build-linux.sh
bash QuickExcel.ScriptWorker/native/build-linux.sh
bash tools/QuickExcel.PlatformSmoke/native/build-linux.sh
```

Linux prerequisites: Python >= 3.12, clang/clang++ 18, CMake >= 3.28, Ninja,
make, GNU development headers, libyaml development headers, pkg-config and Perl. The script builds pinned
Ruby 4.0.9 as a local build tool and ICU 77.1; no system Ruby/ICU is shipped.
The ICU ZIP contains CRLF build inputs, so a separate reproducible LF-normalized
copy is built while the verified original cache stays intact. Header maps are
disabled because the GTK archive omits hmaptool. Linux validation must run as a
normal user with Landlock ABI >= 3 and a systemd user session/cgroup v2.
Building the native library alone is not platform acceptance.

macOS 26 arm64, from the repository root:

```sh
QX_BUILD_PARALLEL=2 bash tools/QuickExcel.JscProbe/build-macos.sh
bash QuickExcel.ScriptWorker/native/build-macos.sh
dotnet publish QuickExcel.Desktop.Avalonia/QuickExcel.Desktop.Avalonia.csproj -c Release -r osx-arm64 --self-contained true -p:QuickExcelNativeRid=osx-arm64 -o artifacts/publish-mac
```

macOS prerequisites: Xcode Command Line Tools (Apple clang), CMake 4 + Ninja,
Python 3.9+ (the preparation scripts avoid 3.11-only APIs), system Ruby and curl.
The first run downloads the pinned WebKit/ICU/TypeScript archives. `macos-sources.json`
adds 104 unchanged WebKit files omitted from the GTK archive; the WebKitFS and
PlatformJSCOnly adjustments are recorded in `source-adjustments.json`. WebKit's
Apple port links the OS `libicucore`, so the pinned ICU build is not shipped on
macOS. The worker sandbox (Seatbelt, parent-death and polling resource watchdogs)
is built by `QuickExcel.ScriptWorker/native/build-macos.sh`. Building these
artifacts is not signing, notarization or device acceptance.

All downloads, sources, builds, logs and uniquely named results stay in ignored
`artifacts/jsc/`. `sources.json` pins archive URLs/hashes and the WebKit commit.
The GTK release omits 27 Windows WTF files: `windows-sources.json` pins those
unmodified files from the same commit. Preparation verifies every cached source
against the archives and the four exact `source-adjustments.json` changes:
two missing includes, unconditional rejection of import attributes, and respect
for an explicit `ENABLE_API_TESTS=OFF` on Linux. Arbitrary source edits fail verification.

JIT/DFG and bundled mimalloc 3.2.8 are enabled; FTL, WebAssembly, remote inspector
and upstream API tests are disabled. Native identity is
`JavaScriptCore/2.54.2/b95c91b1/qx-sync-2`; IPC is 8. Build identity files
record compiler/options/source/binary hashes. JSC is the sole production engine.
J5 Windows and Linux isolation/recovery evidence is recorded independently.
Production build defaults to the JavaScriptCore target; use `-IncludeProbe` only
when running the historical J0 experiments below.

`prepare_compiler.py` creates the trusted official TS 5.9.3 compiler, ES2022
standard libraries, `QuickExcel.Scripting/TypeScript/frontend.js` and product
`quickexcel.d.ts` bundle. The worker verifies its SHA-256 against the host's
embedded identity. Product compilation performs full strict check/emit with an
in-memory host. It preserves ESM and never executes workbook code, resolves OS
files, external packages or configuration, or supplies DOM/Node APIs.

J5 real-worker measurements superseded the J0 probe's 768 MiB compiler budget:
the fixed compiler process budget is now 1 GiB, retaining its five/six-second
deadlines. The 768 MiB failed samples remain in the implementation log.

Execution keeps a 256 MiB process cap, 256 KiB JS stack, one-second native wall
deadline and five-second round. Compilation has its own 1 GiB process cap,
5 MiB JS stack, five-second native deadline/six-second host kill, concurrency 1
and queue 0. Both roles use the same fixed JSC library and OS access isolation.
**JSC-MEM-001 remains open:** a 64 MiB total engine allocation hard cap is not
implemented; the user explicitly deferred it and removed it as a progression gate.
GC thresholds and process limits do not satisfy that separate cap.

Historical J0 experiments remain independently runnable:

```powershell
.\tools\QuickExcel.JscProbe\build.ps1 -IncludeProbe
python .\tools\QuickExcel.JscProbe\run.py
python .\tools\QuickExcel.JscProbe\run.py --performance --repeats 3
```

They intentionally record C API microtask/CPU-watchdog counterexamples; those are
not the product bridge. `bridge_checks.py` checks the actual bridge and extracted
product bootstrap. Managed tests and packaged self-checks additionally exercise
real AppContainer workers. Performance runs use sequential fresh processes;
three samples do not establish p95. Keep failed evidence and compare identical builds.

`prepare_distribution.py` creates notices, pinned corresponding source archives,
Windows supplemental sources, bridge/build adjustments and dynamic replacement
instructions under `artifacts/jsc/distribution`. Release scripts require these
materials, required JSC/ICU/runtime/compiler files and reject probes, fault
fixtures and legacy engine payloads in all releases. JSC/WTF includes LGPL/BSD
and other notices, ICU uses Unicode terms, mimalloc is MIT and TS Apache-2.0;
the old QuickJS MIT notice cannot cover this distribution.
