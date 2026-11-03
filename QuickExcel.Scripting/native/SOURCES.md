# Native runtime provenance

Current production engine: fixed JavaScriptCore / ICU / TypeScript, built with
`jsc/build.ps1` or `jsc/build-linux.sh`. Sources, hashes, licenses, resource
contracts and reproduction are in [jsc/SOURCES.md](jsc/SOURCES.md).
QuickJS production bridge and build targets were removed in J6. Rollback uses
a retained complete package. The following provenance is historical, including
retired commands; the original MIT notice remains in this repository.

## Historical QuickJS runtime provenance

The preceding QuickExcel package embedded **QuickJS-NG 0.17.0**, a maintained fork of Fabrice Bellard and
Charlie Gordon's QuickJS JavaScript engine. It does not use V8 or a managed JS
interpreter. The source is MIT licensed; the complete upstream license is in
`QUICKJS-LICENSE.txt` and remains in historical packages.

- Upstream: https://github.com/quickjs-ng/quickjs
- Pinned release: https://github.com/quickjs-ng/quickjs/releases/tag/v0.17.0
- Source archive: https://github.com/quickjs-ng/quickjs/archive/refs/tags/v0.17.0.zip
- Archive SHA-256: `9DE1BF573606D2AEBBBAB6217E616A66E09E3613518489B37B9B3157CF398DCC`

`build.ps1` verifies the archive hash before extracting and compiles the engine
from source using CMake and the installed Visual Studio C++ compiler. Source
downloads and generated binaries are ignored by Git. No downloaded prebuilt
executable or binary dependency is used. A normal project build runs the script
when the native DLL is missing or native bridge/build inputs are newer. To force a rebuild run:

```powershell
.\QuickExcel.Scripting\native\build.ps1 -Rebuild
```

Requirements to build from a fresh checkout: Windows x64, .NET 10 SDK, CMake
supporting your Visual Studio version, Visual Studio 2022 or newer with C++ build
tools, and internet access for the first source download. The published desktop
app includes the built DLL and does not require these build tools.

The bridge compiles only the core engine, without `quickjs-libc`, OS modules,
filesystem/network loaders, or CLR bindings. API v2 uses an in-memory ES module
loader that resolves only declared logical IDs, with no path or URL resolution.
Each session caps engine
allocation at 64 MiB and JS stack use at 256 KiB; each load and invocation has a
one second native interrupt deadline. Managed requests and native JSON results are
limited to 16 MiB UTF-8 (raised from 4 MiB in the M4 batching slice). The desktop runs this bridge in a separate disposable worker; native smoke
tests still run it in process. Process/Job containment is not a filesystem/network
security sandbox. Imported workbook scripts must still be explicitly enabled by the user.

Calls are serialized and `JS_UpdateStackTop` is called before execution, allowing
sequential access from different managed threads. Each document owns its own
runtime per calculation round. The prototype accepts synchronous scalar formula results and array
arguments; promises and object results are rejected.

M4 adds `qx_was_interrupted` so the worker can discard handles interrupted during
execution or exception stringification. Engine sources, version, archive hash and
license are unchanged. CMake uses the x64 host toolchain in `native/build/host-x64`.
The script gives child tools explicit Windows x64 platform variables, restores
Windows PowerShell's built-in module path, and removes inherited dotnet MSBuild
path overrides before invoking Visual Studio. Restricted build environments may
need access to Visual Studio FileTracker's system paths. Upstream compiler warnings
are preserved, not suppressed or patched out of vendored sources.

The 2026-10-28 module slice adds `qx_load_modules`. The complete reachable graph
shares one native one-second deadline; the host still bounds the entire round.
The engine owns static linking, cycles and once-only evaluation. Initialization
must complete synchronously; no Promise job loop is run. The module map is held
privately by the native session and released with it. The upstream version,
archive SHA-256 and license remain unchanged; no vendored engine source is edited.

The M5 product slice preserves bounded native Error stack diagnostics (logical module
IDs and source lines) through the existing worker failure message. The bridge captures
the engine's original stack getter before loading workbook code; it does not invoke
user-defined `stack` getters. Error text remains bounded to 16 KiB UTF-8 and worker
diagnostics retain the protocol log budget. Engine pin, source hash, license, memory,
time limits, module resolution and AppContainer launch policy remain unchanged.

Run the runtime smoke checks with:

```powershell
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj --artifacts-path .\artifacts\build
```


## Avalonia A6 platform builds

`build.sh` runs on Linux x64 or macOS arm64, verifies the same source ZIP/hash,
extracts into a fresh owned `artifacts/native/<rid>/build-*` directory and emits
`libquickexcel-js.so` / `.dylib` plus source/compiler/binary identity. Linux was
built with Clang 18 on Ubuntu 24.04.5; GCC 13 encountered an upstream optimization
compiler failure, retained in A6 evidence. No engine edits were made. macOS has a
build entry point but no executed device evidence. Unix uses `clock_gettime` and
PIC; the C ABI, logical module loader, limits and license remain unchanged.

Linux additionally builds repository-owned `QuickExcel.ScriptWorker/native/linux-sandbox.c`
with `CC=clang-18 bash QuickExcel.ScriptWorker/native/build-linux.sh`. This produces
the launcher and final TSYNC guard, not a third-party sandbox binary. The policy
uses kernel headers and libc, with no libseccomp runtime dependency. Identity
records contain policy source, compiler and both binary SHA-256 values.

Cross-publish must set both `RuntimeIdentifier` and global `QuickExcelNativeRid`:
RID-neutral referenced projects otherwise select the build host's native bridge.
`scripts/publish-avalonia-unix.ps1` enforces this, uses fresh build directories and
separate RID lock files under artifacts, rejects Windows/test/UI worker payloads,
and records native identities and licenses. Native outputs are never committed.
The ordinary production desktop never uses the in-process bridge as a fallback.

Linux Landlock/seccomp/cgroup requirements, negative probes, lifetime and resource
evidence are in `QuickExcel.Desktop/docs/67-avalonia-platforms.md`. macOS script
execution remains closed; a portable bridge or signed app is not a worker sandbox.
