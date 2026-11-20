# Pinned JavaScriptCore bridge

The product embeds source-built WebKitGTK 2.54.2 / JavaScriptCore commit
`b95c91b1c0ee43cc86d25ff7af73038f67a88e04`, ICU 77.1 and bundled mimalloc 3.2.8.
Exact archive and supplemental Windows/macOS source hashes, six explicit source
adjustments, tools and options are under `tools/QuickExcel.JscProbe/`.
See its [build README](../../../tools/QuickExcel.JscProbe/README.md) and
[implementation evidence](../../../QuickExcel.Desktop/docs/71-jsc-typescript-implementation-log.md).

macOS arm64 is built with `tools/QuickExcel.JscProbe/build-macos.sh` (Apple
clang, CMake/Ninja, Python 3.9+, system Ruby). The GTK archive omits macOS WTF
files, so `macos-sources.json` pins 104 unchanged upstream files with SHA-256;
`WebKitFS.cmake` and `PlatformJSCOnly.cmake` carry the documented framework
header and duplicate-timezone build fixes. WebKit's Apple port links the OS
`libicucore` (macOS 26.6: ICU 78.1) rather than the pinned ICU build. The
macOS worker enforces the 256 MiB process bound with a polling footprint/thread
watchdog, not a cgroup-equivalent hard limit; signing/notarization and
signed-bundle persistent trust remain unverified.

`bridge.cpp` is QuickExcel-owned code compiled into the JSC shared library.
Managed callers see only `qx_create`, `qx_create_compiler`, `qx_load_modules`,
`qx_call`, `qx_set_host_call`, `qx_set_builtin_call`, `qx_destroy`, `qx_free_string`, `qx_was_interrupted`, `qx_is_usable`
and `qx_engine_identity`. No JSC C++ type crosses this ABI. UTF-8 buffers are
copied across it and released by the same allocator; native/IPC output is capped
at 16 MiB and diagnostic text at 16 KiB.

The live model bridge identity is `JavaScriptCore/2.54.2/b95c91b1/qx-libraries-1`
(IPC 13; previous asynchronous bridge was `qx-async-3`). `qx_set_host_call` accepts a bounded synchronous UTF-8 callback; its
reply is owned by the managed caller until the native operation returns, never
freed by the engine allocator. The worker forwards only the validated model
protocol through its inherited IPC pipes, exposing no CLR objects or OS paths.
Registration roles come from native module source stack frames, with URL overrides
disabled; scripts cannot grant themselves another entry's binding by replacing
global properties. The private role selector can also inspect declared entry roles:
network grants may be requested only from the initialization root, before loading
finishes, never from an imported helper or another role. Source pins, hashes and
licenses are unchanged.

Private version adaptation points are `JSModuleLoader::makeModule`, module
record `loadedModules`/link/evaluateSync, `GlobalObjectMethodTable`, `Strong`,
`VM::DrainMicrotaskDelayScope`, microtask queue inspection/clear/action checkpoint,
`VM::ensureTerminationException`/`notifyNeedTermination`, native exception
inspection and raw stack-frame formatting. All are confined to this bridge.
Upgrading WebKit requires rerunning native and isolated-worker contracts.

Modules resolve declared logical IDs entirely from immutable in-memory source.
QuickExcel 7.0 also resolves nine exact built-in IDs through `qx_set_builtin_call`,
registered by the worker before module loading. The callback returns verified
embedded sources, never files or network responses. User IDs take precedence;
NUL, relative paths, URL specifiers and unknown package IDs cannot select a builtin.
Native builtin sources have a separate limit of 16 modules / 4 MiB, with capacity
reserved before graph pointers are retained. Buffers are copied before the next
callback and freed by the managed allocator after loading. Builtins have no entry
roles and consume no user module/source quota. Code, type and adapter hashes bind
the worker handshake, compiler cache and consent policy. Pins and modification
notices are in `tools/QuickExcel.ScriptLibraries/NOTICE.md`; compatibility and
verification are in `QuickExcel.Desktop/docs/86-builtin-script-libraries.md`.
With file API v3 / IPC 10, managed and IPC validation still cap raw user module
sources at 1 MiB. The native executable-source aggregate permits an additional
32 KiB for host-generated `main` registration trailers (at most 64 declared
action entries). UDF contexts receive no action trailers. This preserves the
full raw-source allowance without increasing the user-input budget; engine,
process, stack and time limits remain unchanged. The exact 1 MiB action case
is covered by `ScriptP1Tests.MaximumRawActionSourceBudgetIncludesRoomForHostRegistration`.

The bridge consumes an already settled internal parse Promise, links records
directly, then synchronously evaluates entries in order. Loading never drains user
microtasks. Records stay rooted before recursive linking, preserving cycles,
live bindings and TDZ. Dynamic import, import attributes, deferred imports,
top-level await, unknown IDs and external paths are rejected. API v2 pending
jobs outside explicit action scheduler turns poison the VM; API v1 preserves its
prior inert-job behavior. Actions and the v4 initialization main may return a
Promise. Only private `runAsyncAction` and `settleFetch` calls advance microtasks,
under the same native watchdog. Events, console and UDFs retain their synchronous
contracts. Dynamic import remains forbidden in every role.

Fetch queues bounded requests inside an Action and returns a Promise. Between
native turns, managed code sends the request through the existing authenticated
model IPC to the host broker. The worker remains zero-capability and never opens
network sockets. HTTP waiting consumes the configured host Action round budget but not
the one-second native JavaScript turn budget. Up to sixteen requests are dispatched
serially; the returned action Promise must settle before its transaction commits.
The Fetch objects are QuickExcel-owned code with a host-supplied body bound
(1–10 MiB, default 8 MiB; IPC 13). Runtime response retention is independently
bounded to the greater of 8 MiB and twice that body bound, plus 64 bodies, with
explicit eviction of old bodies. Panel retention is a separate host setting.
Private fetch IPC bodies use padded base64url to avoid JSON escaping expansion;
the existing 16 MiB frame and cumulative model round limits are unchanged.

Each operation arms one steady-clock deadline across load/call/serialization/error
formatting. A separate timer thread requests an uncatchable JSC termination trap;
mutex/generation coordination prevents a late trap from affecting another call.
The termination exception is preallocated. Timeouts, stack/OOM and forbidden
async work permanently invalidate the VM; the worker abandons the whole batch
and the host kills the round without replay. Error inspection does not invoke
user message/stack/toString getters. Script text cannot forge the native invalid flag.

Execution has 1 second per native operation/async continuation turn, 256 KiB JS stack and independent
5-second UDF round / 256 MiB process enforcement. Actions use a separately configured
host round deadline (default 10 seconds); this does not enlarge the native turn or
UDF limits, or override a platform's independent OS lifetime. Windows API v2 retains
one worker and module context per authorized workbook between rounds; idle time
does not consume a round deadline. Linux retains its five-second OS process
lifetime and one-shot workers. Compiler processes lock their role
on first creation and use the same library with 5-second operations, 5 MiB JS
stack, 1 GiB process and 6-second host deadline. Only the hash-verified compiler
bundle is evaluated in that role; user input is data. OS isolation is mandatory
for both roles; development probes alone do not establish it.

The compiler requests GC after 32 MiB allocated per cycle and collects transient
managed bootstrap buffers before native checking. This is a performance/memory
pressure setting, not a total live-heap or engine allocation cap. Dense 1 MiB
sources are separately tested in the real isolated compiler; arbitrary inputs
can still fail the fixed process budget and never yield partial accepted output.

**Open JSC-MEM-001:** the separate 64 MiB total engine allocation hard limit is
not implemented. The user deferred it on 2026-11-03 and explicitly allowed later
stages to proceed. GC thresholds and 256 MiB process enforcement are not an
equivalent implementation. IPC policy names and release manifests disclose this.

JSC/WTF contains LGPL/BSD and other third-party terms. Original copyright and
license headers remain in the source archives. Release packages carry
`JSC-Licenses/` with notices, exact corresponding sources, bridge/build scripts
and rebuilding/dynamic replacement instructions. ClickOnce packages replace the
source archive with one public exact-source address in `JSC-SOURCE.txt`, retaining
the notices and rebuild instructions; the worker itself carries only
its required runtime files. Microsoft VC runtimes retain Microsoft terms.
Historical QuickJS sources/notices and complete rollback packages are preserved.
