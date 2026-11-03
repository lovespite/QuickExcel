# Pinned JavaScriptCore bridge

The product embeds source-built WebKitGTK 2.54.2 / JavaScriptCore commit
`b95c91b1c0ee43cc86d25ff7af73038f67a88e04`, ICU 77.1 and bundled mimalloc 3.2.8.
Exact archive and supplemental Windows source hashes, four explicit source
adjustments, tools and options are under `tools/QuickExcel.JscProbe/`.
See its [build README](../../../tools/QuickExcel.JscProbe/README.md) and
[implementation evidence](../../../QuickExcel.Desktop/docs/71-jsc-typescript-implementation-log.md).

`bridge.cpp` is QuickExcel-owned code compiled into the JSC shared library.
Managed callers see only `qx_create`, `qx_create_compiler`, `qx_load_modules`,
`qx_call`, `qx_destroy`, `qx_free_string`, `qx_was_interrupted`, `qx_is_usable`
and `qx_engine_identity`. No JSC C++ type crosses this ABI. UTF-8 buffers are
copied across it and released by the same allocator; native/IPC output is capped
at 16 MiB and diagnostic text at 16 KiB.

Private version adaptation points are `JSModuleLoader::makeModule`, module
record `loadedModules`/link/evaluateSync, `GlobalObjectMethodTable`, `Strong`,
`VM::DrainMicrotaskDelayScope`, microtask queue inspection/clear,
`VM::ensureTerminationException`/`notifyNeedTermination`, native exception
inspection and raw stack-frame formatting. All are confined to this bridge.
Upgrading WebKit requires rerunning native and isolated-worker contracts.

Modules resolve declared logical IDs entirely from immutable in-memory source.
The bridge consumes an already settled internal parse Promise, links records
directly, then synchronously evaluates entries in order. It never drains user
microtasks. Records stay rooted before recursive linking, preserving cycles,
live bindings and TDZ. Dynamic import, import attributes, deferred imports,
top-level await, unknown IDs and external paths are rejected. API v2 pending
jobs poison the VM; API v1 preserves its prior inert-job behavior.

Each operation arms one steady-clock deadline across load/call/serialization/error
formatting. A separate timer thread requests an uncatchable JSC termination trap;
mutex/generation coordination prevents a late trap from affecting another call.
The termination exception is preallocated. Timeouts, stack/OOM and forbidden
async work permanently invalidate the VM; the worker abandons the whole batch
and the host kills the round without replay. Error inspection does not invoke
user message/stack/toString getters. Script text cannot forge the native invalid flag.

Execution has 1 second per native operation, 256 KiB JS stack and independent
5-second host / 256 MiB process enforcement. Compiler processes lock their role
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
and rebuilding/dynamic replacement instructions; the worker itself carries only
its required runtime files. Microsoft VC runtimes retain Microsoft terms.
Historical QuickJS sources/notices and complete rollback packages are preserved.
