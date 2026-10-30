# AGENTS.md

## Repository scope

This is the independent QuickExcel repository. Keep desktop, Core, infrastructure adapters, scripting, tests, documentation, and generated artifacts here. The sibling `../ExcelIO` is a separate repository containing the ExcelIO libraries. It is a source dependency, not part of this repository. Do not edit, stage, commit, or copy sibling library sources as part of ordinary QuickExcel changes. Library changes require a separately scoped task.

## Build and verification

Run from this repository root on Windows x64 with .NET 10 SDK:

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c Release --artifacts-path .\artifacts\build
.\scripts\build-quickexcel.ps1
```

Always supply `--artifacts-path` for managed builds, tests, runs, and publishes so referenced ExcelIO projects also write their outputs inside QuickExcel. Native runtime build requirements and source provenance are in `QuickExcel.Scripting/native/SOURCES.md`.

Important! DO NOT run the full regression test suite for every feature change or bug fix! Only run the tests related to the modified areas!

## Structure

- `QuickExcel.Core`: pure .NET domain model, document session, commands, history, facade and ports; no ExcelIO, WPF or native runtime references.
- `QuickExcel.Infrastructure`: typed XLSX/XLSQ v2 storage, retained legacy v1 regression adapter, typed workbook calculator with sparse dependencies and incremental evaluation and the QuickJS runtime factory. Current file support and limits are in `QuickExcel.Desktop/docs/09-m2-files-and-types.md`.
- `QuickExcel.Core.Test`: domain/session, typed/legacy storage and native scripting integration tests. File-format fixtures are constructed inline, including malformed ZIP/XML/JSON; keep prior v1 regressions.
- `QuickExcel.Scripting`: managed QuickJS bridge, native sources, and runtime smoke checks.
- `QuickExcel.Desktop`: WPF application and product/architecture documentation.
- `scripts/build-quickexcel.ps1`: test and publish the desktop app; `-RunSmoke -CreateArchive` validates the fresh package and creates a uniquely named ZIP plus SHA-256. `-OutputDirectory` must remain under this repository artifacts directory.
- `Start-QuickExcel.cmd`: launch the published app or build/run from source.

Keep `artifacts`, managed build caches, and native downloaded sources/build output out of Git. Preserve the native source version, SHA-256 verification, and license.

M3 syntax/binding/structure commands live in Core/Formulas and StructureCommands; WorkbookCalculator and TypedEvaluator live in Infrastructure/Calculation. LegacyWorkbookCalculator remains only for historical regression. Current calculation/bridge/export limits are documented in QuickExcel.Desktop/docs/10-m3-calculation.md.

M4 worker contracts live in QuickExcel.ScriptContracts (BCL only), the executable in QuickExcel.ScriptWorker, and the host client/Job in Infrastructure/Scripting. Desktop uses ScriptWorkerRuntimeFactory and ScriptWorkerActionExecutor without in-process fallback. scripts/ScriptWorker.targets copies the complete worker to build/publish output. QuickExcel.Core.Test/WorkerFixture is fault injection only and must never enter a release. Worker lifecycle evidence is in QuickExcel.Desktop/docs/11-m4-worker.md.

M4 package modules use file scriptApiVersion 2 and IPC protocol 4. Imports resolve declared logical IDs only through the native in-memory ES module loader; preserve API v1 script semantics. Module contracts are in QuickExcel.Desktop/docs/12-m4-modules.md. Actions require script-actions and workbook/actions bindings, execute against immutable snapshots, and commit validated patches through DocumentSession as one undoable transaction. Action contracts are in QuickExcel.Desktop/docs/13-m4-actions.md. Ordered UDF batching and the 16 MiB IPC/native bridge budget are in docs/14-m4-batching.md. WorkerProcess/WorkerSandbox are the mandatory zero-capability AppContainer launch path for formulas and actions: suspended creation, token verification, explicit IPC handles, Job at process creation, private read-only payload, environment allowlist, and profile cleanup; no same-permission fallback. The development-only tools/QuickExcel.SandboxProbe (including its read-only administrator WFP collector) must never enter a release. Normal-user Windows execution is needed for profile integration tests; do not weaken isolation to accommodate restricted test runners. Local loopback/DNS evidence and remaining deployment/security work are in docs/15-m4-os-isolation-adr.md. Strict policy and the final M6 OS security gate remain pending.

WorkerSandboxLease owns durable per-user records and exclusive leases for crash recovery; never sweep unjournaled legacy temp/profile names by prefix, follow reparse points, or remove a live lease/Job. QuickExcel.Core.Test/SandboxHostFixture is test-only and must not be published. Recovery contracts and local deployment evidence (including the direct long-path EXE launch failure) are in docs/16-m4-recovery-and-deployment.md. scripts/test-deployment-matrix.ps1 mutates only unique installation copies under artifacts and preserves failed-case evidence; local read-only ACL tests are not actual Program Files or clean-machine validation.

M5 grid/input contracts and evidence are in QuickExcel.Desktop/docs/17-m5-grid-and-input.md. IGridSurface/FastGridSurface are view-only (main-grid migration and current evidence: QuickExcel.Desktop/docs/24-main-grid-migration.md); writes still go through EditorController and DocumentSession. Desktop/Interaction contains WPF-free rules linked into Core.Test. Preserve explicit formula-paste policy, atomic paste/clear history, composition gating, and original-cell draft ownership. scripts/test-m5-grid.ps1 runs interaction/benchmark modes against repository artifacts; run benchmarks separately from tests, retain raw samples, and distinguish synthetic WPF composition from real OS IME verification. M5 and M6 quality gates are not closed by the first UI slice.

UI-004/UI-005 contracts and evidence are in docs/18-m5-structure-and-workspace.md under QuickExcel.Desktop. Core/PresentationCommands owns stable rectangular copy-semantics sorting, style patches, dimensions and sheet deletion. Filter state remains session-only and keyed by SheetId; workbook views, freeze/zoom and sheet visibility now persist under docs/33-xlsx-workbook-views.md (navigation does not mark dirty, view commands do); filtered paste must never silently map across hidden rows. Desktop saves use IVersionedWorkbookStorage/SaveChecked with an opened/saved content fingerprint. Recent-file settings and injected dialog choices must use isolated artifact paths in smoke tests. The Workspace mode in scripts/test-m5-grid.ps1 verifies real WPF handlers; do not claim physical input or complete M6 concurrency/OS/performance gates from it.
