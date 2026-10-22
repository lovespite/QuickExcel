# AGENTS.md

## Repository scope

This is the independent QuickExcel repository. Keep desktop, Core, scripting, tests, documentation, and generated artifacts here. The sibling `../ExcelIO` is a separate repository containing the ExcelIO libraries. It is a source dependency, not part of this repository. Do not edit, stage, commit, or copy sibling library sources as part of ordinary QuickExcel changes. Library changes require a separately scoped task.

## Build and verification

Run from this repository root on Windows x64 with .NET 10 SDK:

```powershell
dotnet build .\QuickExcel.slnx -c Release --artifacts-path .\artifacts\build -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -c Release --artifacts-path .\artifacts\build -v minimal
dotnet run --project .\QuickExcel.Scripting\smoke\QuickExcel.Scripting.Smoke.csproj -c Release --artifacts-path .\artifacts\build
.\scripts\build-quickexcel.ps1
```

Always supply `--artifacts-path` for managed builds, tests, runs, and publishes so referenced ExcelIO projects also write their outputs inside QuickExcel. Native runtime build requirements and source provenance are in `QuickExcel.Scripting/native/SOURCES.md`.

## Structure

- `QuickExcel.Core`: workbook document operations and XLSQ packaging.
- `QuickExcel.Core.Test`: document and native scripting integration tests.
- `QuickExcel.Scripting`: managed QuickJS bridge, native sources, and runtime smoke checks.
- `QuickExcel.Desktop`: WPF application and product/architecture documentation.
- `scripts/build-quickexcel.ps1`: test and publish the desktop app.
- `Start-QuickExcel.cmd`: launch the published app or build/run from source.

Keep `artifacts`, managed build caches, and native downloaded sources/build output out of Git. Preserve the native source version, SHA-256 verification, and license.
