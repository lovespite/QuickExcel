# A0 dependency and WebView probe

This is a disposable test host, not the Avalonia product application. It is deliberately
outside `QuickExcel.slnx`, and must never enter a desktop release or ScriptWorker payload.
The implementation record and acceptance status are in
[62-avalonia-a0-validation.md](../../QuickExcel.Desktop/docs/62-avalonia-a0-validation.md).

On Windows, from the repository root:

```powershell
.\scripts\get-rendering-test-fonts.ps1
.\scripts\test-avalonia-a0.ps1
# Windows virtual HTTPS/worker experiment using the documented WebView2 COM API.
.\scripts\test-avalonia-a0.ps1 -SkipWebBuild -VirtualHttps
# Starts a window for real desktop checks. Close it after completing the observations.
.\scripts\test-avalonia-a0.ps1 -SkipWebBuild -Interactive
```

The script requires .NET 10, Node/npm and the fixed private test fonts. Dependencies,
bundles, browser profiles and evidence stay in ignored `artifacts/avalonia`. The app
never opens a workbook or changes application settings. It captures native module
paths, exercises Avalonia and shared Skia in the same process, exports the existing
three-page rich-text/image/chart report fixture, runs real Monaco/PDF.js workers and
checks the local resource boundary. A successful JSON file alone is insufficient:
the **process exit code must also be zero**.

The default origin is the restricted loopback server shared with Linux. The
optional HTTPS experiment maps only an isolated directory of verified test assets
and a generated test PDF; it retains those files as evidence. It uses a CSP meta
element because folder mappings do not expose arbitrary response headers. This
experiment does not implement the A4 production session/PDF resource service.

Linux uses the same sources and web bundle:

```powershell
dotnet publish tools/QuickExcel.AvaloniaProbe/QuickExcel.AvaloniaProbe.csproj -c Release -r linux-x64 --self-contained true --artifacts-path artifacts/avalonia/build -o artifacts/avalonia/linux-publish -v minimal
```

Copy that output to `BUNDLE/bin`, `artifacts/avalonia/web` to `BUNDLE/web`, and the test
fonts to `BUNDLE/fonts`; preserve executable permissions. Run the Linux script in a
normal user's graphical session. `xvfb-run` is useful for automated compatibility
checks, but its results do not constitute desktop IME/DPI acceptance. See the record
for the tested WebKitGTK packages; no browser sandbox-disabling flags are used.

```bash
bash run-avalonia-a0-linux.sh /path/to/BUNDLE /path/to/new-results
```

Cross-RID publishing can regenerate a shared project's NuGet lock runtime sections.
Before returning to a locked Windows build, restore the normal graph (the fixed
package versions remain unchanged) and review the lock-file diff:

```powershell
dotnet restore tools/QuickExcel.AvaloniaProbe/QuickExcel.AvaloniaProbe.csproj --artifacts-path artifacts/avalonia/build -p:RestoreForceEvaluate=true
```

Independent PDF verification requires test-only PyMuPDF 1.26.7, Pillow 12.0.0 and
numpy 2.3.4 under an artifact directory. Do not ship these tools with the product:

```powershell
python scripts/verify-skia-output.py artifacts/avalonia/WINDOWS_RUN --dependencies artifacts/avalonia/python --report-only
python scripts/verify-skia-output.py artifacts/avalonia/LINUX_RUN --dependencies artifacts/avalonia/python --report-only --compare artifacts/avalonia/WINDOWS_RUN
```
