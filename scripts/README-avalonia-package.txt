QuickExcel Avalonia — Windows x64 (version and identity: release.json)

Run quickxl.exe. The partially trimmed .NET 10 desktop runtime and the isolated
Native AOT script worker are included. JSC/ICU/TypeScript live in ScriptWorker only.
Monaco and PDF.js are bundled locally; Node and an internet connection are not needed.

Microsoft Edge WebView2 Evergreen Runtime is required for script editing and PDF preview.
Install the current x64 Evergreen Runtime from Microsoft's official page:
https://developer.microsoft.com/microsoft-edge/webview2/
For offline devices, download the x64 Evergreen Standalone Installer on a connected
machine, verify its Microsoft Authenticode signature and transfer that installer.
QuickExcel does not silently download/install software or use Edge preview channels.
Existing file/grid work remains available if the WebView Runtime is missing.

Self-check (writes only new diagnostic files; use a new empty destination):
  quickxl.exe --self-check C:\Users\YOUR_NAME\Desktop\QuickExcel-check
Custom isolated settings:
  quickxl.exe --user-data C:\Users\YOUR_NAME\QuickExcel-settings [workbook.xlsq]

Disable all workbook scripts and execution/compiler workers:
  quickxl.exe --no-scripts [workbook.xlsq]
Or set QUICKEXCEL_NO_SCRIPTS=1 in the process, Windows user or machine environment.
The Scripts & API settings page also has a Disable scripts checkbox (immediate).
Any active source overrides signatures and trust; the status bar lists every source.
Command-line/environment values are read at startup. Clearing the checkbox cannot
override them and never silently reauthorizes the current workbook.
The installation self-check fails if scripts are disabled.

On first normal start, compatible WPF UI/recent-file preferences are copied into
LocalAppData\QuickExcel\Avalonia, with an immutable backup and versioned journal.
Existing Avalonia preferences are retained. Certificates/trust pins and browser
profiles are never imported. The original WPF settings and workbooks remain intact.

Keep older tested packages for rollback. Close QuickExcel before switching the
launcher to an older package; do not overwrite workbooks or copy new trust state
into the older application. release.json and the ZIP SHA-256 identify this package.

Physical IME/DPI/accessibility/printing and clean-device acceptance are deferred.
Linux/macOS support and existing M4/M6 security gates are tracked separately.
Licenses are in this directory, licenses (including grid provenance) and Web/licenses.

JavaScriptCore is the only script engine. JS API v1 remains compatible; API v2
modules can use JavaScript or TypeScript. TypeScript is fully checked by the
isolated compiler before Apply/Run; failed checks keep the draft and previous
applied source. Files save the original source. Older packages reject the
required typescript-sources feature: keep an original file backup before migration.
Compilation uses a separate 1 GiB / 5-second native / 6-second host budget;
execution retains 256 MiB process / 256 KiB JS stack / 1-second operation / 5-second round.
OPEN JSC-MEM-001: the 64 MiB total engine allocation hard limit is not implemented.
It is explicitly deferred and is not a progression gate. A process limit or GC
threshold is not an equivalent engine limit. JSC-Licenses contains the upstream
notices, corresponding sources and instructions for rebuilding/replacing libraries.
Rollback replaces the whole host/worker package; never mix engine or IPC versions.
