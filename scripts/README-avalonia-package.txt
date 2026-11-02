QuickExcel Avalonia 0.6.0-preview.5 — Windows x64

Run QuickExcel.Avalonia.exe. .NET 10 and the isolated script worker are included.
Monaco and PDF.js are bundled locally; Node and an internet connection are not needed.

Microsoft Edge WebView2 Evergreen Runtime is required for script editing and PDF preview.
Install the current x64 Evergreen Runtime from Microsoft's official page:
https://developer.microsoft.com/microsoft-edge/webview2/
For offline devices, download the x64 Evergreen Standalone Installer on a connected
machine, verify its Microsoft Authenticode signature and transfer that installer.
QuickExcel does not silently download/install software or use Edge preview channels.
Existing file/grid work remains available if the WebView Runtime is missing.

Self-check (writes only new diagnostic files; use a new empty destination):
  QuickExcel.Avalonia.exe --self-check C:\Users\YOUR_NAME\Desktop\QuickExcel-check
Custom isolated settings:
  QuickExcel.Avalonia.exe --user-data C:\Users\YOUR_NAME\QuickExcel-settings [workbook.xlsq]

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
