# A0 dependency provenance

The exact NuGet graph/content hashes are in `packages.lock.json`; npm archive URLs
and integrity digests are in [package-lock.json](../../QuickExcel.Web/package-lock.json).
`scripts/export-avalonia-dependencies.ps1` verifies cached archive hashes, NuGet
normalized content identities and package signatures, and exports the full graph.
`npm ci --ignore-scripts` checks the npm lock identities. No XPF package, commercial
control, trial key or license service is used by this probe.

| Added dependency | Fixed source / version | License evidence |
| --- | --- | --- |
| Avalonia framework, Desktop, Fluent, Skia, platform backends | 12.1.4, [commit 2d79ba5](https://github.com/AvaloniaUI/Avalonia/tree/2d79ba55fae1b1e71c69f93a668080c9029e6a31) | `Licenses/AVALONIA-LICENSE.txt` from this commit |
| Avalonia.Controls.WebView | 12.1.1, [commit 43d26d9](https://github.com/AvaloniaUI/Avalonia.Controls.WebView/tree/43d26d90312b338d1d82bd12b215c16b5f3fa195) | MIT, `Licenses/AVALONIA-WEBVIEW-LICENSE.txt`; package records this commit but leaves repository URL empty |
| Avalonia.Angle.Windows.Natives | 2.1.27548.20260419, [commit 1c89805](https://github.com/AvaloniaUI/angle/tree/1c89805903c1482166356d3b950d474973180e61) | BSD notice from the exact package, `Licenses/ANGLE-LICENSE.txt` |
| Avalonia.BuildServices | 11.3.2, [commit 777f975](https://github.com/AvaloniaUI/Avalonia.BuildServices/tree/777f975b0a0cecf0311273711d56697212c558c0) | MIT, `Licenses/AVALONIA-BUILDSERVICES-LICENSE.txt`; build-time only |
| MicroCom.Runtime | 0.11.6, [commit 76785ef](https://github.com/kekekeks/MicroCom/tree/76785efcafd91b5902fd19dd11145f6dd655b7b4) | MIT, `Licenses/MICROCOM-LICENSE.txt` |
| Tmds.DBus.Protocol | 0.94.1, [commit b4a7fed](https://github.com/tmds/Tmds.DBus/tree/b4a7fed0b878f74cb54f7cca84d2889af4e596ba) | MIT, `Licenses/TMDS-DBUS-LICENSE.txt` |
| Microsoft.Web.WebView2 SDK | 1.0.4022.49, exact NuGet archive/hash in the lock and dependency inventory | BSD-style Microsoft notice from the exact package, `Licenses/WEBVIEW2-SDK-LICENSE.txt`; only the public Core COM projection is referenced, no WPF/WinForms controls or SDK build targets |
| PDF.js (`pdfjs-dist`) | 6.5.438, exact registry tarball/integrity in npm lock | Apache-2.0; build copies LICENSE, plus all applicable decoder/font notices with their assets |
| Monaco (`monaco-editor`) | 0.57.0, exact registry tarball/integrity in npm lock | MIT; build copies LICENSE and ThirdPartyNotices.txt |
| esbuild | 0.28.2, exact registry tarball/integrity in npm lock | MIT; build-time only, license copied for provenance |

The pre-existing SkiaSharp 4.154.0 / HarfBuzzSharp 14.2.1.400 / RichTextKit / PdfPig /
LibTiff notices remain in [Rendering.Skia/SOURCES.md](../../QuickExcel.Rendering.Skia/SOURCES.md).
Their versions were not downgraded for Avalonia. The transitive browser-only Skia
3.119.4 / HarfBuzz 8.3.1.3 WebAssembly packages appear in the restore graph but are
not loaded by the desktop probe. Runtime evidence identifies the actual desktop
native libraries. No browser .NET target is supported here.

Only the built browser resources are served. Node, npm optional native canvas,
esbuild binaries and all `node_modules` stay in artifacts and are not copied into
the web bundle. The manifest records each delivered resource's size and SHA-256.
Npm's optional packages for other operating systems remain listed honestly in the
dependency inventory, even if they were not installed on Windows.

The system WebView2 Runtime and WebKitGTK are separate installed components under
their own terms; they are not reclassified as MIT or bundled by this probe. The
tested Ubuntu/WebKit version, display mode and limitations are recorded in the A0
implementation record. Test fonts, QEMU/Ubuntu, GDB and independent PDF validators
remain test-only assets and are excluded from product and ScriptWorker releases.

The Windows virtual-host experiment uses Avalonia's public platform handle and the
SDK's public `ICoreWebView2_3` COM metadata. Those embedded TypeIdentifier types are
hidden from C# consumers, so the probe invokes their public methods through
reflection. It does not use CoreWebView2's private constructor, copy upstream
implementation code, or hard-code COM vtable offsets. See Microsoft's documented
[virtual-host API](https://learn.microsoft.com/en-us/microsoft-edge/webview2/reference/win32/icorewebview2_3)
and [local-content choices](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/working-with-local-content).
