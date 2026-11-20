# Third-party components and licenses

QuickExcel includes the following third-party materials. Each license applies to its respective component.

| Component | License | Location | Use in QuickExcel |
| --- | --- | --- | --- |
| RichTextKit | Apache-2.0 | `third_party/RichTextKit` | Text layout, line breaking, and HarfBuzz integration. Original copyright headers and license retained. |
| Office Scripts API definitions | CC BY 4.0 | `third_party/office-scripts-docs` | Input for generating script type declarations. QuickExcel implements the behavior independently. |
| Avalonia, Avalonia WebView, BuildServices, MicroCom, Tmds.DBus, Fluent System Icons | MIT | Licenses | Desktop interface, platform integration, and icons. |
| Markdown.Avalonia.Tight, ColorDocument.Avalonia, ColorTextBlock.Avalonia | MIT | Licenses/MARKDOWN-AVALONIA-LICENSE.txt | Native Markdown rendering for bundled license documents. |
| FastWpfGrid | MIT | licenses/QuickExcel.Grid.Shared | Grid algorithms and their source attribution. |
| ANGLE | BSD and accompanying notices | Licenses/ANGLE-LICENSE.txt | Graphics backend. |
| SkiaSharp and HarfBuzzSharp | MIT and the individual native library licenses | SKIA and HARFBUZZ license and third-party notice files | Graphics and fonts. |
| LibTiff.NET and PdfPig | BSD-3-Clause and Apache-2.0, respectively | LIBTIFF-LICENSE.txt and PDFPIG-LICENSE.txt | Images and PDF. |
| AngleSharp and OpenMcdf | MIT and MPL-2.0, respectively | Licenses/ANGLESHARP-LICENSE.txt and Licenses/OPENMCDF-LICENSE.txt | HTML and compound document readers. |
| Monaco Editor, PDF.js, esbuild | MIT, Apache-2.0, and MIT, respectively | Web/licenses | Editor, PDF preview, and build tooling. Fonts, character maps, and WebAssembly components include their own licenses. |
| JavaScriptCore / WebKit, ICU, mimalloc, TypeScript | Individual component licenses | JSC-Licenses | Copyright and license notices and rebuild instructions are included. ClickOnce packages provide the exact corresponding source, including QuickExcel modifications and build scripts, through the address in JSC-SOURCE.txt; other packages include the source archive. Microsoft terms apply to the Windows VC++ runtime. The macOS arm64 package links the OS `libicucore` under Apple's operating-system terms. |
| .NET runtime, including Native AOT runtime support | MIT and accompanying third-party terms | DOTNET-LICENSE and DOTNET-THIRD-PARTY-NOTICES in the self-contained distribution | Windows 4.0 uses .NET 10.0.15: a partially trimmed desktop runtime and Native AOT isolated worker. Other platforms retain their separately recorded runtime builds. |
| .NET ClickOnce launcher (Microsoft.DotNet.Mage 10.0.0) | MIT and accompanying third-party terms | Licenses/CLICKONCE-LICENSE.txt and CLICKONCE-THIRD-PARTY-NOTICES.txt | Windows ClickOnce entry point; generated launcher resources identify QuickExcel. |
| Microsoft Edge WebView2 Runtime; Visual Studio ClickOnce setup engine | Microsoft redistribution terms | Licenses/WINDOWS-INSTALLER-NOTICES.txt | The Windows installer includes the original Microsoft-signed WebView2 Evergreen online bootstrapper and the SDK-generated setup bootstrapper. |

## Windows ClickOnce distribution

The ClickOnce launcher comes from Microsoft.DotNet.Mage 10.0.0, upstream
`dotnet/deployment-tools` revision `c7242c203f7d302a24714497964510cdccccac3d`.
QuickExcel configures its launch target and signs the resulting launcher; upstream
copyright and MIT notices are retained. The Visual Studio bootstrapper engine is
redistributed under the installed Visual Studio SDK's redistribution terms.
WebView2 is a separate Microsoft component, not covered by QuickExcel's or
Avalonia's licenses. Its original installer is unmodified. See
`Licenses/WINDOWS-INSTALLER-NOTICES.txt` for official terms and sources.

## Markdown.Avalonia

Copyright (c) 2010 Bevan Arps, 2020 Whistyun. QuickExcel uses the unmodified
[Markdown.Avalonia](https://github.com/whistyun/Markdown.Avalonia) Tight package
12.0.2 and its ColorDocument.Avalonia (Copyright (c) 2024 whistyun) and
ColorTextBlock.Avalonia (Copyright (c) 2020 whistyun) dependencies under the MIT
license, from revision `f3f65f3f82b44f569958f0116c188e7a8e60686f`.
The complete upstream license is distributed as
`Licenses/MARKDOWN-AVALONIA-LICENSE.txt`. Markdown is rendered as native Avalonia
controls; the optional HTML, SVG and syntax-highlighting packages are not included.

## Fluent UI System Icons

Copyright (c) 2020 Microsoft Corporation. The Ribbon uses 94 selected 20px Regular
SVGs from revision `64e5aa757d1b70c24a283256ddcb8269d0f58773` of
[Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons), under
the MIT license (`Licenses/FLUENT-SYSTEM-ICONS-LICENSE.txt` in the distribution).
The original SVGs are unmodified; QuickExcel converts their paths and fill rules
to native Avalonia geometry. Original paths and SHA-256 hashes are recorded in
`third_party/fluentui-system-icons/manifest.json`. The expanded selection covers
formatting, data tools, tables, printing, protection and automation commands.

## Office Scripts API definitions

Copyright Microsoft Corporation and contributors. The source is the [Microsoft Office Scripts API reference repository](https://github.com/OfficeDev/office-scripts-docs-reference). The pinned commit and checksums are recorded in `third_party/office-scripts-docs/source.lock.json`.

QuickExcel modifies the declarations by renaming the namespace, removing documentation comments and the OfficeScript namespace, and adding QuickExcel extensions. Distributions containing the declarations or their derivatives must retain the copyright notice, source attribution, [CC BY 4.0 license reference](https://creativecommons.org/licenses/by/4.0/legalcode), and modification notice.

See [the Office Scripts notice](third_party/office-scripts-docs/NOTICE.md), also included as `Licenses/OFFICE-SCRIPTS-NOTICE.md` in the application package.

## RichTextKit

The [RichTextKit source record](third_party/RichTextKit/SOURCES.md) identifies the upstream revision, archive SHA-256, and QuickExcel modifications.
