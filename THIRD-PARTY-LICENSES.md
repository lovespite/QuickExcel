# Third-party components and licenses

QuickExcel includes the following third-party materials. Each license applies to its respective component.

| Component | License | Location | Use in QuickExcel |
| --- | --- | --- | --- |
| RichTextKit | Apache-2.0 | `third_party/RichTextKit` | Text layout, line breaking, and HarfBuzz integration. Original copyright headers and license retained. |
| Office Scripts API definitions | CC BY 4.0 | `third_party/office-scripts-docs` | Input for generating script type declarations. QuickExcel implements the behavior independently. |
| Avalonia, Avalonia WebView, BuildServices, MicroCom, Tmds.DBus, Fluent System Icons | MIT | Licenses | Desktop interface, platform integration, and icons. |
| FastWpfGrid | MIT | licenses/QuickExcel.Grid.Shared | Grid algorithms and their source attribution. |
| ANGLE | BSD and accompanying notices | Licenses/ANGLE-LICENSE.txt | Graphics backend. |
| SkiaSharp and HarfBuzzSharp | MIT and the individual native library licenses | SKIA and HARFBUZZ license and third-party notice files | Graphics and fonts. |
| LibTiff.NET and PdfPig | BSD-3-Clause and Apache-2.0, respectively | LIBTIFF-LICENSE.txt and PDFPIG-LICENSE.txt | Images and PDF. |
| AngleSharp and OpenMcdf | MIT and MPL-2.0, respectively | Licenses/ANGLESHARP-LICENSE.txt and Licenses/OPENMCDF-LICENSE.txt | HTML and compound document readers. |
| Monaco Editor, PDF.js, esbuild | MIT, Apache-2.0, and MIT, respectively | Web/licenses | Editor, PDF preview, and build tooling. Fonts, character maps, and WebAssembly components include their own licenses. |
| JavaScriptCore / WebKit, ICU, mimalloc, TypeScript | Individual component licenses | JSC-Licenses | The scripting package includes copyright and license notices, corresponding source, and rebuild instructions. Microsoft terms apply to the Windows VC++ runtime. |
| .NET runtime | MIT and accompanying third-party terms | DOTNET-LICENSE and DOTNET-THIRD-PARTY-NOTICES in the self-contained distribution | Desktop application and isolated processes. |

## Office Scripts API definitions

Copyright Microsoft Corporation and contributors. The source is the [Microsoft Office Scripts API reference repository](https://github.com/OfficeDev/office-scripts-docs-reference). The pinned commit and checksums are recorded in `third_party/office-scripts-docs/source.lock.json`.

QuickExcel modifies the declarations by renaming the namespace, removing documentation comments and the OfficeScript namespace, and adding QuickExcel extensions. Distributions containing the declarations or their derivatives must retain the copyright notice, source attribution, [CC BY 4.0 license reference](https://creativecommons.org/licenses/by/4.0/legalcode), and modification notice.

See [the Office Scripts notice](third_party/office-scripts-docs/NOTICE.md), also included as `Licenses/OFFICE-SCRIPTS-NOTICE.md` in the application package.

## RichTextKit

The [RichTextKit source record](third_party/RichTextKit/SOURCES.md) identifies the upstream revision, archive SHA-256, and QuickExcel modifications.
