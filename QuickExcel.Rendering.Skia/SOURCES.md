# Shared rendering dependencies

Pinned production packages (NuGet content hashes in `packages.lock.json`):

| Component | Version / source | License and distribution |
| --- | --- | --- |
| SkiaSharp, HarfBuzz integration and Skia native assets | 4.154.0, mono/SkiaSharp commit `8771a4343c7ca2e5edf18f79a3b7b23c77fff60d` | MIT binding; complete native third-party notices copied from the exact native NuGet package |
| HarfBuzzSharp and native assets | 14.2.1.400, dependency selected by SkiaSharp.HarfBuzz 4.154.0 | MIT binding and native third-party notices from the exact package |
| BitMiracle.LibTiff.NET | 2.4.660, https://www.nuget.org/packages/BitMiracle.LibTiff.NET/2.4.660 | BSD-style license, IJG/LibTIFF/ComponentAce notices retained in `Licenses/LIBTIFF-LICENSE.txt`; original https://raw.githubusercontent.com/BitMiracle/libtiff.net/master/license.txt and https://bitmiracle.github.io/libtiff.net/help/articles/license.html |
| PdfPig | 0.1.16, UglyToad/PdfPig commit `a7bb35662bbbf405efddad50aedc9bcdcf515afc` | Apache-2.0 plus PDFBox/FontBox notices; complete LICENSE from that commit is shipped |
| RichTextKit | Locally adapted 0.6.170 sources | See [source and changes](../third_party/RichTextKit/SOURCES.md), Apache-2.0; license ships as RICH-TEXT-KIT-LICENSE.txt |

All `Licenses/*.txt` files are copied to build and publish output. Windows x64 publication includes the Windows native DLLs; Linux x64 smoke includes its ELF native libraries. macOS native assets are transitive package assets, **not evidence of a tested macOS target**. Core has no graphics dependency. Shared rendering has no WPF, Infrastructure, script worker or OS printing dependency. The grid's WPF adapter references the shared renderer; the worker does not.

Font assets downloaded by `scripts/get-rendering-test-fonts.ps1` are test-only OFL assets in repository artifacts. Product rendering discovers installed system fonts; replacements and missing scalars are recorded in bounded `RenderDiagnostic` collections with locations. Tests must not silently use system fonts in fixed-font comparisons.

MuPDF/PyMuPDF is an independent, test-only PDF parser/rasterizer used by the Python verification scripts. It is not referenced by a managed product project and is not redistributed in the desktop package. QEMU/Ubuntu and Python dependencies used for local Linux evidence likewise remain in ignored artifacts.

Upgrade these packages together with native loading, the fixed-font Windows/Linux smoke, image formats, actual PDF Unicode extraction and local image comparison. Installing a package or recognizing `%PDF-` is not a validation result.

`PdfOutputStream` is QuickExcel's bounded output adapter. It captures write/flush exceptions inside the native callback and rethrows on the managed export path. The pinned binding's [SKManagedWStream](https://github.com/mono/SkiaSharp/blob/8771a4343c7ca2e5edf18f79a3b7b23c77fff60d/binding/SkiaSharp/SKManagedWStream.cs) calls the underlying managed stream directly. Tests cover injected write/flush errors and actual Linux ENOSPC; this adapter does not copy upstream source.
