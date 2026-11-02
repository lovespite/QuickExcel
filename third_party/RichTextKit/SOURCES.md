# RichTextKit layout source

Source: https://github.com/toptensoftware/RichTextKit/tree/6362fcdc33ea43b359ada2da946bf025a6d14630 (NuGet 0.6.170), Apache-2.0. Original copyright headers and LICENSE.txt are retained. The editor is not used or included.

Archive: https://github.com/toptensoftware/RichTextKit/archive/6362fcdc33ea43b359ada2da946bf025a6d14630.zip

SHA-256: `9D03F7AB4D66A6EB47FCAD20247E216418227984B901CDDF159C533C71B185E0`.

QuickExcel changes:

- A local net10.0 project pins SkiaSharp/HarfBuzz 4.154.0; no upstream build/signing tools or editor.
- Correct UTF-16 surrogate decoding: addition of 0x10000, not bitwise OR.
- Expose explicit typeface/shaper release; the upstream static dictionary otherwise retains every disposed report's native font and HarfBuzz resources indefinitely.
- Add per-TextBlock fallback matcher so immutable explicit-font fixtures never discover host fonts implicitly. QuickExcel resolves grapheme fonts before shaping and supplies a null matcher.
- Dispose temporary SKFonts in fallback enumeration.

QuickExcel uses this source for bidi, Unicode line breaking and HarfBuzz layout. It emits its own text blobs with UTF-8 and clusters for PDF ActualText, and owns their disposal. This is not a general RichTextKit editor integration.
