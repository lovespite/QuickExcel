using System.Globalization;
using System.Text;
using QuickExcel.Domain;
using SkiaSharp;
using Topten.RichTextKit;
using RichStyle = Topten.RichTextKit.Style;

namespace QuickExcel.Rendering.Skia;

/// <summary>Owned shaped glyph runs retain UTF-8 text and clusters for PDF Unicode mapping.</summary>
public sealed class SkiaTextLayout : IDisposable
{
    private sealed record Run(SKTextBlob Blob, SKColor Color, float OffsetY, float Left, float Right, float Baseline,
        float Size, string? Underline, bool Strike);
    private sealed class CellRunStyle : RichStyle
    {
        public string? CellUnderline { get; init; }
        public bool CellStrike { get; init; }
    }
    private readonly List<Run> _runs = [];
    public TextMetrics Metrics { get; private set; } = new(0, 0, 0, 0, 0);
    public IReadOnlyList<string> FontIdentities { get; private set; } = [];
    public float Indent { get; private set; }
    private static readonly object LayoutGate = new();
    private sealed class NoImplicitFallback : ICharacterMatcher
    {
        public SKTypeface MatchCharacter(string familyName, int weight, int width, SKFontStyleSlant slant, string[] bcp47, int character) => null!;
    }
    private static readonly ICharacterMatcher Fallback = new NoImplicitFallback();
    internal static void ReleaseFonts(IEnumerable<SKTypeface> faces)
    {
        lock (LayoutGate)
        {
            foreach (var face in faces) { TextBlock.ReleaseTypeface(face); face.Dispose(); }
            TextBlock.ResetPooledMemory();
        }
    }
    internal static void ReleaseShapingCaches(IEnumerable<SKTypeface> faces)
    {
        lock (LayoutGate)
        {
            // The report still owns the exact font faces. Only recreatable HarfBuzz
            // font tables and temporary layout pools are released between operations.
            foreach (var face in faces) TextBlock.ReleaseTypeface(face);
            TextBlock.ResetPooledMemory();
        }
    }

    internal static SkiaTextLayout Create(TextLayoutRequest request, SkiaFonts fonts, Action<string, string> diagnostic)
    {
        // RichTextKit uses shared shape/run pools. Serialize layout; retained blobs are independent and immutable.
        lock (LayoutGate)
        {
            var result = new SkiaTextLayout();
            try { result.Build(request, fonts, diagnostic); return result; }
            catch { result.Dispose(); throw; }
        }
    }

    private static void LayoutParagraphs(TextLayoutRequest request, SkiaFonts fonts, Action<string, string> diagnostic,
        float factor, float? maxWidth, Action<TextBlock, bool> visit)
    {
        var style = request.Style;
        var block = new TextBlock { FontMapper = fonts, CharacterMatcher = Fallback, EllipsisEnabled = false, MaxWidth = maxWidth };
        var rich = request.RichText?.Text == request.Text ? request.RichText : null;
        RichStyle? breakStyle = null;
        var previousWasCr = false;
        void FinishParagraph()
        {
            var empty = block.Length == 0;
            // A hard break reserves a line even at the end of a cell. Use the
            // break's font for its height; the measuring space is never drawn.
            if (empty && breakStyle is not null) block.AddText(" ", breakStyle);
            try { block.Layout(); visit(block, empty); }
            finally { block.Clear(); }
        }
        foreach (var run in rich?.Runs.AsEnumerable() ?? [new TextRun(request.Text)])
        {
            var f = run.Font;
            var fontRequest = new FontRequest(f?.ResolvedName ?? style.FontName ?? "Microsoft YaHei UI", f?.EastAsianName ?? style.EastAsianFontName,
                f?.Size ?? style.FontSize ?? request.DefaultFontSize * 72 / 96, f?.Bold ?? style.Bold, f?.Italic ?? style.Italic);
            var segments = StringInfo.GetTextElementEnumerator(run.Text.Replace("\t", "    "));
            SKTypeface? previous = null;
            var buffer = new StringBuilder();
            RichStyle StyleFor(SKTypeface face)
            {
                return new CellRunStyle { FontFamily = fonts.Key(face), FontSize = (float)(fontRequest.Points * 96 / 72) * factor,
                    FontWeight = fontRequest.Bold ? 700 : 400, FontItalic = fontRequest.Italic,
                    CellUnderline = f?.Underline ?? style.Underline, CellStrike = f?.Strike ?? style.Strike,
                    TextColor = SkiaRenderer.Color(f?.Color ?? style.FontColor, SKColors.Black),
                    FontVariant = f?.Vertical switch { "superscript" => FontVariant.SuperScript, "subscript" => FontVariant.SubScript, _ => FontVariant.Normal } };
            }
            void Flush()
            {
                if (previous is null || buffer.Length == 0) return;
                var scalars = buffer.ToString().EnumerateRunes().Select(r => r.Value).ToArray();
                block.AddText(new Topten.RichTextKit.Utils.Slice<int>(scalars), StyleFor(previous));
                buffer.Clear();
            }
            while (segments.MoveNext())
            {
                var element = segments.GetTextElement();
                // CRLF is one grapheme within a run but can straddle rich-text runs.
                var followsCr = previousWasCr;
                previousWasCr = element == "\r";
                if (followsCr && element == "\n") continue;
                if (element is "\n" or "\r" or "\r\n" or "\v" or "\f" or "\u0085" or "\u2028" or "\u2029")
                {
                    Flush();
                    breakStyle = StyleFor(fonts.Resolve(" ", fontRequest, diagnostic));
                    FinishParagraph();
                    continue;
                }
                // Interlinear annotation delimiters also have no visible glyph.
                // Leave joiners, bidi controls and selectors in the shaping input.
                if (char.IsControl(element[0]) || element[0] is >= '\uFFF9' and <= '\uFFFB') continue;
                var face = fonts.Resolve(element, fontRequest, diagnostic);
                if (face != previous) { Flush(); previous = face; }
                buffer.Append(element);
            }
            Flush();
        }
        FinishParagraph();
    }

    private void Build(TextLayoutRequest request, SkiaFonts fonts, Action<string, string> diagnostic)
    {
        var style = request.Style;
        var text = request.Text;
        float NaturalWidth(string value)
        {
            var width = 0f;
            LayoutParagraphs(request with { Text = value }, fonts, diagnostic, 1, null, (block, empty) =>
            {
                if (!empty) width = Math.Max(width, block.FontRuns.Sum(r => r.Width));
            });
            return width;
        }
        if (style.Indent > 0) Indent = NaturalWidth(" ") * 3 * style.Indent;
        if (request is { FillOffset: >= 0, FillCharacter: { } fill } && double.IsFinite(request.Width))
        {
            var width = NaturalWidth(fill.ToString());
            var count = width > 0 ? Math.Clamp((int)Math.Floor((request.Width - Indent - NaturalWidth(text)) / width), 0, 4096) : 0;
            if (request.FillOffset <= text.Length) text = text.Insert(request.FillOffset, new string(fill, count));
        }
        var factor = 1f;
        if (style.ShrinkToFit && !style.WrapText && double.IsFinite(request.Width))
        {
            var natural = NaturalWidth(text);
            if (natural > request.Width - Indent) factor = Math.Max(.02f, (float)(request.Width - Indent) / Math.Max(1, natural));
        }
        float maxWidth = 0, height = 0;
        var lineCount = 0;
        var identities = new HashSet<string>();
        float? wrapWidth = style.WrapText && style.TextRotation == 0 && double.IsFinite(request.Width)
            ? Math.Max(1, (float)request.Width - Indent) : null;
        LayoutParagraphs(request with { Text = text }, fonts, diagnostic, factor, wrapWidth, (layout, empty) =>
        {
            var top = height;
            height += layout.MeasuredHeight;
            lineCount += layout.LineCount;
            if (empty) return;
            maxWidth = Math.Max(maxWidth, layout.Lines.Max(l => l.Runs.Sum(r => r.Width)));
            foreach (var r in layout.FontRuns) identities.Add($"{r.Typeface.FamilyName}; weight={fonts.Weight(r.Typeface)}; italic={r.Typeface.IsItalic}");
            foreach (var line in layout.Lines)
            foreach (var run in line.Runs)
            {
                if (run.Glyphs.Length == 0 || run.CodePoints.Length == 0) continue;
                var codepoints = run.CodePoints.ToArray();
                var value = string.Concat(codepoints.Select(char.ConvertFromUtf32));
                var bytes = Encoding.UTF8.GetBytes(value);
                var offsets = new int[codepoints.Length + 1];
                for (var i = 0; i < codepoints.Length; i++) offsets[i + 1] = offsets[i] + new Rune(codepoints[i]).Utf8SequenceLength;
                var variant = run.Style.FontVariant;
                var scale = variant == FontVariant.Normal ? 1 : .65f;
                var dy = variant switch { FontVariant.SuperScript => -run.Style.FontSize * .35f, FontVariant.SubScript => run.Style.FontSize * .1f, _ => 0 };
                using var font = new SKFont(run.Typeface, run.Style.FontSize * scale) { Edging = SKFontEdging.Antialias, Subpixel = true, Hinting = SKFontHinting.None };
                using var builder = new SKTextBlobBuilder();
                var b = builder.AllocateRawPositionedTextRun(font, run.Glyphs.Length, bytes.Length);
                run.Glyphs.AsSpan().CopyTo(b.Glyphs);
                run.GlyphPositions.AsSpan().CopyTo(b.Positions);
                bytes.CopyTo(b.Text);
                for (var i = 0; i < run.Clusters.Length; i++) b.Clusters[i] = (uint)offsets[Math.Clamp(run.Clusters[i] - run.Start, 0, codepoints.Length - 1)];
                var blob = builder.Build();
                if (blob is null) continue;
                var decoration = (CellRunStyle)run.Style;
                // RichTextKit positions lines and bidi runs. Apply horizontal cell alignment per line.
                var dx = double.IsFinite(request.Width) ? style.Horizontal switch
                { 2 => ((float)request.Width - Indent - line.Width) / 2, 3 => (float)request.Width - Indent - line.Width, _ => 0 } : 0;
                if (style.TextRotation != 0) dx = 0;
                _runs.Add(new(blob, run.Style.TextColor, top + dy, run.XCoord + dx, run.XCoord + run.Width + dx,
                    line.YCoord + line.BaseLine, run.Style.FontSize, decoration.CellUnderline, decoration.CellStrike));
                _offsets.Add(dx);
            }
        });
        var angle = style.TextRotation is > 0 and <= 180 ? style.TextRotation <= 90 ? -style.TextRotation : style.TextRotation - 90 : 0;
        var radians = angle * Math.PI / 180;
        Metrics = new(maxWidth, height, Math.Abs(maxWidth * Math.Cos(radians)) + Math.Abs(height * Math.Sin(radians)),
            Math.Abs(maxWidth * Math.Sin(radians)) + Math.Abs(height * Math.Cos(radians)), lineCount);
        FontIdentities = identities.ToArray();
    }
    private readonly List<float> _offsets = [];
    public void Draw(SKCanvas canvas, float x, float y, bool monochrome = false)
    {
        using var paint = new SKPaint { IsAntialias = true };
        for (var i = 0; i < _runs.Count; i++)
        {
            var run = _runs[i]; paint.Color = monochrome ? SKColors.Black : run.Color;
            canvas.DrawText(run.Blob, x + _offsets[i], y + run.OffsetY, paint);
            paint.StrokeWidth = Math.Max(.7f, run.Size / 16);
            var baseline = y + run.Baseline + run.OffsetY;
            if (run.Underline is not (null or "none"))
            {
                canvas.DrawLine(x + run.Left, baseline + 2, x + run.Right, baseline + 2, paint);
                if (run.Underline is "double" or "doubleAccounting") canvas.DrawLine(x + run.Left, baseline + 4, x + run.Right, baseline + 4, paint);
            }
            if (run.Strike) canvas.DrawLine(x + run.Left, baseline - run.Size * .3f, x + run.Right, baseline - run.Size * .3f, paint);
        }
    }
    public void Dispose() { foreach (var run in _runs) run.Blob.Dispose(); _runs.Clear(); }
}
