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

    private void Build(TextLayoutRequest request, SkiaFonts fonts, Action<string, string> diagnostic)
    {
        var style = request.Style;
        var text = request.Text;
        var decorations = new Dictionary<RichStyle, (string? Underline, bool Strike)>();
        TextBlock Make(string value, float factor, bool constrain)
        {
            var block = new TextBlock { FontMapper = fonts, CharacterMatcher = Fallback, EllipsisEnabled = false,
                MaxWidth = constrain && style.WrapText && style.TextRotation == 0 && double.IsFinite(request.Width) ? Math.Max(1, (float)request.Width - Indent) : null,
                MaxLines = null };
            var rich = request.RichText?.Text == value ? request.RichText : null;
            foreach (var run in rich?.Runs.AsEnumerable() ?? [new TextRun(value)])
            {
                var f = run.Font;
                var fontRequest = new FontRequest(f?.ResolvedName ?? style.FontName ?? "Microsoft YaHei UI", f?.EastAsianName ?? style.EastAsianFontName,
                    f?.Size ?? style.FontSize ?? request.DefaultFontSize * 72 / 96, f?.Bold ?? style.Bold, f?.Italic ?? style.Italic);
                // RichTextKit shapes tabs as missing glyphs. Expand only the display
                // projection, per run, so measuring/wrapping and drawing agree while
                // the workbook text and inherited rich-run fonts stay untouched.
                var segments = StringInfo.GetTextElementEnumerator(run.Text.Replace("\t", "    "));
                SKTypeface? previous = null; var buffer = new StringBuilder();
                void Flush()
                {
                    if (previous is null || buffer.Length == 0) return;
                    var rs = new RichStyle { FontFamily = fonts.Key(previous), FontSize = (float)(fontRequest.Points * 96 / 72) * factor,
                        FontWeight = fontRequest.Bold ? 700 : 400, FontItalic = fontRequest.Italic,
                        TextColor = SkiaRenderer.Color(f?.Color ?? style.FontColor, SKColors.Black),
                        FontVariant = f?.Vertical switch { "superscript" => FontVariant.SuperScript, "subscript" => FontVariant.SubScript, _ => FontVariant.Normal } };
                    decorations[rs] = (f?.Underline ?? style.Underline, f?.Strike ?? style.Strike);
                    // RichTextKit 0.6.170's UTF-16 decoder ORs the surrogate offset instead of adding it.
                    // Supply scalar values directly so U+20000 and higher keep their real identity.
                    block.AddText(new Topten.RichTextKit.Utils.Slice<int>(buffer.ToString().EnumerateRunes().Select(r => r.Value).ToArray()), rs); buffer.Clear();
                }
                while (segments.MoveNext())
                {
                    var element = segments.GetTextElement();
                    var face = fonts.Resolve(element, fontRequest, diagnostic);
                    if (face != previous) { Flush(); previous = face; }
                    buffer.Append(element);
                }
                Flush();
            }
            block.Layout(); return block;
        }
        if (style.Indent > 0) { var space = Make(" ", 1, false); Indent = space.FontRuns.Sum(r => r.Width) * 3 * style.Indent; space.Clear(); }
        var layout = Make(text, 1, false);
        if (request is { FillOffset: >= 0, FillCharacter: { } fill } && double.IsFinite(request.Width))
        {
            var unit = Make(fill.ToString(), 1, false);
            var width = unit.FontRuns.Sum(r => r.Width); unit.Clear();
            var count = width > 0 ? Math.Clamp((int)Math.Floor((request.Width - Indent - layout.MeasuredWidth) / width), 0, 4096) : 0;
            if (request.FillOffset <= text.Length) text = text.Insert(request.FillOffset, new string(fill, count));
            layout.Clear(); layout = Make(text, 1, false);
        }
        var factor = style.ShrinkToFit && !style.WrapText && double.IsFinite(request.Width) && layout.MeasuredWidth > request.Width - Indent
            ? Math.Max(.02f, (float)(request.Width - Indent) / Math.Max(1, layout.MeasuredWidth)) : 1;
        if (factor != 1 || style.WrapText) { layout.Clear(); layout = Make(text, factor, true); }
        try
        {
            var width = layout.Lines.Select(l => l.Runs.Sum(r => r.Width)).DefaultIfEmpty().Max();
            var height = layout.MeasuredHeight;
            var angle = style.TextRotation is > 0 and <= 180 ? style.TextRotation <= 90 ? -style.TextRotation : style.TextRotation - 90 : 0;
            var radians = angle * Math.PI / 180;
            Metrics = new(width, height, Math.Abs(width * Math.Cos(radians)) + Math.Abs(height * Math.Sin(radians)),
                Math.Abs(width * Math.Sin(radians)) + Math.Abs(height * Math.Cos(radians)), layout.LineCount);
            FontIdentities = layout.FontRuns.Select(r => $"{r.Typeface.FamilyName}; weight={fonts.Weight(r.Typeface)}; italic={r.Typeface.IsItalic}").Distinct().ToArray();
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
                var decoration = decorations[(RichStyle)run.Style];
                // RichTextKit positions lines and bidi runs. Apply horizontal cell alignment per line.
                var dx = double.IsFinite(request.Width) ? style.Horizontal switch
                { 2 => ((float)request.Width - Indent - line.Width) / 2, 3 => (float)request.Width - Indent - line.Width, _ => 0 } : 0;
                if (style.TextRotation != 0) dx = 0;
                _runs.Add(new(blob, run.Style.TextColor, dy, run.XCoord + dx, run.XCoord + run.Width + dx,
                    line.YCoord + line.BaseLine, run.Style.FontSize, decoration.Underline, decoration.Strike));
                _offsets.Add(dx);
            }
        }
        finally { layout.Clear(); }
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
