using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Immutable UTF-16 selection edits. Insertions inherit the font to their left, replacements the first replaced character.</summary>
public static class RichTextEditing
{
    public static long EstimatedBytes(RichCellText text) => 128L + text.Text.Length * 4L + text.Runs.Sum(r => 256L +
        (r.Font is { } font ? 2L * ((font.Name?.Length ?? 0) + (font.ResolvedName?.Length ?? 0) + (font.EastAsianName?.Length ?? 0)) : 0));
    public static RichCellText Replace(RichCellText source, int start, int length, string text)
    {
        Validate(source, start, length); WorkbookLimits.ValidateText(text);
        var inherit = FontAt(source, length > 0 ? start : Math.Max(0, start - 1));
        return Normalize(Slice(source, 0, start).Concat([new TextRun(text, inherit)]).Concat(Slice(source, start + length, source.Text.Length - start - length)));
    }

    public static RichCellText Format(RichCellText source, int start, int length, TextRunFont patch, StyleDefinition style)
    {
        Validate(source, start, length);
        if (length == 0) throw new ArgumentException(MessageText.Get("请先选择要设置格式的文字。"));
        TextRun Apply(TextRun run)
        {
            var font = run.Font ?? new TextRunFont(style.FontName, style.FontSize, style.Bold, style.Italic, style.FontColor,
                Scheme: style.FontScheme, EastAsianName: style.EastAsianFontName, ResolvedName: style.FontName, Underline: style.Underline, Strike: style.Strike);
            return run with { Font = font with
            {
                Name = patch.Name ?? font.Name, ResolvedName = patch.Name ?? font.ResolvedName,
                EastAsianName = patch.Name is null ? font.EastAsianName : null, Scheme = patch.Name is null ? font.Scheme : null,
                Size = patch.Size ?? font.Size, Bold = patch.Bold ?? font.Bold, Italic = patch.Italic ?? font.Italic,
                Color = patch.Color ?? font.Color, RawRgb = patch.Color ?? font.RawRgb,
                ThemeColor = patch.Color is null ? font.ThemeColor : null, IndexedColor = patch.Color is null ? font.IndexedColor : null,
                AutoColor = patch.Color is null ? font.AutoColor : null, Tint = patch.Color is null ? font.Tint : null
            } };
        }
        return Normalize(Slice(source, 0, start).Concat(Slice(source, start, length).Select(Apply)).Concat(Slice(source, start + length, source.Text.Length - start - length)));
    }

    public static TextRunFont? FontAt(RichCellText source, int offset)
    {
        var current = 0;
        foreach (var run in source.Runs) { current += run.Text.Length; if (offset < current) return run.Font; }
        return source.Runs[^1].Font;
    }
    private static IEnumerable<TextRun> Slice(RichCellText source, int start, int length)
    {
        var offset = 0;
        foreach (var run in source.Runs)
        {
            var left = Math.Max(start, offset); var right = Math.Min(start + length, offset + run.Text.Length);
            if (right > left) yield return run with { Text = run.Text[(left - offset)..(right - offset)] };
            offset += run.Text.Length;
        }
    }
    private static RichCellText Normalize(IEnumerable<TextRun> input)
    {
        var runs = new List<TextRun>();
        var text = new System.Text.StringBuilder(); TextRunFont? font = null;
        foreach (var run in input)
        {
            if (run.Text.Length == 0) continue;
            if (text.Length > 0 && font != run.Font) { runs.Add(new(text.ToString(), font)); text.Clear(); }
            font = run.Font; text.Append(run.Text);
        }
        if (text.Length > 0) runs.Add(new(text.ToString(), font));
        return new(runs.Count == 0 ? [new TextRun("")] : runs);
    }
    private static void Validate(RichCellText source, int start, int length)
    {
        if (start < 0 || length < 0 || start > source.Text.Length - length) throw new ArgumentOutOfRangeException(nameof(start));
        foreach (var offset in new[] { start, start + length })
            if (offset > 0 && offset < source.Text.Length && char.IsHighSurrogate(source.Text[offset - 1]) && char.IsLowSurrogate(source.Text[offset]))
                throw new ArgumentException(MessageText.Get("选择范围不能拆开 Unicode 字符。"));
    }
}
