using System.Collections.Immutable;

namespace QuickExcel.Domain;

public sealed record TextRunFont(string? Name = null, double? Size = null, bool? Bold = null, bool? Italic = null,
    string? Color = null, int? ThemeColor = null, int? IndexedColor = null, bool? AutoColor = null, double? Tint = null,
    string? Scheme = null, string? EastAsianName = null, string? ResolvedName = null,
    string? Underline = null, bool? Strike = null, string? Vertical = null, int? Family = null, int? Charset = null, string? RawRgb = null);
public sealed record TextRun(string Text, TextRunFont? Font = null);

/// <summary>Run boundaries and explicit font overrides are input data, separate from a cell's base style.</summary>
public sealed class RichCellText : IEquatable<RichCellText>
{
    public RichCellText(IEnumerable<TextRun> runs)
    {
        ArgumentNullException.ThrowIfNull(runs);
        Runs = runs.Take(32768).ToImmutableArray();
        if (Runs.Length is < 1 or > 32767 || Runs.Any(r => r is null || r.Text is null)) throw new ArgumentException("富文本分段无效或超出容量。");
        if (Runs.Sum(r => (long)r.Text.Length) > WorkbookLimits.MaxCellTextLength) throw new ArgumentException("富文本总长度超限。");
        Runs = Runs.Select(r => r.Font is { } font ? r with { Font = font with { ResolvedName = font.ResolvedName ?? font.Name,
            RawRgb = font.RawRgb ?? (font.ThemeColor is null && font.IndexedColor is null && font.AutoColor is null ? font.Color : null) } } : r).ToImmutableArray();
        Text = string.Concat(Runs.Select(r => r.Text)); WorkbookLimits.ValidateText(Text);
        foreach (var font in Runs.Select(r => r.Font).OfType<TextRunFont>())
        {
            if (font.Size is { } size && (!double.IsFinite(size) || size is <= 0 or > 409) ||
                font.Tint is { } tint && (!double.IsFinite(tint) || tint is < -1 or > 1) ||
                font.ThemeColor is < 0 or > 11 || font.IndexedColor is < 0 or > 65 && font.IndexedColor != 81 || font.Family is < 0 or > 5 || font.Charset is < 0 or > 255 ||
                font.Scheme is not (null or "major" or "minor" or "none") ||
                font.Underline is not (null or "none" or "single" or "double" or "singleAccounting" or "doubleAccounting") ||
                font.Vertical is not (null or "baseline" or "superscript" or "subscript")) throw new ArgumentException("富文本字体属性无效。");
            if (font.Color is { } color && (color.Length != 8 || color.Any(c => !Uri.IsHexDigit(c)))) throw new ArgumentException("富文本颜色无效。");
            if (font.RawRgb is { } raw && (raw.Length != 8 || raw.Any(c => !Uri.IsHexDigit(c)))) throw new ArgumentException("富文本原始颜色无效。");
            foreach (var name in new[] { font.Name, font.EastAsianName, font.ResolvedName })
                if (name is not null && (string.IsNullOrWhiteSpace(name) || name.Length > 255)) throw new ArgumentException("字体名称无效。");
        }
    }
    public ImmutableArray<TextRun> Runs { get; }
    public string Text { get; }
    public bool Equals(RichCellText? other) => other is not null && Runs.SequenceEqual(other.Runs);
    public override bool Equals(object? obj) => obj is RichCellText other && Equals(other);
    public override int GetHashCode() { var hash = new HashCode(); foreach (var run in Runs) hash.Add(run); return hash.ToHashCode(); }
}

public sealed record CellHyperlink(string? Target, string? Location = null, string? Display = null, string? Tooltip = null);
public sealed record CellNote(string Author, RichCellText Text);

/// <summary>Imported cell attachments. VML is an opaque, validated note-only drawing with fixed coordinates.</summary>
public sealed class WorksheetAnnotations
{
    public WorksheetAnnotations(IEnumerable<KeyValuePair<CellPosition, CellHyperlink>>? links = null,
        IEnumerable<KeyValuePair<CellPosition, CellNote>>? notes = null, IEnumerable<byte>? noteDrawing = null)
    {
        Links = links?.Take(65531).ToImmutableDictionary() ?? ImmutableDictionary<CellPosition, CellHyperlink>.Empty;
        Notes = notes?.Take(65531).ToImmutableDictionary() ?? ImmutableDictionary<CellPosition, CellNote>.Empty;
        NoteDrawing = noteDrawing?.Take(1024 * 1024 + 1).ToImmutableArray() ?? [];
        if (Links.Count > 65530 || Notes.Count > 65530 || NoteDrawing.Length > 1024 * 1024) throw new ArgumentException("链接/批注容量超限。");
        foreach (var p in Links.Keys.Concat(Notes.Keys)) WorkbookLimits.ValidatePosition(p.Row, p.Column);
        foreach (var link in Links.Values)
        {
            if (link is null || string.IsNullOrEmpty(link.Target) && string.IsNullOrEmpty(link.Location)) throw new ArgumentException("链接缺少目标。");
            foreach (var text in new[] { link.Target, link.Location, link.Display, link.Tooltip })
                if (text is { Length: > 32767 } || text?.Any(c => c == '\0') == true) throw new ArgumentException("链接字段无效。");
        }
        foreach (var note in Notes.Values)
            if (note is null || note.Text is null || note.Author is null || note.Author.Length > 32767) throw new ArgumentException("批注字段无效。");
    }
    public ImmutableDictionary<CellPosition, CellHyperlink> Links { get; }
    public ImmutableDictionary<CellPosition, CellNote> Notes { get; }
    public ImmutableArray<byte> NoteDrawing { get; }
    public bool IsEmpty => Links.IsEmpty && Notes.IsEmpty && NoteDrawing.IsEmpty;
    public bool ContentEquals(WorksheetAnnotations other) => Links.Count == other.Links.Count && Links.All(p => other.Links.GetValueOrDefault(p.Key) == p.Value)
        && Notes.Count == other.Notes.Count && Notes.All(p => other.Notes.GetValueOrDefault(p.Key) == p.Value) && NoteDrawing.SequenceEqual(other.NoteDrawing);
}
