using System.Collections.Immutable;
using QuickExcel.Formulas;

namespace QuickExcel.Domain;

public enum WorksheetChartKind { Column, Line, Pie }
public sealed record ChartSeries(string Name, string? NameReference, string Categories, string Values, string Color = "FF4472C4")
{
    public ImmutableArray<string> PointColors { get; init; } = [];
}
public sealed record WorksheetChart
{
    public WorksheetChartKind Kind { get; init; }
    public string Title { get; init; } = "";
    public DrawingAnchor Anchor { get; init; } = new() { Kind = DrawingAnchorKind.OneCell, Width = 5_334_000, Height = 3_048_000 };
    public ImmutableArray<ChartSeries> Series { get; init; } = [];
    public string? Legend { get; init; } = "r";
    public bool VisibleOnly { get; init; } = true;
    public double? Minimum { get; init; }
    public double? Maximum { get; init; }
    public string NumberFormat { get; init; } = "General";
    // Validated dependency-free chart XML retains harmless Office extensions and presentation details.
    public ImmutableArray<byte> OriginalXml { get; init; } = [];
    public long EstimatedBytes => 512L + OriginalXml.Length + 2L * (Title.Length + NumberFormat.Length) + Series.Sum(s => 256L + s.PointColors.Length * 48L + (s.Name.Length + (s.NameReference?.Length ?? 0) + s.Categories.Length + s.Values.Length) * 2L);
    public void Validate()
    {
        Anchor.Validate(); WorkbookLimits.ValidateText(Title); WorkbookLimits.ValidateText(NumberFormat);
        if (!Enum.IsDefined(Kind) || Series.IsDefaultOrEmpty || Series.Length > 16 || Kind == WorksheetChartKind.Pie && Series.Length != 1 ||
            OriginalXml.IsDefault || OriginalXml.Length > 512 * 1024 || Legend is not (null or "l" or "r" or "t" or "b") ||
            Minimum is { } min && !double.IsFinite(min) || Maximum is { } max && !double.IsFinite(max) || Minimum >= Maximum)
            throw new ArgumentException("图表类型、系列、图例、坐标轴或容量参数无效。");
        foreach (var series in Series)
        {
            WorkbookLimits.ValidateText(series.Name);
            var categories = ParseReference(series.Categories); var values = ParseReference(series.Values);
            if (Count(categories) != Count(values) || Count(values) > 4096 || series.NameReference is { } name && Count(ParseReference(name)) != 1 ||
                series.Color.Length != 8 || series.Color.Any(c => !Uri.IsHexDigit(c)) || series.PointColors.IsDefault ||
                !series.PointColors.IsEmpty && (series.PointColors.Length != Count(values) || series.PointColors.Any(color => color.Length != 8 || color.Any(c => !Uri.IsHexDigit(c))))) throw new ArgumentException("图表系列必须包含等长的单行/单列引用，最多 4096 个点。");
        }
    }
    public static ReferenceNode ParseReference(string formula)
    {
        if (FormulaSyntax.Parse("=" + formula.TrimStart('=')) is not ReferenceNode r || r.Sheet is null || r.WholeColumns || r.Row != r.EndRow && r.Column != r.EndColumn)
            throw new NotSupportedException("图表数据必须直接引用本工作簿单张表中的单行或单列区域。");
        WorkbookLimits.ValidatePosition(r.Row, r.Column); WorkbookLimits.ValidatePosition(r.EndRow, r.EndColumn);
        if (r.EndRow < r.Row || r.EndColumn < r.Column) throw new ArgumentException("图表引用反向。");
        return r;
    }
    public static int Count(ReferenceNode r) => Math.Max(r.EndRow - r.Row + 1, r.EndColumn - r.Column + 1);
    public WorksheetChart MapReferences(Func<string, string> map) => this with { Series = Series.Select(s => s with { NameReference = s.NameReference is null ? null : map(s.NameReference), Categories = map(s.Categories), Values = map(s.Values) }).ToImmutableArray() };
    public bool ContentEquals(WorksheetChart other) => Kind == other.Kind && Title == other.Title && Anchor == other.Anchor && Legend == other.Legend && VisibleOnly == other.VisibleOnly &&
        Minimum == other.Minimum && Maximum == other.Maximum && NumberFormat == other.NumberFormat && Series.Length == other.Series.Length && Series.Zip(other.Series).All(p =>
            p.First.Name == p.Second.Name && p.First.Color == p.Second.Color && p.First.PointColors.SequenceEqual(p.Second.PointColors) && Same(p.First.NameReference, p.Second.NameReference) && Same(p.First.Categories, p.Second.Categories) && Same(p.First.Values, p.Second.Values));
    private static bool Same(string? a, string? b) => a is null ? b is null : b is not null && FormulaSyntax.Print(FormulaSyntax.Parse("=" + a.TrimStart('='))) == FormulaSyntax.Print(FormulaSyntax.Parse("=" + b.TrimStart('=')));
}
