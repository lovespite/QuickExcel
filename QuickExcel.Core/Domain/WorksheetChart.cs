using System.Collections.Immutable;
using QuickExcel.Formulas;

namespace QuickExcel.Domain;

public enum WorksheetChartKind { Column, Line, Pie, Scatter, SmoothLine, Waterfall, Funnel }
public sealed record ChartStylePart(string RelationshipId, string Kind, ImmutableArray<byte> Bytes);
public sealed record ChartTextStyle(string Color = "FF404040", string? FontFamily = null, double Size = 9, bool Bold = false, bool Italic = false);
public sealed record ChartAppearance
{
    public string Background { get; init; } = "FFFFFFFF";
    public string PlotBackground { get; init; } = "00FFFFFF";
    public string BorderColor { get; init; } = "FFD3D3D3";
    public double BorderWidth { get; init; } = 1;
    public string GridColor { get; init; } = "FFD3D3D3";
    public ChartTextStyle Title { get; init; } = new(Size: 12);
    public ChartTextStyle Axis { get; init; } = new();
    public ChartTextStyle Legend { get; init; } = new();
}
public sealed record ChartSeries(string Name, string? NameReference, string Categories, string Values, string Color = "FF4472C4")
{
    // Empty entries in waterfall charts retain automatic increase/decrease/total colors.
    public ImmutableArray<string> PointColors { get; init; } = [];
    public ImmutableArray<int> TotalPoints { get; init; } = [];
    public bool NumericCategories { get; init; }
    public bool Smooth { get; init; }
    public bool ConnectPoints { get; init; } = true;
    public string Marker { get; init; } = "none";
    public double MarkerSize { get; init; } = 5;
    public double LineWidth { get; init; } = 2;
    public bool ShowValues { get; init; }
    public bool ShowCategories { get; init; }
    public bool ShowSeriesName { get; init; }
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
    public double? XMinimum { get; init; }
    public double? XMaximum { get; init; }
    public string XNumberFormat { get; init; } = "General";
    public int GapWidth { get; init; } = 150;
    public int Overlap { get; init; }
    public bool ShowConnectors { get; init; } = true;
    public string IncreaseColor { get; init; } = "FF4472C4";
    public string DecreaseColor { get; init; } = "FFED7D31";
    public string TotalColor { get; init; } = "FFA5A5A5";
    public ChartAppearance Appearance { get; init; } = new();
    // Validated, package-local style/color parts only; no external or executable payloads.
    public ImmutableArray<ChartStylePart> StyleParts { get; init; } = [];
    public ImmutableArray<byte> OriginalXml { get; init; } = [];
    public bool IsExtended => Kind is WorksheetChartKind.Waterfall or WorksheetChartKind.Funnel;
    public long EstimatedBytes => 1024L + OriginalXml.Length + StyleParts.Sum(p => p.Bytes.Length + 128L) + 2L * (Title.Length + NumberFormat.Length + XNumberFormat.Length) + Series.Sum(s => 384L + s.TotalPoints.Length * 4L + s.PointColors.Length * 48L + (s.Name.Length + (s.NameReference?.Length ?? 0) + s.Categories.Length + s.Values.Length) * 2L);
    public void Validate()
    {
        Anchor.Validate(); WorkbookLimits.ValidateText(Title); WorkbookLimits.ValidateText(NumberFormat);
        if (!Enum.IsDefined(Kind) || Series.IsDefaultOrEmpty || Series.Length > 16 || Kind is WorksheetChartKind.Pie or WorksheetChartKind.Waterfall or WorksheetChartKind.Funnel && Series.Length != 1 ||
            OriginalXml.IsDefault || OriginalXml.Length > 512 * 1024 || Legend is not (null or "l" or "r" or "t" or "b") ||
            Minimum is { } min && !double.IsFinite(min) || Maximum is { } max && !double.IsFinite(max) || Minimum >= Maximum ||
            XMinimum is { } xmin && !double.IsFinite(xmin) || XMaximum is { } xmax && !double.IsFinite(xmax) || XMinimum >= XMaximum ||
            GapWidth is < 0 or > 500 || Overlap is < -100 or > 100 || StyleParts.IsDefault || StyleParts.Length > 2 ||
            StyleParts.Select(p => p.Kind).Distinct().Count() != StyleParts.Length || StyleParts.Select(p => p.RelationshipId).Distinct().Count() != StyleParts.Length ||
            StyleParts.Any(p => p.Kind is not ("chartStyle" or "chartColorStyle") || string.IsNullOrWhiteSpace(p.RelationshipId) || p.Bytes.IsDefaultOrEmpty || p.Bytes.Length > 512 * 1024))
            throw new ArgumentException("图表类型、系列、图例、坐标轴或容量参数无效。");
        WorkbookLimits.ValidateText(XNumberFormat);
        if (IsExtended && (!VisibleOnly || Series.Any(s => s.NumericCategories))) throw new ArgumentException("瀑布图和漏斗图使用可见数据及分类标签。");
        static bool Color(string value) => value.Length == 8 && value.All(Uri.IsHexDigit);
        foreach (var color in new[] { IncreaseColor, DecreaseColor, TotalColor, Appearance.Background, Appearance.PlotBackground, Appearance.BorderColor, Appearance.GridColor })
            if (!Color(color)) throw new ArgumentException("图表颜色无效。");
        if (!double.IsFinite(Appearance.BorderWidth) || Appearance.BorderWidth is < 0 or > 100) throw new ArgumentException("图表边框宽度无效。");
        foreach (var text in new[] { Appearance.Title, Appearance.Axis, Appearance.Legend })
        {
            if (!Color(text.Color) || !double.IsFinite(text.Size) || text.Size is < 1 or > 400) throw new ArgumentException("图表文字样式无效。");
            WorkbookLimits.ValidateText(text.FontFamily ?? "");
        }
        foreach (var series in Series)
        {
            WorkbookLimits.ValidateText(series.Name);
            var categories = ParseReference(series.Categories); var values = ParseReference(series.Values);
            if (Count(categories) != Count(values) || Count(values) > 4096 || series.NameReference is { } name && Count(ParseReference(name)) != 1 ||
                !Color(series.Color) || series.PointColors.IsDefault || series.TotalPoints.IsDefault ||
                series.TotalPoints.Distinct().Count() != series.TotalPoints.Length || series.TotalPoints.Any(i => i < 0 || i >= Count(values)) ||
                series.Marker is not ("none" or "auto" or "circle" or "square" or "diamond" or "triangle" or "x" or "plus" or "star" or "dot" or "dash") ||
                !double.IsFinite(series.MarkerSize) || series.MarkerSize is < 2 or > 72 || !double.IsFinite(series.LineWidth) || series.LineWidth is < 0 or > 100 ||
                !series.PointColors.IsEmpty && (series.PointColors.Length != Count(values) || series.PointColors.Any(color => !(Kind == WorksheetChartKind.Waterfall && color.Length == 0) && !Color(color)))) throw new ArgumentException("图表系列必须包含等长的单行/单列引用，最多 4096 个点。");
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
        Minimum == other.Minimum && Maximum == other.Maximum && NumberFormat == other.NumberFormat && XMinimum == other.XMinimum && XMaximum == other.XMaximum && XNumberFormat == other.XNumberFormat && GapWidth == other.GapWidth && Overlap == other.Overlap &&
        ShowConnectors == other.ShowConnectors && IncreaseColor == other.IncreaseColor && DecreaseColor == other.DecreaseColor && TotalColor == other.TotalColor && Appearance == other.Appearance &&
        (OriginalXml.IsEmpty && StyleParts.IsEmpty || other.OriginalXml.IsEmpty && other.StyleParts.IsEmpty ||
            StyleParts.Length == other.StyleParts.Length && StyleParts.Zip(other.StyleParts).All(p => p.First.Kind == p.Second.Kind && p.First.Bytes.SequenceEqual(p.Second.Bytes))) && Series.Length == other.Series.Length && Series.Zip(other.Series).All(p =>
            p.First with { NameReference = null, Categories = "", Values = "", PointColors = [], TotalPoints = [] } == p.Second with { NameReference = null, Categories = "", Values = "", PointColors = [], TotalPoints = [] } &&
            p.First.PointColors.SequenceEqual(p.Second.PointColors) && p.First.TotalPoints.SequenceEqual(p.Second.TotalPoints) && Same(p.First.NameReference, p.Second.NameReference) && Same(p.First.Categories, p.Second.Categories) && Same(p.First.Values, p.Second.Values));
    private static bool Same(string? a, string? b) => a is null ? b is null : b is not null && FormulaSyntax.Print(FormulaSyntax.Parse("=" + a.TrimStart('='))) == FormulaSyntax.Print(FormulaSyntax.Parse("=" + b.TrimStart('=')));
}
