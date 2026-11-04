using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed class SetChartCommand(SheetId sheetId, int index, WorksheetChart? chart) : DocumentCommand
{
    public static WorksheetChart ChangeKind(WorksheetChart chart, WorksheetChartKind kind)
    {
        if (chart.Kind == kind) return chart;
        if (kind is WorksheetChartKind.Pie or WorksheetChartKind.Waterfall or WorksheetChartKind.Funnel && chart.Series.Length != 1)
            throw new ArgumentException("饼图、瀑布图和漏斗图需要单个系列，请选择两列数据范围。");
        string[] palette = ["FF4472C4", "FFED7D31", "FFA5A5A5", "FFFFC000", "FF5B9BD5", "FF70AD47"];
        var axes = kind is not (WorksheetChartKind.Pie or WorksheetChartKind.Funnel);
        return chart with
        {
            Kind = kind, OriginalXml = [], StyleParts = [], VisibleOnly = kind is WorksheetChartKind.Waterfall or WorksheetChartKind.Funnel || chart.VisibleOnly,
            Minimum = axes ? chart.Minimum : null, Maximum = axes ? chart.Maximum : null, NumberFormat = axes ? chart.NumberFormat : "General",
            XMinimum = null, XMaximum = null, XNumberFormat = "General", GapWidth = kind is WorksheetChartKind.Column or WorksheetChartKind.Waterfall ? chart.GapWidth : 150, Overlap = 0,
            ShowConnectors = true, IncreaseColor = "FF4472C4", DecreaseColor = "FFED7D31", TotalColor = "FFA5A5A5",
            Appearance = chart.Appearance with { Title = chart.Title.Length == 0 ? new(Size: 12) : chart.Appearance.Title,
                Axis = axes ? chart.Appearance.Axis : new(), GridColor = axes ? chart.Appearance.GridColor : "FFD3D3D3", Legend = chart.Legend is null ? new() : chart.Appearance.Legend },
            Series = [.. chart.Series.Select(s => s with { NumericCategories = kind == WorksheetChartKind.Scatter, Smooth = kind == WorksheetChartKind.SmoothLine,
                ConnectPoints = kind != WorksheetChartKind.Scatter, Marker = kind == WorksheetChartKind.Scatter ? "circle" : "none", MarkerSize = 5,
                LineWidth = 2, TotalPoints = [], ShowValues = kind is WorksheetChartKind.Waterfall or WorksheetChartKind.Funnel,
                PointColors = kind is WorksheetChartKind.Pie or WorksheetChartKind.Funnel ? [.. Enumerable.Range(0, WorksheetChart.Count(WorksheetChart.ParseReference(s.Values))).Select(i => palette[i % palette.Length])] : [] })]
        };
    }

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        if (index < 0 || index > sheet.Data.Charts.Length || chart is null && index == sheet.Data.Charts.Length) throw new ArgumentOutOfRangeException(nameof(index));
        chart?.Validate();
        if (chart is not null && chart.Anchor.Bounds(sheet) is { Width: > 0, Height: > 0 } bounds)
            _ = DrawingGeometry.FromBounds(sheet, bounds, DrawingAnchorKind.TwoCell);
        if (chart is not null) foreach (var reference in chart.Series.SelectMany(s => new[] { s.NameReference, s.Categories, s.Values }).OfType<string>())
            if (!snapshot.Sheets.Any(s => s.Name.Equals(WorksheetChart.ParseReference(reference).Sheet, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("图表引用的工作表不存在。");
        var charts = chart is null ? sheet.Data.Charts.RemoveAt(index) : index == sheet.Data.Charts.Length ? sheet.Data.Charts.Add(chart) : sheet.Data.Charts.SetItem(index, chart);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(s.Data with { Charts = charts }) : s)), [], true, CalculationUnchanged: true);
    }

    public static WorksheetChart FromRange(WorksheetSnapshot sheet, CellRectangle range, WorksheetChartKind kind, DrawingAnchor anchor, string title = "")
    {
        if (range.Rows < 2 || range.Columns < 2 || range.Columns > 17 || kind is WorksheetChartKind.Pie or WorksheetChartKind.Waterfall or WorksheetChartKind.Funnel && range.Columns != 2)
            throw new ArgumentException("首行为系列标题，首列为分类（散点图为 X 值）；需要 2–17 列，饼图、瀑布图和漏斗图需要 2 列。");
        string Ref(int top, int left, int bottom, int right) => FormulaSyntax.Print(new ReferenceNode(sheet.Name, top, left, bottom, right, true, true, true, true, IsRange: top != bottom || left != right));
        string[] palette = ["FF4472C4", "FFED7D31", "FFA5A5A5", "FFFFC000", "FF5B9BD5", "FF70AD47"];
        return new() { Kind = kind, Title = title, Anchor = anchor, Series = [.. Enumerable.Range(range.Left + 1, range.Columns - 1).Select((c, i) =>
            new ChartSeries("", Ref(range.Top, c, range.Top, c), Ref(range.Top + 1, range.Left, range.Bottom, range.Left), Ref(range.Top + 1, c, range.Bottom, c), palette[i % palette.Length])
            { NumericCategories = kind == WorksheetChartKind.Scatter, ConnectPoints = kind != WorksheetChartKind.Scatter,
                Smooth = kind == WorksheetChartKind.SmoothLine, Marker = kind == WorksheetChartKind.Scatter ? "circle" : "none",
                ShowValues = kind is WorksheetChartKind.Waterfall or WorksheetChartKind.Funnel,
                PointColors = kind is WorksheetChartKind.Pie or WorksheetChartKind.Funnel ? [.. Enumerable.Range(0, range.Rows - 1).Select(p => palette[p % palette.Length])] : [] })] };
    }
}

public sealed record ChartSeriesData(string Name, string[] Categories, double?[] Values, string Color, string[] PointColors)
{
    public int SeriesIndex { get; init; }
    public int[] SourceIndices { get; init; } = [];
    public double?[] XValues { get; init; } = [];
    public bool[] Totals { get; init; } = [];
}
public static class ChartData
{
    public static IReadOnlyList<ChartSeriesData> Read(WorkbookSnapshot book, WorksheetChart chart, bool includeFiltered = false)
    {
        CellSnapshot?[] Read(string formula, out WorksheetSnapshot source, out ReferenceNode reference)
        {
            reference = WorksheetChart.ParseReference(formula); var name = reference.Sheet;
            source = book.Sheets.FirstOrDefault(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) ?? throw new InvalidOperationException("图表源工作表不存在。");
            var cells = new CellSnapshot?[WorksheetChart.Count(reference)];
            for (var i = 0; i < cells.Length; i++)
            {
                var r = reference.Row + (reference.EndRow > reference.Row ? i : 0); var c = reference.Column + (reference.EndColumn > reference.Column ? i : 0);
                var cell = source.GetCell(r, c);
                if (cell?.Content.Kind == CellContentKind.Formula && cell.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error))
                    throw new InvalidOperationException("图表源公式尚未重算，旧图形已暂停显示。");
                cells[i] = cell;
            }
            return cells;
        }
        string Display(CellSnapshot? cell, WorksheetSnapshot source, int row, int col) => CellDisplayFormatter.Format(cell?.Value ?? CellValue.Blank, book.Styles[source.GetStyleId(row, col)], book.DateSystem);
        var output = new List<ChartSeriesData>();
        foreach (var (series, seriesIndex) in chart.Series.Select((s, i) => (s, i)))
        {
            if (series.IsFiltered && !includeFiltered) continue;
            var values = Read(series.Values, out var valueSheet, out var vr); var categories = Read(series.Categories, out var categorySheet, out var cr);
            var hidden = chart.VisibleOnly ? WorksheetFilterEngine.Evaluate(book, valueSheet.Id) : [];
            var categoryHidden = chart.VisibleOnly ? WorksheetFilterEngine.Evaluate(book, categorySheet.Id) : [];
            bool Visible(WorksheetSnapshot sheet, ReferenceNode r, int i, IReadOnlySet<int> excluded) =>
                !excluded.Contains(r.Row + (r.EndRow > r.Row ? i : 0)) && sheet.Rows.GetValueOrDefault(r.Row + (r.EndRow > r.Row ? i : 0))?.Hidden != true && sheet.Columns.GetValueOrDefault(r.Column + (r.EndColumn > r.Column ? i : 0))?.Hidden != true;
            var excluded = series.HiddenPoints.ToHashSet();
            var selected = Enumerable.Range(0, values.Length).Where(i => (includeFiltered || !excluded.Contains(i)) &&
                (!chart.VisibleOnly || Visible(valueSheet, vr, i, hidden) && Visible(categorySheet, cr, i, categoryHidden))).ToArray();
            var name = series.Name;
            if (series.NameReference is { } title) { var cell = Read(title, out var s, out var r)[0]; name = Display(cell, s, r.Row, r.Column); }
            output.Add(new(name, selected.Select(i => Display(categories[i], categorySheet, cr.Row + (cr.EndRow > cr.Row ? i : 0), cr.Column + (cr.EndColumn > cr.Column ? i : 0))).ToArray(),
                selected.Select(i => values[i]?.Value.Kind == CellValueKind.Number ? (double?)values[i]!.Value.NumberValue : null).ToArray(), series.Color,
                selected.Select(i => series.PointColors.IsEmpty ? series.Color : series.PointColors[i]).ToArray())
            {
                SeriesIndex = seriesIndex, SourceIndices = selected,
                XValues = selected.Select(i => categories[i]?.Value.Kind == CellValueKind.Number ? (double?)categories[i]!.Value.NumberValue : null).ToArray(),
                Totals = selected.Select(i => series.TotalPoints.Contains(i)).ToArray()
            });
        }
        if (output.Count > 1 && chart.Kind is WorksheetChartKind.Column or WorksheetChartKind.Line or WorksheetChartKind.SmoothLine &&
            output.Select(s => FormulaSyntax.Print(WorksheetChart.ParseReference(chart.Series[s.SeriesIndex].Categories))).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1)
        {
            // Different series can exclude different source points. Keep a shared category axis;
            // a missing value is a gap, never a shift onto the neighboring category.
            var positions = output.SelectMany(s => s.SourceIndices).Distinct().Order().ToArray();
            var labels = output.SelectMany(s => s.SourceIndices.Select((p, i) => (Position: p, Label: s.Categories[i]))).GroupBy(p => p.Position).ToDictionary(g => g.Key, g => g.First().Label);
            for (var s = 0; s < output.Count; s++)
            {
                var data = output[s]; if (data.SourceIndices.SequenceEqual(positions)) continue;
                var index = data.SourceIndices.Select((p, i) => (Position: p, Index: i)).ToDictionary(p => p.Position, p => p.Index);
                output[s] = data with { SourceIndices = positions, Categories = positions.Select(p => labels[p]).ToArray(),
                    Values = positions.Select(p => index.TryGetValue(p, out var i) ? data.Values[i] : null).ToArray(),
                    XValues = positions.Select(p => index.TryGetValue(p, out var i) ? data.XValues[i] : null).ToArray(),
                    PointColors = positions.Select(p => index.TryGetValue(p, out var i) ? data.PointColors[i] : data.Color).ToArray(),
                    Totals = positions.Select(p => index.TryGetValue(p, out var i) && data.Totals[i]).ToArray() };
            }
        }
        return output;
    }

    public static (double Start, double End)[] Waterfall(ChartSeriesData series)
    {
        var total = 0d;
        return series.Values.Select((value, i) =>
        {
            if (value is null) return (total, total);
            var start = i < series.Totals.Length && series.Totals[i] ? 0 : total;
            var end = start + (value ?? 0);
            if (!double.IsFinite(end)) throw new InvalidOperationException("瀑布图累计值超出数值范围。");
            total = end; return (start, end);
        }).ToArray();
    }
}
