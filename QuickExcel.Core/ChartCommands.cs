using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed class SetChartCommand(SheetId sheetId, int index, WorksheetChart? chart) : DocumentCommand
{
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
        if (range.Rows < 2 || range.Columns < 2 || range.Columns > 17 || kind == WorksheetChartKind.Pie && range.Columns != 2)
            throw new ArgumentException("首行为系列标题，首列为分类；需要 2–17 列，饼图需要 2 列。");
        string Ref(int top, int left, int bottom, int right) => FormulaSyntax.Print(new ReferenceNode(sheet.Name, top, left, bottom, right, true, true, true, true, IsRange: top != bottom || left != right));
        string[] palette = ["FF4472C4", "FFED7D31", "FFA5A5A5", "FFFFC000", "FF5B9BD5", "FF70AD47"];
        return new() { Kind = kind, Title = title, Anchor = anchor, Series = [.. Enumerable.Range(range.Left + 1, range.Columns - 1).Select((c, i) =>
            new ChartSeries("", Ref(range.Top, c, range.Top, c), Ref(range.Top + 1, range.Left, range.Bottom, range.Left), Ref(range.Top + 1, c, range.Bottom, c), palette[i % palette.Length])
            { PointColors = kind == WorksheetChartKind.Pie ? [.. Enumerable.Range(0, range.Rows - 1).Select(p => palette[p % palette.Length])] : [] })] };
    }
}

public sealed record ChartSeriesData(string Name, string[] Categories, double?[] Values, string Color, string[] PointColors);
public static class ChartData
{
    public static IReadOnlyList<ChartSeriesData> Read(WorkbookSnapshot book, WorksheetChart chart)
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
        foreach (var series in chart.Series)
        {
            var values = Read(series.Values, out var valueSheet, out var vr); var categories = Read(series.Categories, out var categorySheet, out var cr);
            var hidden = chart.VisibleOnly ? WorksheetFilterEngine.Evaluate(book, valueSheet.Id) : [];
            var selected = Enumerable.Range(0, values.Length).Where(i => !chart.VisibleOnly || !hidden.Contains(vr.Row + (vr.EndRow > vr.Row ? i : 0)) && valueSheet.Rows.GetValueOrDefault(vr.Row + (vr.EndRow > vr.Row ? i : 0))?.Hidden != true && valueSheet.Columns.GetValueOrDefault(vr.Column + (vr.EndColumn > vr.Column ? i : 0))?.Hidden != true).ToArray();
            var name = series.Name;
            if (series.NameReference is { } title) { var cell = Read(title, out var s, out var r)[0]; name = Display(cell, s, r.Row, r.Column); }
            output.Add(new(name, selected.Select(i => Display(categories[i], categorySheet, cr.Row + (cr.EndRow > cr.Row ? i : 0), cr.Column + (cr.EndColumn > cr.Column ? i : 0))).ToArray(),
                selected.Select(i => values[i]?.Value.Kind == CellValueKind.Number ? (double?)values[i]!.Value.NumberValue : null).ToArray(), series.Color,
                selected.Select(i => series.PointColors.IsEmpty ? series.Color : series.PointColors[i]).ToArray()));
        }
        return output;
    }
}
