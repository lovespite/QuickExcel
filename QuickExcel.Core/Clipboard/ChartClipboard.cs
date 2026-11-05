using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

/// <summary>A chart and a bounded, independent snapshot of its source values. No source formulas or links travel.</summary>
public sealed class ChartClipboard
{
    public const int MaxCells = 20_000;
    public WorkbookSnapshot Book { get; }
    public WorksheetChart Chart { get; }
    private ChartClipboard(WorkbookSnapshot book, WorksheetChart chart) { Book = book; Chart = chart; }

    public static ChartClipboard Capture(WorkbookSnapshot book, WorksheetChart chart)
    {
        chart.Validate();
        var geometry = book.Sheets.FirstOrDefault(s => s.Data.Charts.Contains(chart)) ?? book.Sheets[0];
        var bounds = chart.Anchor.Bounds(geometry);
        var sheet = new WorksheetSnapshot(SheetId.New(), "图表数据");
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        var styles = new List<StyleDefinition> { new() };
        var references = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string Copy(string formula)
        {
            var r = WorksheetChart.ParseReference(formula); var key = FormulaSyntax.Print(r);
            if (references.TryGetValue(key, out var existing)) return existing;
            var source = book.Sheets.SingleOrDefault(s => s.Name.Equals(r.Sheet, StringComparison.OrdinalIgnoreCase)) ?? throw new InvalidOperationException(MessageText.Get("图表源工作表不存在。"));
            var count = WorksheetChart.Count(r); var column = references.Count;
            if (cells.Count + count > MaxCells) throw new NotSupportedException(MessageText.Get("图表剪贴板数据最多 20,000 格。"));
            for (var i = 0; i < count; i++)
            {
                var row = r.Row + (r.EndRow > r.Row ? i : 0); var col = r.Column + (r.EndColumn > r.Column ? i : 0);
                var cell = source.GetCell(row, col);
                if (cell?.Content.Kind == CellContentKind.Formula && cell.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error))
                    throw new InvalidOperationException(MessageText.Get("图表源公式尚未重算，不能复制旧结果。"));
                var style = book.Styles[source.GetStyleId(row, col)] with { NamedStyleId = 0, FontScheme = null };
                var id = styles.IndexOf(style); if (id < 0) { id = styles.Count; styles.Add(style); }
                cells.Add(new(i, column), new(CellContent.FromLiteral(cell?.Value ?? CellValue.Blank, id)));
            }
            var mapped = FormulaSyntax.Print(new ReferenceNode(sheet.Name, 0, column, count - 1, column, true, true, true, true, IsRange: count > 1));
            references.Add(key, mapped); return mapped;
        }
        // Materialize row/column and worksheet-filter visibility as chart point filters in the independent data sheet.
        var visible = ChartData.Read(book, chart with { Series = [.. chart.Series.Select(s => s with { IsFiltered = false, HiddenPoints = [] })] });
        var copied = chart.MapReferences(Copy) with
        {
            // Re-emit modeled appearance with resolved colors/fonts: source theme indices are not portable.
            OriginalXml = [], StyleParts = [],
            Anchor = DrawingGeometry.FromBounds(sheet, new(0, 0, bounds.Width, bounds.Height), DrawingAnchorKind.OneCell, chart.Anchor.Rotation),
            Series = [.. chart.MapReferences(Copy).Series.Select((s, i) => s with
            {
                HiddenPoints = [.. s.HiddenPoints.Union(Enumerable.Range(0, WorksheetChart.Count(WorksheetChart.ParseReference(s.Values))).Except(visible[i].SourceIndices)).Order()]
            })]
        };
        sheet = new(sheet.Id, sheet.Name, cells);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, book.DateSystem, [sheet], styles: styles);
        copied.Validate(); return new(snapshot, copied);
    }
}

/// <summary>Insert the editable chart and its value data sheet as one undoable change.</summary>
public sealed class PasteChartCommand(ChartClipboard source, SheetId sheetId, CellPosition position) : DocumentCommand
{
    private readonly Guid _chartId = Guid.NewGuid();
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var destination = snapshot.GetSheet(sheetId); WorkbookLimits.ValidatePosition(position.Row, position.Column);
        var name = "图表数据"; var suffix = 1;
        while (snapshot.Sheets.Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase))) name = "图表数据" + ++suffix;
        var data = source.Book.Sheets.Single(); var styles = snapshot.Styles.ToList(); var map = new int[source.Book.Styles.Length];
        for (var i = 0; i < map.Length; i++)
        {
            var id = styles.IndexOf(source.Book.Styles[i]); if (id < 0) { id = styles.Count; styles.Add(source.Book.Styles[i]); } map[i] = id;
        }
        var imported = new WorksheetSnapshot(SheetId.New(), name, data.Cells.Select(p => new KeyValuePair<CellPosition, CellSnapshot>(p.Key, new(p.Value.Content.WithStyle(map[p.Value.Content.StyleId])))));
        var chart = source.Chart.MapReferences(f => FormulaSyntax.Print(WorksheetChart.ParseReference(f) with { Sheet = name }));
        var bounds = chart.Anchor.Bounds(data);
        chart = chart with { ScriptId = _chartId, Anchor = DrawingGeometry.FromBounds(destination, new(DrawingGeometry.ColumnStart(destination, position.Column), DrawingGeometry.RowStart(destination, position.Row), bounds.Width, bounds.Height), DrawingAnchorKind.OneCell, chart.Anchor.Rotation) };
        // Values are independent snapshots. Preserve numeric/date-axis interpretation across workbook epochs.
        if (source.Book.DateSystem != snapshot.DateSystem && source.Chart.Series.Any(s => s.NumericCategories))
            throw new NotSupportedException(MessageText.Get("数值分类图表的数据日期系统与目标不同；请选择粘贴为图片。"));
        var staged = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, snapshot.Sheets.Add(imported), snapshot.ScriptSource, styles,
            snapshot.ScriptModules, snapshot.FileState, snapshot.DefinedNames, snapshot.ActiveSheetId, snapshot.Window, snapshot.TextCompatibilityVersion, snapshot.PackageFiles);
        var result = new SetChartCommand(sheetId, destination.Data.Charts.Length, chart).Prepare(staged);
        return result with { StructureChanged = true, CalculationUnchanged = false };
    }
}
