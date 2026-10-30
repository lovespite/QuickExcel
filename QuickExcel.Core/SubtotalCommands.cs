using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed record SubtotalGroupPreview(int FirstRow, int LastRow, int SummaryRow, string Label);
public sealed class SubtotalPreview
{
    internal SubtotalPreview(WorkbookSnapshot source, CommandDraft draft, IEnumerable<SubtotalGroupPreview> groups, int grandTotal)
    { Source = source; Draft = draft; Groups = groups.ToImmutableArray(); GrandTotalRow = grandTotal; }
    internal WorkbookSnapshot Source { get; }
    internal CommandDraft Draft { get; }
    public WorkbookSnapshot Result => Draft.Snapshot;
    public ImmutableArray<SubtotalGroupPreview> Groups { get; }
    public int GrandTotalRow { get; }
    public int InsertedRows => Groups.Length + 1;
}

public sealed class ApplySubtotalsCommand(SubtotalPreview preview) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(preview);
        if (!ApplyCleaningCommand.SameInputs(snapshot, preview.Source)) throw new InvalidOperationException("分类汇总预览已过期，请重新预览；没有修改工作簿。");
        var result = preview.Result;
        // Keep navigation that may have changed during the preview dialog.
        result = new(result.DocumentId, result.Revision, result.DateSystem, result.Sheets.Select(s => s.WithView(snapshot.GetSheet(s.Id).View)),
            result.ScriptSource, result.Styles, result.ScriptModules, result.FileState, result.DefinedNames, snapshot.ActiveSheetId, snapshot.Window, result.TextCompatibilityVersion);
        return StructureEdits.Draft(snapshot, result);
    }
}

public static class SubtotalCommands
{
    private static readonly string[] Labels = ["", "平均值", "计数", "计数", "最大值", "最小值", "乘积", "标准偏差", "总体标准偏差", "汇总", "方差", "总体方差"];
    public static SubtotalPreview Preview(WorkbookSnapshot book, SheetId sheetId, CellRectangle range, int groupColumn, IEnumerable<int> totalColumns, int function = 9)
    {
        WorkbookClipboard.ValidateRange(range);
        if (range.Bottom <= range.Top || groupColumn < range.Left || groupColumn > range.Right || function is < 1 or > 11)
            throw new ArgumentException("分类汇总需要标题和数据行、区域内分类列以及 1–11 的汇总函数。");
        var totals = totalColumns.Take(257).ToArray();
        if (totals.Length is 0 or > 256 || totals.Any(c => c < range.Left || c > range.Right || c == groupColumn) || totals.Distinct().Count() != totals.Length)
            throw new ArgumentException("请选择不重复的汇总列，且不能与分类列相同。");
        if (book.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
        var sheet = book.GetSheet(sheetId); StructureEdits.RequireSimpleStructure(sheet); StructureEdits.RequireTransformableFunctions(book);
        if (!sheet.Data.IsEmpty || WorksheetOutline.HasGroups(sheet) || sheet.Rows.Values.Any(r => r.Hidden) || sheet.Columns.Values.Any(c => c.Hidden) ||
            !sheet.Printing.RowBreaks.IsEmpty || !sheet.Printing.ColumnBreaks.IsEmpty || book.Sheets.Any(s => !s.Data.Validations.IsEmpty) ||
            SubtotalFunctions.HasUnmodeledFilters(book))
            throw new NotSupportedException("分类汇总仅处理无筛选、表、分组、验证、条件格式、隐藏行列或未建模对象的普通区域；请先复制到普通工作表。");
        if (sheet.Cells.Values.Any(c => c.Content.Kind == CellContentKind.Formula && SubtotalFunctions.Contains(FormulaSyntax.Parse(c.Content.FormulaText))))
            throw new NotSupportedException("工作表已有 SUBTOTAL；本次不替换或嵌套生成既有分类汇总。");
        string Key(int row)
        {
            var cell = sheet.GetCell(row, groupColumn);
            if (cell?.Content.Kind == CellContentKind.Formula && (cell.Calculation is not { State: CalculationState.Fresh or CalculationState.Error, Value: not null } || cell.Calculation.InputRevision != book.Revision))
                throw new InvalidOperationException("分类列含待重算公式，请先重算。");
            return WorksheetFilterEngine.ValueLabel(cell?.Value ?? CellValue.Blank, TablePresentation.NumberStyle(sheet, book.Styles, row, groupColumn), book.DateSystem);
        }
        var groups = new List<(int First, int Last, string Label)>(); var start = range.Top + 1; var label = Key(start);
        for (var row = start + 1; row <= range.Bottom + 1; row++)
        {
            var next = row <= range.Bottom ? Key(row) : null;
            if (next is not null && string.Equals(next, label, StringComparison.OrdinalIgnoreCase)) continue;
            groups.Add((start, row - 1, label)); start = row; label = next ?? "";
        }
        if (groups.Count > 256) throw new InvalidOperationException("一次最多生成 256 个连续分类汇总，请先整理数据或缩小区域。");
        var insertions = groups.Select(g => g.Last + 1).Append(range.Bottom + 1).Order().ToArray();
        int? Position(int value) => checked(value + insertions.Count(i => i <= value));
        (int, int)? Interval(int first, int last) => (Position(first)!.Value, Position(last)!.Value);
        var transformed = StructureEdits.Axis(book, sheet, SheetAxis.Row, Position, Interval);
        var output = transformed.GetSheet(sheetId); var styles = transformed.Styles.ToList(); var edits = new List<CellEdit>();
        var rows = output.Rows.ToBuilder(); var previews = new List<SubtotalGroupPreview>();
        void Summary(int row, int from, int to, string text, int styleSource)
        {
            for (var col = range.Left; col <= range.Right; col++)
            {
                var style = book.Styles[sheet.GetStyleId(styleSource, col)] with { Bold = true }; var id = styles.IndexOf(style);
                if (id < 0) { if (styles.Count >= 4096) throw new InvalidOperationException("汇总样式超过 4096。"); id = styles.Count; styles.Add(style); }
                var column = FormulaSyntax.ColumnName(col);
                var content = col == groupColumn ? CellContent.FromLiteral(CellValue.FromText(text), id) : totals.Contains(col) ?
                    CellContent.FromFormula($"=SUBTOTAL({function},{column}{from + 1}:{column}{to + 1})", id) : CellContent.FromLiteral(CellValue.Blank, id);
                edits.Add(new(new(sheetId, row, col), content));
            }
        }
        foreach (var group in groups)
        {
            var first = Position(group.First)!.Value; var last = Position(group.Last)!.Value; var summary = last + 1;
            Summary(summary, first, last, group.Label + " " + Labels[function], group.Last);
            for (var row = first; row <= last; row++) rows[row] = (rows.GetValueOrDefault(row) ?? new()) with { OutlineLevel = 2 };
            rows[summary] = new() { OutlineLevel = 1 };
            previews.Add(new(first, last, summary, group.Label));
        }
        var lastDetail = Position(range.Bottom)!.Value; var grand = lastDetail + 2;
        Summary(grand, range.Top + 1, lastDetail, function is 2 or 3 ? "总计数" : function == 9 ? "总计" : "总计" + Labels[function], range.Bottom);
        transformed = WorksheetOutline.Replace(transformed, output, rows.ToImmutable(), output.Columns, new());
        transformed = new(transformed.DocumentId, transformed.Revision, transformed.DateSystem, transformed.Sheets, transformed.ScriptSource, styles,
            transformed.ScriptModules, transformed.FileState, transformed.DefinedNames, transformed.ActiveSheetId, transformed.Window, transformed.TextCompatibilityVersion);
        transformed = new SetCellsCommand(edits).Prepare(transformed).Snapshot;
        return new(book, StructureEdits.Draft(book, transformed), previews, grand);
    }
}
