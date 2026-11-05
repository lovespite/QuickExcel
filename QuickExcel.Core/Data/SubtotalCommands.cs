using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed record SubtotalGroupPreview(int FirstRow, int LastRow, int SummaryRow, string Label);
public sealed record SubtotalOptions(bool ReplaceCurrent = true, bool PageBreaks = false, bool SummaryBelow = true);
public sealed class SubtotalPreview
{
    internal SubtotalPreview(WorkbookSnapshot source, CommandDraft draft, IEnumerable<SubtotalGroupPreview> groups, int grandTotal)
    { Source = source; Draft = draft; Groups = groups.ToImmutableArray(); GrandTotalRow = grandTotal; }
    internal WorkbookSnapshot Source { get; }
    internal CommandDraft Draft { get; }
    public WorkbookSnapshot Result => Draft.Snapshot;
    public ImmutableArray<SubtotalGroupPreview> Groups { get; }
    public int GrandTotalRow { get; }
    public int InsertedRows => GrandTotalRow < 0 ? 0 : Groups.Length + 1;
    public int RemovedRows { get; internal init; }
    public CellRectangle Range { get; internal init; }
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
            result.ScriptSource, result.Styles, result.ScriptModules, result.FileState, result.DefinedNames, snapshot.ActiveSheetId, snapshot.Window, result.TextCompatibilityVersion, result.PackageFiles);
        return StructureEdits.Draft(snapshot, result);
    }
}

public static class SubtotalCommands
{
    private static readonly string[] Labels = ["", "平均值", "计数", "计数", "最大值", "最小值", "乘积", "标准偏差", "总体标准偏差", "汇总", "方差", "总体方差"];
    public static SubtotalPreview Preview(WorkbookSnapshot book, SheetId sheetId, CellRectangle range, int groupColumn, IEnumerable<int> totalColumns, int function = 9, SubtotalOptions? options = null)
    {
        options ??= new();
        var original = book;
        WorkbookClipboard.ValidateRange(range);
        if (range.Count > DataCleaning.MaxCells) throw new ArgumentException("分类汇总一次最多处理 20,000 格。");
        if (range.Bottom <= range.Top || groupColumn < range.Left || groupColumn > range.Right || function is < 1 or > 11)
            throw new ArgumentException("分类汇总需要标题和数据行、区域内分类列以及 1–11 的汇总函数。");
        var totals = totalColumns.Take(257).ToArray();
        if (totals.Length is 0 or > 256 || totals.Any(c => c < range.Left || c > range.Right || c == groupColumn) || totals.Distinct().Count() != totals.Length)
            throw new ArgumentException("请选择不重复的汇总列，且不能与分类列相同。");
        if (book.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
        var sheet = book.GetSheet(sheetId); StructureEdits.RequireSimpleStructure(sheet); StructureEdits.RequireTransformableFunctions(book);
        var existing = SummaryRows(sheet, range);
        var removed = 0;
        if (existing.Count > 0 && options.ReplaceCurrent)
        {
            var removal = Remove(book, sheetId, range); removed = removal.RemovedRows; book = removal.Result; range = removal.Range; sheet = book.GetSheet(sheetId);
            existing = [];
        }
        var priorRows = sheet.Rows;
        if (existing.Count > 0)
        {
            if ((sheet.Data.Outline?.SummaryBelow ?? true) != options.SummaryBelow) throw new ArgumentException("追加汇总须沿用现有汇总位置；改变位置时请勾选替换当前分类汇总。");
            if (priorRows.Values.Any(r => r.OutlineLevel >= 7)) throw new InvalidOperationException("追加汇总会超过 7 层大纲。");
            book = ClearOutline(book, sheet); sheet = book.GetSheet(sheetId);
        }
        if (!(sheet.Data with { ConditionalFormats = [], Outline = null, DefaultColumnWidth = null }).IsEmpty || WorksheetOutline.HasGroups(sheet) || sheet.Rows.Values.Any(r => r.Hidden) || sheet.Columns.Values.Any(c => c.Hidden) ||
            book.Sheets.Any(s => !s.Data.Validations.IsEmpty) ||
            SubtotalFunctions.HasUnmodeledFilters(book))
            throw new NotSupportedException("分类汇总需要无筛选、表、验证、手动分组、隐藏行列或未建模对象的普通区域；整批操作未应用。");
        if (sheet.Cells.Any(p => p.Value.Content.Kind == CellContentKind.Formula && SubtotalFunctions.Contains(FormulaSyntax.Parse(p.Value.Content.FormulaText)) && !existing.Contains(p.Key.Row)))
            throw new NotSupportedException("工作表包含不能识别为当前分类汇总的 SUBTOTAL，不能安全替换或嵌套。");
        string Key(int row)
        {
            var cell = sheet.GetCell(row, groupColumn);
            if (cell?.Content.Kind == CellContentKind.Formula && (cell.Calculation is not { State: CalculationState.Fresh or CalculationState.Error, Value: not null } || cell.Calculation.InputRevision != book.Revision))
                throw new InvalidOperationException("分类列含待重算公式，请先重算。");
            return WorksheetFilterEngine.ValueLabel(cell?.Value ?? CellValue.Blank, TablePresentation.NumberStyle(sheet, book.Styles, row, groupColumn), book.DateSystem);
        }
        var details = Enumerable.Range(range.Top + 1, range.Rows - 1).Where(r => !existing.Contains(r)).ToArray();
        if (details.Length == 0) throw new ArgumentException("区域内没有明细数据。");
        var groups = new List<(int First, int Last, string Label)>(); var start = range.Top + 1; var label = Key(details[0]);
        for (var i = 1; i <= details.Length; i++)
        {
            var row = i < details.Length ? details[i] : range.Bottom + 1;
            var next = i < details.Length ? Key(row) : null;
            if (next is not null && string.Equals(next, label, StringComparison.OrdinalIgnoreCase)) continue;
            if (!options.SummaryBelow) while (row > start && existing.Contains(row - 1)) row--;
            groups.Add((start, row - 1, label)); start = row; label = next ?? "";
        }
        if (groups.Count > 256) throw new InvalidOperationException("一次最多生成 256 个连续分类汇总，请先整理数据或缩小区域。");
        var insertions = (options.SummaryBelow ? groups.Select(g => g.Last + 1).Append(range.Bottom + 1) : groups.Select(g => g.First).Append(range.Top + 1)).Order().ToArray();
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
            var first = Position(group.First)!.Value; var last = Position(group.Last)!.Value; var summary = options.SummaryBelow ? last + 1 : first - 1;
            Summary(summary, first, last, group.Label + " " + Labels[function], group.Last);
            for (var oldRow = group.First; oldRow <= group.Last; oldRow++)
            {
                var row = Position(oldRow)!.Value;
                rows[row] = (rows.GetValueOrDefault(row) ?? new()) with { OutlineLevel = existing.Count == 0 ? 2 : Math.Max(2, (priorRows.GetValueOrDefault(oldRow)?.OutlineLevel ?? 0) + 1) };
            }
            rows[summary] = new() { OutlineLevel = 1 };
            previews.Add(new(first, last, summary, group.Label));
        }
        var lastDetail = Position(range.Bottom)!.Value; var grand = options.SummaryBelow ? lastDetail + 2 : range.Top + 1;
        Summary(grand, options.SummaryBelow ? range.Top + 1 : grand + 1, lastDetail, function is 2 or 3 ? "总计数" : function == 9 ? "总计" : "总计" + Labels[function], range.Bottom);
        transformed = WorksheetOutline.Replace(transformed, output, rows.ToImmutable(), output.Columns, new(SummaryBelow: options.SummaryBelow));
        if (options.PageBreaks)
        {
            output = transformed.GetSheet(sheetId);
            var breaks = previews.Skip(1).Select(g => new PrintPageBreak(options.SummaryBelow ? g.FirstRow : g.SummaryRow, 0, WorkbookLimits.MaxColumns - 1, true));
            transformed = transformed.WithSheets(transformed.Sheets.Select(s => s.Id == sheetId ? s.WithPrinting(s.Printing with { RowBreaks = s.Printing.RowBreaks.Concat(breaks).DistinctBy(b => b.Id).OrderBy(b => b.Id).ToImmutableArray() }) : s));
        }
        transformed = new(transformed.DocumentId, transformed.Revision, transformed.DateSystem, transformed.Sheets, transformed.ScriptSource, styles,
            transformed.ScriptModules, transformed.FileState, transformed.DefinedNames, transformed.ActiveSheetId, transformed.Window, transformed.TextCompatibilityVersion, transformed.PackageFiles);
        transformed = new SetCellsCommand(edits).Prepare(transformed).Snapshot;
        return new(original, StructureEdits.Draft(original, transformed), previews, grand) { RemovedRows = removed, Range = range with { Bottom = range.Bottom + groups.Count + 1 } };
    }

    private static HashSet<int> SummaryRows(WorksheetSnapshot sheet, CellRectangle range)
    {
        var groups = WorksheetOutline.Groups(sheet, SheetAxis.Row);
        var rows = new HashSet<int>();
        foreach (var group in groups)
        {
            if (group.Summary <= range.Top || group.Summary > range.Bottom) continue;
            var cells = sheet.Cells.Where(p => p.Key.Row == group.Summary && (p.Value.Content.Kind == CellContentKind.Formula || p.Value.Value.Kind != CellValueKind.Blank)).ToArray();
            var formulas = cells.Where(p => p.Value.Content.Kind == CellContentKind.Formula).ToArray();
            if (formulas.Length == 0) continue;
            var valid = formulas.All(p => FormulaSyntax.Parse(p.Value.Content.FormulaText) is FunctionNode { Name: "SUBTOTAL", Arguments.Count: 2 } f &&
                f.Arguments[0] is ValueNode { Value.Kind: CellValueKind.Number } n && n.Value.NumberValue is >= 1 and <= 11 &&
                f.Arguments[1] is ReferenceNode r && (r.Sheet is null || r.Sheet.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase)) &&
                r.Column == p.Key.Column && r.EndColumn == p.Key.Column && r.Row >= group.First && r.EndRow <= group.Last &&
                r.Row > range.Top && r.EndRow <= range.Bottom);
            if (!valid || cells.Any(p => p.Key.Column < range.Left || p.Key.Column > range.Right) || cells.Count(p => p.Value.Content.Kind == CellContentKind.Literal) != 1)
                throw new NotSupportedException("汇总行包含额外数据或自定义公式，无法安全删除或替换；请先检查完整区域。");
            rows.Add(group.Summary);
        }
        // When another aggregate is nested, a retained grand total can sit inside a new
        // outer group rather than at its boundary. Recognize only the generated total row shape.
        foreach (var row in sheet.Cells.Where(p => p.Key.Row > range.Top && p.Key.Row <= range.Bottom && p.Value.Content.Kind == CellContentKind.Literal &&
            p.Value.Value.Kind == CellValueKind.Text && p.Value.Value.TextValue.StartsWith("总计", StringComparison.Ordinal)).Select(p => p.Key.Row).Distinct())
        {
            if (rows.Contains(row) || !groups.Any(g => row >= g.First && row <= g.Last)) continue;
            var cells = sheet.Cells.Where(p => p.Key.Row == row && (p.Value.Content.Kind == CellContentKind.Formula || p.Value.Value.Kind != CellValueKind.Blank)).ToArray();
            var formulas = cells.Where(p => p.Value.Content.Kind == CellContentKind.Formula).ToArray();
            if (formulas.Length == 0) continue;
            if (cells.Any(p => p.Key.Column < range.Left || p.Key.Column > range.Right) || cells.Count(p => p.Value.Content.Kind == CellContentKind.Literal) != 1 ||
                formulas.Any(p => FormulaSyntax.Parse(p.Value.Content.FormulaText) is not FunctionNode { Name: "SUBTOTAL", Arguments.Count: 2 } f ||
                    f.Arguments[0] is not ValueNode { Value.Kind: CellValueKind.Number } n || n.Value.NumberValue is < 1 or > 11 ||
                    f.Arguments[1] is not ReferenceNode r || r.Column != p.Key.Column || r.EndColumn != p.Key.Column || r.Row <= range.Top || r.EndRow > range.Bottom ||
                    r.Row <= row && r.EndRow >= row || r.Sheet is not null && !r.Sheet.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase)))
                throw new NotSupportedException("总计行包含额外数据或自定义公式，不能安全删除。");
            rows.Add(row);
        }
        if (rows.Count > 0 && (groups.Any(g => g.Axis == SheetAxis.Row && (!rows.Contains(g.Summary) || g.First <= range.Top || g.Last > range.Bottom)) || WorksheetOutline.Groups(sheet, SheetAxis.Column).Count > 0))
            throw new NotSupportedException("请选择包含全部分类汇总的大纲区域；手动分组不能自动替换。");
        return rows;
    }

    private static WorkbookSnapshot ClearOutline(WorkbookSnapshot book, WorksheetSnapshot sheet) => WorksheetOutline.Replace(book, sheet,
        sheet.Rows.ToDictionary(p => p.Key, p => p.Value with { OutlineLevel = 0, Collapsed = false, Hidden = p.Value.Hidden && p.Value.OutlineLevel == 0 }), sheet.Columns, null);

    public static SubtotalPreview Remove(WorkbookSnapshot book, SheetId sheetId, CellRectangle range)
    {
        WorkbookClipboard.ValidateRange(range);
        if (book.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
        var sheet = book.GetSheet(sheetId); StructureEdits.RequireSimpleStructure(sheet); StructureEdits.RequireTransformableFunctions(book);
        var removed = SummaryRows(sheet, range);
        if (removed.Count == 0) throw new InvalidOperationException("当前区域没有可安全删除的分类汇总。");
        // Breaks at generated group boundaries are removed; unrelated print settings remain mapped.
        var boundaries = WorksheetOutline.Groups(sheet, SheetAxis.Row).Select(g => (sheet.Data.Outline?.SummaryBelow ?? true) ? g.First : g.Summary).ToHashSet();
        var clean = ClearOutline(book, sheet);
        clean = clean.WithSheets(clean.Sheets.Select(s => s.Id == sheetId ? s.WithPrinting(s.Printing with { RowBreaks = s.Printing.RowBreaks.Where(b => !boundaries.Contains(b.Id)).ToImmutableArray() }) : s));
        int? Position(int value) => removed.Contains(value) ? null : value - removed.Count(r => r < value);
        (int, int)? Interval(int first, int last)
        {
            while (first <= last && removed.Contains(first)) first++;
            while (last >= first && removed.Contains(last)) last--;
            return first > last ? null : (Position(first)!.Value, Position(last)!.Value);
        }
        var result = StructureEdits.Axis(clean, clean.GetSheet(sheetId), SheetAxis.Row, Position, Interval);
        return new(book, StructureEdits.Draft(book, result), [], -1) { RemovedRows = removed.Count, Range = range with { Bottom = range.Bottom - removed.Count } };
    }
}
