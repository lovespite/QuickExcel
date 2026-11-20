using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

/// <summary>Cut a rectangle, including its complete merges and tables, from one immutable workbook state.</summary>
public sealed class MoveCellsCommand(SheetId sourceSheet, CellPosition first, CellPosition last,
    SheetId destinationSheet, CellPosition destination) : DocumentCommand
{
    private readonly CellRectangle _source = new(first.Row, first.Column, last.Row, last.Column);
    private readonly int _rows = destination.Row - first.Row;
    private readonly int _columns = destination.Column - first.Column;

    internal static void ValidateSource(WorkbookSnapshot book, WorksheetSnapshot sheet, CellRectangle range)
    {
        WorkbookClipboard.RequireCompleteMerges(sheet, range);
        if (sheet.Data.Tables.Any(t => ValidationSafety.Overlaps(t.Range, range) && !range.Contains(t.Range)))
            throw new InvalidOperationException(MessageText.Get("剪切需要选中完整的 Excel 表格，包括表头和汇总行。"));
        ValidateArea(book, sheet, range);
        RequireCompleteSpills(sheet, range, null);
    }

    private static void ValidateArea(WorkbookSnapshot book, WorksheetSnapshot sheet, CellRectangle range)
    {
        WorkbookClipboard.RequireNoDrawings(sheet, range);
        WorkbookClipboard.RequireVisibleDimensions(sheet, range);
        if (sheet.Printing.HasReferences)
            throw new NotSupportedException(MessageText.Get("含打印区域或重复标题的工作表暂不支持移动单元格。"));
        if (sheet.Data.Protection?.Sheet == true)
            throw new InvalidOperationException(MessageText.Get("请先取消工作表保护，再移动单元格。"));
        if (!sheet.Data.Validations.IsEmpty || !sheet.Data.ConditionalFormats.IsEmpty ||
            book.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("conditionalFormatting", StringComparison.OrdinalIgnoreCase) ||
                d.Feature.Contains("dataValidation", StringComparison.OrdinalIgnoreCase) || d.Feature.Contains("sortState", StringComparison.OrdinalIgnoreCase)) == true ||
            SubtotalFunctions.HasUnmodeledFilters(book))
            throw new NotSupportedException(MessageText.Get("工作表含验证或条件格式，或文件含尚未建模的表格、筛选、排序或规则；暂不能安全剪切。"));
        if (sheet.Data.Filter is { } filter && ValidationSafety.Overlaps(range, filter.Range) ||
            sheet.Data.Sort is { } sort && ValidationSafety.Overlaps(range, sort.Range))
            throw new NotSupportedException(MessageText.Get("剪切区域与独立的工作表筛选或排序区域相交，请先清除该状态。"));
        if (WorksheetFilterEngine.Evaluate(book, sheet.Id, WorksheetFilterEngine.CanUseSourceCaches(book)).Any(r => r >= range.Top && r <= range.Bottom))
            throw new InvalidOperationException(MessageText.Get("剪切区域含筛选隐藏行，请先清除筛选条件。"));
    }

    private static void RequireCompleteSpills(WorksheetSnapshot sheet, CellRectangle area, CellRectangle? vacated)
    {
        foreach (var (position, cell) in sheet.Cells)
        {
            if (cell.Calculation?.Array is not { } array) continue;
            var spill = new CellRectangle(position.Row, position.Column, position.Row + array.Rows - 1, position.Column + array.Columns - 1);
            if (ValidationSafety.Overlaps(area, spill) && !area.Contains(spill) && vacated?.Contains(spill) != true)
                throw new InvalidOperationException(MessageText.Get("剪切需要包含整个数组溢出区域。"));
        }
    }

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorkbookLimits.ValidatePosition(first.Row, first.Column);
        WorkbookLimits.ValidatePosition(last.Row, last.Column);
        if (first.Row > last.Row || first.Column > last.Column) throw new ArgumentException(MessageText.Get("移动区域顺序无效。"));
        var output = _source.Offset(_rows, _columns);
        WorkbookLimits.ValidatePosition(output.Top, output.Left);
        WorkbookLimits.ValidatePosition(output.Bottom, output.Right);
        var source = snapshot.GetSheet(sourceSheet);
        var target = snapshot.GetSheet(destinationSheet);
        var sameSheet = sourceSheet == destinationSheet;
        if (sameSheet && _source == output) return new(snapshot, [], false);

        ValidateSource(snapshot, source, _source);
        ValidateArea(snapshot, target, output);
        RequireCompleteSpills(target, output, sameSheet ? _source : null);
        var movedTables = source.Data.Tables.Where(t => _source.Contains(t.Range)).Select(t => t.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (target.Data.Tables.Any(t => !movedTables.Contains(t.Name) && ValidationSafety.Overlaps(output, t.Range)))
            throw new InvalidOperationException(MessageText.Get("目标与现有 Excel 表格相交，请选择表格以外的位置。"));
        var movedMerges = source.MergedCells.Select(CellRectangle.Parse).Where(_source.Contains).ToArray();
        if (target.MergedCells.Select(CellRectangle.Parse).Any(m => ValidationSafety.Overlaps(m, output) && !output.Contains(m) && !(sameSheet && _source.Contains(m))))
            throw new InvalidOperationException(MessageText.Get("目标只覆盖了合并单元格的一部分，请选择能容纳整个区域的位置。"));
        foreach (var (p, cell) in target.Cells)
            if (output.Contains(p.Row, p.Column) && !(sameSheet && _source.Contains(p.Row, p.Column)) && cell.Content.RichText is not null &&
                source.GetCell(p.Row - _rows, p.Column - _columns)?.Content.RichText is null)
                throw new InvalidOperationException(MessageText.Get("剪切目标包含富文本，请先明确转换为普通文本；源区域尚未修改。"));

        var tableNames = snapshot.Sheets.SelectMany(s => s.Data.Tables).Select(t => t.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        FormulaNode Transform(FormulaNode node, SheetId context)
        {
            // Both Orders and Orders[Qty] can carry an explicit sheet qualifier.
            bool MovedTable(string? sheet, string? table) => table is not null && movedTables.Contains(table) &&
                string.Equals(sheet, source.Name, StringComparison.OrdinalIgnoreCase);
            if (node is StructuredReferenceNode structured && MovedTable(structured.Sheet, structured.Table))
                return structured with { Sheet = target.Name };
            if (node is NameNode name && MovedTable(name.Sheet, name.Name))
                // Excel reads Sheet!Orders as a defined name. The globally unique structured
                // reference Orders[#Data] retains the table binding, including inside LET.
                return new StructuredReferenceNode(name.Name);
            if (node is not ReferenceNode r || !string.Equals(r.Sheet ?? snapshot.GetSheet(context).Name, source.Name, StringComparison.OrdinalIgnoreCase))
                return node;
            var reference = new CellRectangle(r.Row, r.Column, r.EndRow, r.EndColumn);
            if (!ValidationSafety.Overlaps(reference, _source)) return node;
            if (!_source.Contains(reference))
                throw new NotSupportedException(MessageText.Get("移动仅覆盖了引用区域的一部分，无法保留连续区域语义，操作未应用。"));
            return r with { Sheet = target.Name, Row = r.Row + _rows, EndRow = r.EndRow + _rows, Column = r.Column + _columns, EndColumn = r.EndColumn + _columns };
        }

        // Mapped A1 references now point explicitly to their destination. Remaining unqualified
        // A1 references and sheet-scoped names in a travelling formula still belong to its old sheet.
        string InTargetContext(string formula)
        {
            if (sameSheet) return formula;
            var expression = FormulaSyntax.Map(FormulaSyntax.Parse(formula), node => node switch
            {
                ReferenceNode { Sheet: null } r => r with { Sheet = source.Name },
                NameNode { Sheet: null } n when !tableNames.Contains(n.Name) => n with { Sheet = source.Name },
                _ => node
            });
            return "=" + FormulaSyntax.Print(expression);
        }

        var transformed = StructureEdits.Transform(snapshot, Transform);
        var captured = transformed.GetSheet(sourceSheet);
        var tables = captured.Data.Tables.Where(t => movedTables.Contains(t.Name)).Select(t => MoveTable(t, InTargetContext)).ToArray();
        var result = transformed.WithSheets(transformed.Sheets.Select(sheet =>
        {
            var from = sheet.Id == sourceSheet;
            var to = sheet.Id == destinationSheet;
            if (!from && !to) return sheet;
            var cells = sheet.Cells.ToBuilder();
            foreach (var p in sheet.Cells.Keys)
                if (from && _source.Contains(p.Row, p.Column) || to && output.Contains(p.Row, p.Column)) cells.Remove(p);
            if (to)
            {
                foreach (var (p, cell) in captured.Cells)
                {
                    if (!_source.Contains(p.Row, p.Column)) continue;
                    cells[new(p.Row + _rows, p.Column + _columns)] = !sameSheet && cell.Content.Kind == CellContentKind.Formula
                        ? new(cell.Content.WithFormula(InTargetContext(cell.Content.FormulaText))) : cell;
                }
                MoveInheritedBlankStyles(source, target, output, cells);
            }
            var merges = sheet.MergedCells.Where(text =>
            {
                var merge = CellRectangle.Parse(text);
                return !(from && _source.Contains(merge) || to && output.Contains(merge));
            }).ToList();
            if (to) merges.AddRange(movedMerges.Select(m => m.Offset(_rows, _columns).ToString()));
            var remaining = sheet.Data.Tables.Where(t => !movedTables.Contains(t.Name));
            var ignored = from ? IgnoredErrorRanges.Replace(sheet.Data.IgnoredErrors, _source) : sheet.Data.IgnoredErrors;
            if (to) ignored = IgnoredErrorRanges.Replace(ignored, output,
                IgnoredErrorRanges.Clip(captured.Data.IgnoredErrors, _source).Select(e => e with { Range = e.Range.Offset(_rows, _columns) }));
            var data = sheet.Data with { Tables = (to ? remaining.Concat(tables) : remaining).ToImmutableArray(), IgnoredErrors = ignored };
            return new WorksheetSnapshot(sheet.Id, sheet.Name, cells, sheet.Rows, sheet.Columns, merges, sheet.Images, sheet.ShowGridLines,
                sheet.DefaultRowHeight, sheet.TabColor, sheet.View, sheet.Visibility, sheet.Printing, sheet.Annotations, data);
        }));
        return StructureEdits.Draft(snapshot, result);
    }

    private WorksheetTable MoveTable(WorksheetTable table, Func<string, string> inTargetContext)
    {
        var filter = table.Filter is { } f ? new WorksheetAutoFilter(f.Range.Offset(_rows, _columns),
            f.Columns.ToDictionary(p => p.Key + _columns, p => p.Value)) : null;
        var sort = table.Sort is { } s ? new WorksheetSortState(s.Range.Offset(_rows, _columns),
            s.Keys.Select(k => new SortKey(k.Column + _columns, k.Descending, k.CaseSensitive, k.Basis, k.PreferredColor, k.CustomOrder)), s.SortMethod) : null;
        return new(table.Id, table.Name, table.Range.Offset(_rows, _columns), table.Columns.Select(c => c with
        {
            CalculatedFormula = c.CalculatedFormula is { } calculated ? inTargetContext(calculated) : null,
            TotalsFormula = c.TotalsFormula is { } total ? inTargetContext(total) : null
        }), table.HeaderRow, table.TotalsRow, table.TotalsRowShown, table.Style, filter, sort, table.InsertRow);
    }

    private void MoveInheritedBlankStyles(WorksheetSnapshot source, WorksheetSnapshot target, CellRectangle output,
        ImmutableDictionary<CellPosition, CellSnapshot>.Builder cells)
    {
        static bool HasStyle(WorksheetSnapshot sheet, CellRectangle range) =>
            sheet.Rows.Any(p => p.Key >= range.Top && p.Key <= range.Bottom && p.Value.ApplyStyle && p.Value.StyleId != 0) ||
            sheet.Columns.Any(p => p.Key >= range.Left && p.Key <= range.Right && p.Value.StyleId != 0);
        if (!HasStyle(source, _source) && !HasStyle(target, output)) return;
        if (_source.Count > WorkbookClipboard.MaxCells) throw new NotSupportedException(MessageText.Get("移动含继承格式的区域最多支持 20,000 格。"));
        for (var row = _source.Top; row <= _source.Bottom; row++)
        for (var column = _source.Left; column <= _source.Right; column++)
        {
            if (source.Cells.ContainsKey(new(row, column))) continue;
            var style = source.GetInheritedStyleId(row, column);
            var to = new CellPosition(row + _rows, column + _columns);
            if (style != target.GetInheritedStyleId(to.Row, to.Column)) cells[to] = new(CellContent.FromLiteral(CellValue.Blank, style));
        }
    }
}
