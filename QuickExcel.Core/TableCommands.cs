using QuickExcel.Domain;

namespace QuickExcel.Core;

internal static class TableSafety
{
    public static void Validate(WorkbookSnapshot snapshot)
    {
        var ids = new HashSet<int>(); var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase); var count = 0;
        var styles = new Dictionary<string, TableStyleInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in snapshot.Sheets)
        {
            var occupied = new List<CellRectangle>();
            foreach (var table in sheet.Data.Tables)
            {
                if (++count > 256 || !ids.Add(table.Id) || !names.Add(table.Name)) throw new ArgumentException("表数量超过 256，或表编号/名称重复。");
                if (snapshot.DefinedNames.Any(n => n.Name.Equals(table.Name, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("表名与定义名称冲突。");
                if (occupied.Any(r => Overlaps(r, table.Range)) || sheet.MergedCells.Select(CellRectangle.Parse).Any(r => Overlaps(r, table.Range)) ||
                    sheet.Data.Filter is { } filter && Overlaps(filter.Range, table.Range)) throw new ArgumentException("表不能与其他表、工作表筛选或合并单元格重叠。");
                occupied.Add(table.Range);
                if (table.Style.Name is { } styleName)
                {
                    if (styles.TryGetValue(styleName, out var existing) && !existing.Elements.SequenceEqual(table.Style.Elements))
                        throw new ArgumentException("同名表样式必须具有相同定义。");
                    styles[styleName] = table.Style;
                }
            }
        }
    }
    private static bool Overlaps(CellRectangle a, CellRectangle b) => a.Top <= b.Bottom && b.Top <= a.Bottom && a.Left <= b.Right && b.Left <= a.Right;
    public static void ValidateEdit(WorkbookSnapshot before, WorkbookSnapshot after)
    {
        foreach (var sheet in before.Sheets)
        {
            var next = after.Sheets.FirstOrDefault(s => s.Id == sheet.Id); if (next is null) continue;
            foreach (var table in sheet.Data.Tables)
            {
                var nextTable = next.Data.Tables.FirstOrDefault(t => t.Id == table.Id);
                if (nextTable is null || nextTable.Range != table.Range) continue;
                for (var column = table.Range.Left; column <= table.Range.Right; column++)
                {
                    bool Changed(int row)
                    {
                        var old = sheet.GetCell(row, column)?.Content; var current = next.GetCell(row, column)?.Content;
                        return old?.Kind != current?.Kind || (old?.Kind == CellContentKind.Formula ? old.FormulaText != current!.FormulaText : old?.LiteralValue != current?.LiteralValue);
                    }
                    if (table.HeaderRow && nextTable.Columns[column - table.Range.Left].Name == table.Columns[column - table.Range.Left].Name && Changed(table.Range.Top))
                        throw new InvalidOperationException("当前表头受控编辑尚不支持改列名，请保持表头内容；本次修改未应用。");
                    var definition = table.Columns[column - table.Range.Left];
                    if (table.TotalsRow && definition == nextTable.Columns[column - table.Range.Left] &&
                        (definition.TotalsFunction is not null || definition.TotalsFormula is not null || definition.TotalsLabel is not null) && Changed(table.Range.Bottom))
                        throw new InvalidOperationException("当前汇总行定义受保护，请在 Excel 修改汇总标签/公式；本次修改未应用。");
                }
            }
        }
    }
}

public sealed class CreateTableCommand(SheetId sheetId, CellRectangle range, string name, TableStyleInfo? style = null) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        if (range.Bottom <= range.Top) throw new ArgumentException("创建表需要表头和至少一行数据。");
        var columns = Enumerable.Range(range.Left, range.Right - range.Left + 1).Select(c =>
        {
            var cell = sheet.GetCell(range.Top, c);
            if (cell?.Content.Kind != CellContentKind.Literal || cell.Value.Kind != CellValueKind.Text || cell.Content.RichText is not null)
                throw new ArgumentException("表头须是非空且不重复的普通文本；请先整理表头。");
            return new TableColumn(c - range.Left + 1, cell.Value.TextValue);
        }).ToArray();
        var id = checked(snapshot.Sheets.SelectMany(s => s.Data.Tables).Select(t => t.Id).DefaultIfEmpty().Max() + 1);
        var table = new WorksheetTable(id, name, range, columns, style: style, filter: new(range));
        var data = sheet.Data with { Tables = sheet.Data.Tables.Add(table) };
        var next = snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(data) : s));
        return new(next, [], true);
    }
}

/// <summary>First table slice: resize the bottom boundary only; total rows and calculated columns need table-aware editing.</summary>
public sealed class ResizeTableCommand(SheetId sheetId, int tableId, int bottomRow) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId); var table = sheet.Data.Tables.Single(t => t.Id == tableId);
        if (table.Range.Bottom == bottomRow) return new(snapshot, [], false);
        if (table.TotalsRow) throw new NotSupportedException("含汇总行的表请使用表内插删行，底边调整暂未支持移动汇总格；本次没有修改。");
        if (bottomRow < table.Range.Top + (table.HeaderRow ? 1 : 0)) throw new ArgumentException("表范围至少保留一行数据。");
        var range = table.Range with { Bottom = bottomRow };
        var working = bottomRow < table.Range.Bottom ? StructureEdits.TransformAt(snapshot, (node, context, position) =>
            node is QuickExcel.Formulas.StructuredReferenceNode { Table: null } reference && context == sheetId && position is { } p && !range.Contains(p.Row, p.Column) &&
                QuickExcel.Formulas.StructuredReferences.Resolve(snapshot, context, position, reference)?.Table.Id == tableId ? reference with { Table = table.Name } : node) : snapshot;
        var resized = table.WithRange(range, table.Filter is null ? null : new(range, table.Filter.Columns));
        var data = sheet.Data with { Tables = sheet.Data.Tables.Replace(table, resized) };
        var next = working.WithSheets(working.Sheets.Select(s => s.Id == sheetId ? s.WithData(data) : s));
        next = TableStructure.FillInserted(next, sheet, p => p.Row > table.Range.Bottom && p.Row <= bottomRow);
        _ = WorksheetFilterEngine.Evaluate(next, sheetId);
        return StructureEdits.Draft(snapshot, next);
    }
}

public sealed class SetTableFilterCommand(SheetId sheetId, int tableId, WorksheetAutoFilter? filter) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorksheetOutline.RequireEditable(snapshot);
        var sheet = snapshot.GetSheet(sheetId); var table = sheet.Data.Tables.Single(t => t.Id == tableId); var updated = table.WithFilter(filter);
        if (table.ContentEquals(updated)) return new(snapshot, [], false);
        var next = snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithData(sheet.Data with { Tables = sheet.Data.Tables.Replace(table, updated) }) : s));
        _ = WorksheetFilterEngine.Evaluate(next, sheetId);
        return new(next, [], false) { CalculationUnchanged = true };
    }
}
