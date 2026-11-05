using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed class RenameTableCommand(SheetId sheetId, int tableId, string name) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorksheetTable.ValidateName(name); var table = snapshot.GetSheet(sheetId).Data.Tables.Single(t => t.Id == tableId);
        if (table.Name == name) return new(snapshot, [], false);
        var updated = StructureEdits.Transform(snapshot, (node, _) => node switch
        {
            StructuredReferenceNode r when string.Equals(r.Table, table.Name, StringComparison.OrdinalIgnoreCase) => r with { Table = name },
            NameNode n when n.Name.Equals(table.Name, StringComparison.OrdinalIgnoreCase) => n with { Name = name },
            _ => node
        }, (sheet, cells) => sheet.WithCells(cells).WithData(sheet.Data with
        { Tables = sheet.Data.Tables.Select(t => t.Id == tableId ? t.WithName(name) : t).ToImmutableArray() }));
        return StructureEdits.Draft(snapshot, updated);
    }
}

public sealed class RenameTableColumnCommand(SheetId sheetId, int tableId, int columnId, string name) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var source = snapshot.GetSheet(sheetId); var table = source.Data.Tables.Single(t => t.Id == tableId); var column = table.Columns.Single(c => c.Id == columnId);
        if (column.Name == name) return new(snapshot, [], false);
        _ = table.WithColumns(table.Columns.Replace(column, column with { Name = name })); // Validate before transforming any formula.
        var updated = StructureEdits.TransformAt(snapshot, (node, context, position) =>
        {
            if (node is not StructuredReferenceNode r || StructuredReferences.Resolve(snapshot, context, position, r)?.Table.Id != tableId) return node;
            return r with { FirstColumn = string.Equals(r.FirstColumn, column.Name, StringComparison.OrdinalIgnoreCase) ? name : r.FirstColumn,
                LastColumn = string.Equals(r.LastColumn, column.Name, StringComparison.OrdinalIgnoreCase) ? name : r.LastColumn };
        }, (sheet, cells) =>
        {
            if (sheet.Id != sheetId) return sheet.WithCells(cells);
            var current = sheet.Data.Tables.Single(t => t.Id == tableId); var index = current.Columns.Select(c => c.Id).ToList().IndexOf(columnId);
            var renamed = current.WithColumns(current.Columns.SetItem(index, current.Columns[index] with { Name = name }));
            if (current.HeaderRow)
            {
                var position = new CellPosition(current.Range.Top, current.Range.Left + index);
                if (source.GetCell(position.Row, position.Column)?.Content.RichText is not null) throw new NotSupportedException(MessageText.Get("富文本表头改名需要先明确转换为普通文本。"));
                cells = cells.SetItem(position, new(CellContent.FromLiteral(CellValue.FromText(name), source.GetStyleId(position.Row, position.Column))));
            }
            return sheet.WithCells(cells).WithData(sheet.Data with { Tables = sheet.Data.Tables.Replace(current, renamed) });
        });
        return StructureEdits.Draft(snapshot, updated);
    }
}

/// <summary>Explicitly replace an entire calculated column, including existing exceptions, in one transaction.</summary>
public sealed class SetTableCalculatedColumnCommand(SheetId sheetId, int tableId, int columnId, string formula) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId); var table = sheet.Data.Tables.Single(t => t.Id == tableId);
        var index = table.Columns.Select(c => c.Id).ToList().IndexOf(columnId); if (index < 0) throw new ArgumentException(MessageText.Get("表列不存在。"));
        var first = new CellPosition(table.Range.Top + (table.HeaderRow ? 1 : 0), table.Range.Left + index);
        _ = FormulaBinding.Bind(FormulaSyntax.Parse(formula), snapshot, sheetId, first);
        var updated = table.WithColumns(table.Columns.SetItem(index, table.Columns[index] with { CalculatedFormula = formula })).WithInsertRow(false);
        var cells = sheet.Cells.ToBuilder();
        for (var row = first.Row; row <= table.Range.Bottom - (table.TotalsRow ? 1 : 0); row++)
        {
            if (sheet.GetCell(row, first.Column)?.Content.RichText is not null) throw new NotSupportedException(MessageText.Get("计算列包含富文本，请先明确转换为普通文本。"));
            cells[new(row, first.Column)] = new(CellContent.FromFormula(FormulaCopy.Translate(formula, first, new(row, first.Column)), sheet.GetStyleId(row, first.Column)));
        }
        var after = snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithCells(cells.ToImmutable()).WithData(s.Data with { Tables = s.Data.Tables.Replace(table, updated) }) : s));
        return StructureEdits.Draft(snapshot, after);
    }
}

internal static class TableStructure
{
    internal static WorksheetData NormalizeTemplates(WorksheetSnapshot before, WorksheetData after, SheetAxis axis,
        Func<int, int?> position, Func<FormulaNode, SheetId, CellPosition?, FormulaNode> transform)
    {
        return after with { Tables = after.Tables.Select(table =>
        {
            var original = before.Data.Tables.Single(t => t.Id == table.Id);
            return table.WithColumns(table.Columns.Select((column, index) =>
            {
                var old = original.Columns.FirstOrDefault(c => c.Id == column.Id);
                if (old?.CalculatedFormula is not { } formula) return column;
                var oldIndex = original.Columns.IndexOf(old); var first = new CellPosition(original.Range.Top + (original.HeaderRow ? 1 : 0), original.Range.Left + oldIndex);
                var source = first;
                if (axis == SheetAxis.Row)
                {
                    while (source.Row <= original.Range.Bottom - (original.TotalsRow ? 1 : 0) && position(source.Row) is null) source = source with { Row = source.Row + 1 };
                    if (source.Row > original.Range.Bottom - (original.TotalsRow ? 1 : 0)) return column with { CalculatedFormula = formula }; // Excel retains the reusable template after deleting every record.
                }
                var copied = FormulaCopy.Translate(formula, first, source);
                var expression = FormulaSyntax.Map(FormulaSyntax.Parse(copied), node => transform(node, before.Id, source));
                var mapped = axis == SheetAxis.Row ? source with { Row = position(source.Row)!.Value } : source with { Column = position(source.Column)!.Value };
                var next = new CellPosition(table.Range.Top + (table.HeaderRow ? 1 : 0), table.Range.Left + index);
                return column with { CalculatedFormula = FormulaCopy.Translate("=" + FormulaSyntax.Print(expression), mapped, next) };
            }));
        }).ToImmutableArray() };
    }

    internal static FormulaNode TransformReference(WorkbookSnapshot snapshot, WorksheetSnapshot target, SheetAxis axis,
        Func<int, int?> position, Func<int, int, (int Start, int End)?> interval, FormulaNode node, SheetId context, CellPosition? currentCell)
    {
        var reference = node as StructuredReferenceNode ?? (node is NameNode n ? new StructuredReferenceNode(n.Name, Sheet: n.Sheet) : null);
        if (reference is null || StructuredReferences.Resolve(snapshot, context, currentCell, reference) is not { } resolved || resolved.Sheet.Id != target.Id) return node;
        var table = resolved.Table; var range = table.Range;
        FormulaNode Missing() => new ValueNode(CellValue.FromError(CellErrorCode.Reference));
        if (interval(axis == SheetAxis.Row ? range.Top : range.Left, axis == SheetAxis.Row ? range.Bottom : range.Right) is null) return Missing();
        if (axis != SheetAxis.Column || reference.FirstColumn is null) return node;
        var names = table.Columns.Select(c => c.Name).ToList();
        var left = names.FindIndex(n => n.Equals(reference.FirstColumn, StringComparison.OrdinalIgnoreCase));
        var right = reference.LastColumn is null ? left : names.FindIndex(n => n.Equals(reference.LastColumn, StringComparison.OrdinalIgnoreCase));
        if (left < 0 || right < left) return node;
        if (position(range.Left + left) is null || position(range.Left + right) is null) return Missing();
        var surviving = Enumerable.Range(left, right - left + 1).Where(c => position(range.Left + c).HasValue).ToArray();
        if (surviving.Length == 0) return Missing();
        return reference with { FirstColumn = names[surviving[0]], LastColumn = reference.LastColumn is null ? null : names[surviving[^1]] };
    }

    internal static WorkbookSnapshot FillInserted(WorkbookSnapshot snapshot, WorksheetSnapshot source, Func<CellPosition, bool> inserted)
    {
        var sheet = snapshot.GetSheet(source.Id); var cells = sheet.Cells.ToBuilder();
        foreach (var table in sheet.Data.Tables)
        {
            for (var i = 0; i < table.Columns.Length; i++)
            {
                var column = table.Columns[i]; var x = table.Range.Left + i;
                if (table.HeaderRow && inserted(new(table.Range.Top, x)))
                    cells[new(table.Range.Top, x)] = new(CellContent.FromLiteral(CellValue.FromText(column.Name), sheet.GetStyleId(table.Range.Top, x)));
                if (column.CalculatedFormula is not { } formula) continue;
                var first = new CellPosition(table.Range.Top + (table.HeaderRow ? 1 : 0), x);
                for (var row = first.Row; row <= table.Range.Bottom - (table.TotalsRow ? 1 : 0); row++)
                    if (inserted(new(row, x)) && (!cells.TryGetValue(new(row, x), out var cell) || cell.Content.Kind == CellContentKind.Literal && cell.Value.Kind == CellValueKind.Blank))
                        cells[new(row, x)] = new(CellContent.FromFormula(FormulaCopy.Translate(formula, first, new(row, x)), sheet.GetStyleId(row, x)));
            }
        }
        return snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheet.Id ? s.WithCells(cells.ToImmutable()) : s));
    }
}
