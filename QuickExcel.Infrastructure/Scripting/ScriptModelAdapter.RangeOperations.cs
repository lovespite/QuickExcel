using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private object? RangeObjectApi(WorksheetSnapshot sheet, CellRectangle range, string member, JsonElement[] args)
    {
        switch (member)
        {
            case "getUsedRange":
            {
                Arity(args, 0, 1);
                var valuesOnly = args.Length == 1 && Bool(args[0]);
                if (sheet.Cells.Count > MaxReadCells)
                    throw new ModelFailure("BudgetExceeded", "已用区域扫描超过读取预算。");
                var positions = sheet.Cells
                    .Where(p => range.Contains(p.Key.Row, p.Key.Column) &&
                        (!valuesOnly || p.Value.Content.Kind == CellContentKind.Formula || p.Value.Content.LiteralValue.Kind != CellValueKind.Blank))
                    .Select(p => p.Key).ToArray();
                Spend(Math.Max(1, positions.Length), false);
                return positions.Length == 0 ? null : RangeHandle(sheet.Id, new(
                    positions.Min(p => p.Row), positions.Min(p => p.Column),
                    positions.Max(p => p.Row), positions.Max(p => p.Column)));
            }
            case "getMergedAreas":
            {
                Arity(args, 0);
                var merges = sheet.MergedCells.Select(CellRectangle.Parse).Where(item => Overlap(range, item)).ToArray();
                Spend(Math.Max(1, merges.Length), false);
                return Handle("RangeAreas", new Dictionary<string, object?> { ["sheetId"] = sheet.Id.ToString(), ["addresses"] = merges.Select(item => item.ToString()).ToArray() });
            }
            case "getTables":
            {
                Arity(args, 0, 1);
                var fullyContained = args.Length == 1 && Bool(args[0]);
                var tables = sheet.Data.Tables.Where(table => fullyContained
                    ? range.Top <= table.Range.Top && range.Left <= table.Range.Left &&
                      range.Bottom >= table.Range.Bottom && range.Right >= table.Range.Right
                    : Overlap(range, table.Range)).ToArray();
                Spend(1, false);
                return tables.Select(table => TableHandle(sheet.Id, table.Id)).ToArray();
            }
            case "getRowHidden":
            {
                Arity(args, 0);
                Spend(1, false);
                return Enumerable.Range(range.Top, range.Rows).All(row => sheet.Rows.GetValueOrDefault(row)?.Hidden == true);
            }
            case "setRowHidden":
                Arity(args, 1);
                SetAxisHidden(sheet, SheetAxis.Row, range.Top, range.Rows, Bool(args[0]));
                return null;
            case "getColumnHidden":
            {
                Arity(args, 0);
                Spend(1, false);
                return Enumerable.Range(range.Left, range.Columns).All(column => sheet.Columns.GetValueOrDefault(column)?.Hidden == true);
            }
            case "setColumnHidden":
                Arity(args, 1);
                SetAxisHidden(sheet, SheetAxis.Column, range.Left, range.Columns, Bool(args[0]));
                return null;
            default:
                throw Unsupported(member);
        }
    }

    private void SetAxisHidden(WorksheetSnapshot sheet, SheetAxis axis, int first, int count, bool hidden)
    {
        if (count > MaxWriteCells) throw new ModelFailure("BudgetExceeded", "单次最多修改 20000 行/列。");
        ExecuteObject(sheet.Id, new SetAxisMetadataCommand(sheet.Id, axis, first, count, hidden: hidden), Math.Max(1, count));
    }

    private object? RangeStructureApi(WorksheetSnapshot sheet, CellRectangle range, string member, JsonElement[] args)
    {
        switch (member)
        {
            case "merge":
            {
                Arity(args, 0, 1);
                var across = args.Length == 1 && Bool(args[0]);
                if (range.Count == 1) throw Invalid("请选择多个单元格。");
                if (!across)
                {
                    Write(sheet.Id, range, new MergeCellsCommand(sheet.Id, range));
                    return null;
                }
                for (var row = range.Top; row <= range.Bottom; row++)
                    Write(sheet.Id, range with { Top = row, Bottom = row }, new MergeCellsCommand(sheet.Id, range with { Top = row, Bottom = row }));
                return null;
            }
            case "unmerge":
                Arity(args, 0);
                Write(sheet.Id, range, new MergeCellsCommand(sheet.Id, range, unmerge: true));
                return null;
            case "insert":
            {
                Arity(args, 1);
                var direction = EnumName("InsertShiftDirection", args[0]);
                if (direction == "down")
                {
                    RequireEntireRows(range, "insert(down)");
                    InsertRows(sheet, range.Top, range.Rows);
                    return RangeHandle(sheet.Id, range);
                }
                RequireEntireColumns(range, "insert(right)");
                InsertColumns(sheet, range.Left, range.Columns);
                return RangeHandle(sheet.Id, range);
            }
            case "delete":
            {
                Arity(args, 1);
                var direction = EnumName("DeleteShiftDirection", args[0]);
                if (direction == "up")
                {
                    RequireEntireRows(range, "delete(up)");
                    DeleteAxis(sheet, SheetAxis.Row, range.Top, range.Rows);
                    return null;
                }
                RequireEntireColumns(range, "delete(left)");
                DeleteAxis(sheet, SheetAxis.Column, range.Left, range.Columns);
                return null;
            }
            default:
                throw Unsupported(member);
        }
    }

    private static void RequireEntireRows(CellRectangle range, string operation)
    {
        if (range.Left != 0 || range.Right != WorkbookLimits.MaxColumns - 1)
            throw Unsupported(operation + " 目前只支持整行区域；请先调用 getEntireRow()。");
    }

    private static void RequireEntireColumns(CellRectangle range, string operation)
    {
        if (range.Top != 0 || range.Bottom != WorkbookLimits.MaxRows - 1)
            throw Unsupported(operation + " 目前只支持整列区域；请先调用 getEntireColumn()。");
    }

    private void InsertRows(WorksheetSnapshot sheet, int first, int count)
    {
        if (count >= WorkbookLimits.MaxRows) throw Invalid("插入的行数超出工作表范围。");
        ExecuteObject(sheet.Id, new ShiftAxisCommand(sheet.Id, SheetAxis.Row, first, count), Math.Min(MaxWriteCells, Math.Max(1, count)));
    }

    private void InsertColumns(WorksheetSnapshot sheet, int first, int count)
    {
        if (count >= WorkbookLimits.MaxColumns) throw Invalid("插入的列数超出工作表范围。");
        ExecuteObject(sheet.Id, new ShiftAxisCommand(sheet.Id, SheetAxis.Column, first, count), Math.Min(MaxWriteCells, Math.Max(1, count)));
    }

    private void DeleteAxis(WorksheetSnapshot sheet, SheetAxis axis, int first, int count)
    {
        var limit = axis == SheetAxis.Row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
        if (count >= limit) throw Invalid("不能删除整个工作表的全部行列。");
        ExecuteObject(sheet.Id, new ShiftAxisCommand(sheet.Id, axis, first, count, delete: true), Math.Min(MaxWriteCells, Math.Max(1, count)));
    }
}
