using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private object? TableApi(WorksheetSnapshot sheet, JsonElement data, string name, JsonElement[] args)
    {
        var table = data.TryGetProperty("tableId", out var id) ? sheet.Data.Tables.FirstOrDefault(t => t.Id == Int(id))
            ?? throw new ModelFailure("InvalidReference", "表句柄已失效。") : null;
        var type = name.Split('.')[0]; var member = name.Split('.')[1];
        if (type is "AutoFilter" or "Filter") return FilterApi(sheet, table, data, type, member, args);
        if (table is null) throw Invalid("缺少表身份。");
        if (type == "TableSort") return TableSortApi(sheet, table, member, args);
        var range = table.Range;
        if (type == "TableColumn")
        {
            var columnId = Int(data.GetProperty("columnId"));
            var index = table.Columns.Select(c => c.Id).ToList().IndexOf(columnId);
            if (index < 0) throw new ModelFailure("InvalidReference", "表列句柄已失效。");
            var column = table.Columns[index];
            if (member == "setName") { Arity(args, 1); Write(sheet.Id, range, new RenameTableColumnCommand(sheet.Id, table.Id, columnId, Text(args[0]))); return null; }
            Arity(args, 0); Spend(1, false);
            if (member == "getId") return column.Id;
            if (member == "getIndex") return index;
            if (member == "getName") return column.Name;
            if (member == "getFilter") return Handle("Filter", new { sheetId = sheet.Id.ToString(), tableId = table.Id, columnId });
            range = range with { Left = range.Left + index, Right = range.Left + index };
        }
        object Column(TableColumn c) => Handle("TableColumn", new { sheetId = sheet.Id.ToString(), tableId = table.Id, columnId = c.Id });
        switch (member)
        {
            case "getRange": case "getHeaderRowRange": case "getTotalRowRange": case "getRangeBetweenHeaderAndTotal":
                Arity(args, 0); Spend(1, false);
                if (member == "getHeaderRowRange") { if (!table.HeaderRow) throw new ModelFailure("ItemNotFound", "表没有表头。"); range = range with { Bottom = range.Top }; }
                if (member == "getTotalRowRange") { if (!table.TotalsRow) throw new ModelFailure("ItemNotFound", "表没有汇总行。"); range = range with { Top = range.Bottom }; }
                if (member == "getRangeBetweenHeaderAndTotal") range = range with { Top = range.Top + (table.HeaderRow ? 1 : 0), Bottom = range.Bottom - (table.TotalsRow ? 1 : 0) };
                if (range.Top > range.Bottom) throw new ModelFailure("ItemNotFound", "表没有数据行。");
                return RangeHandle(sheet.Id, range);
            case "getId": case "getLegacyId": case "getName": case "getRowCount": case "getShowHeaders": case "getShowTotals":
            case "getHighlightFirstColumn": case "getHighlightLastColumn": case "getShowBandedColumns": case "getShowBandedRows": case "getPredefinedTableStyle": case "getShowFilterButton":
                Arity(args, 0); Spend(1, false);
                return member switch
                {
                    "getId" or "getLegacyId" => table.Id.ToString(), "getName" => table.Name,
                    "getRowCount" => range.Bottom - range.Top + 1 - (table.HeaderRow ? 1 : 0) - (table.TotalsRow ? 1 : 0),
                    "getShowHeaders" => table.HeaderRow, "getShowTotals" => table.TotalsRow,
                    "getHighlightFirstColumn" => table.Style.FirstColumn, "getHighlightLastColumn" => table.Style.LastColumn,
                    "getShowBandedColumns" => table.Style.ColumnStripes, "getShowBandedRows" => table.Style.RowStripes,
                    "getPredefinedTableStyle" => table.Style.Name ?? "", _ => table.Filter is not null && table.Filter.Columns.Values.All(c => c.ShowButton)
                };
            case "getWorksheet": Arity(args, 0); return Handle("Worksheet", new { sheetId = sheet.Id.ToString() });
            case "getSort": case "getAutoFilter":
                Arity(args, 0); return Handle(member == "getSort" ? "TableSort" : "AutoFilter", new { sheetId = sheet.Id.ToString(), tableId = table.Id });
            case "getColumns": Arity(args, 0); Spend(1, false); return table.Columns.Select(Column).ToArray();
            case "getColumn": case "getColumnById": case "getColumnByName":
                Arity(args, 1); Spend(1, false);
                var found = args[0].ValueKind == JsonValueKind.Number && member != "getColumnByName"
                    ? table.Columns.FirstOrDefault(c => c.Id == Int(args[0])) : table.Columns.FirstOrDefault(c => c.Name.Equals(Text(args[0]), StringComparison.OrdinalIgnoreCase));
                return found is null ? null : Column(found);
            case "setName": Arity(args, 1); Write(sheet.Id, range, new RenameTableCommand(sheet.Id, table.Id, Text(args[0]))); return null;
            case "setPredefinedTableStyle": case "setHighlightFirstColumn": case "setHighlightLastColumn": case "setShowBandedColumns": case "setShowBandedRows":
                Arity(args, 1);
                var style = member switch
                {
                    "setPredefinedTableStyle" => TableStyles.Resolve(Model.View, Text(args[0]), table.Style.RowStripes, table.Style.ColumnStripes, table.Style.FirstColumn, table.Style.LastColumn),
                    "setHighlightFirstColumn" => table.Style with { FirstColumn = Bool(args[0]) },
                    "setHighlightLastColumn" => table.Style with { LastColumn = Bool(args[0]) },
                    "setShowBandedColumns" => table.Style with { ColumnStripes = Bool(args[0]) },
                    _ => table.Style with { RowStripes = Bool(args[0]) }
                };
                Write(sheet.Id, range, new SetScriptTableStyleCommand(sheet.Id, table.Id, style)); return null;
            case "resize":
                Arity(args, 1); var next = Address(args[0], sheet.Id);
                if (next.Top != range.Top || next.Left != range.Left || next.Right != range.Right) throw Unsupported("表 resize 当前只支持调整底边。");
                Write(sheet.Id, range with { Bottom = Math.Max(range.Bottom, next.Bottom) }, new ResizeTableCommand(sheet.Id, table.Id, next.Bottom)); return null;
            case "addRow": case "addRows":
                Arity(args, 0, 2);
                if (table.TotalsRow) throw Unsupported("含汇总行的表追加需要移动汇总格，尚未支持。");
                var rowCount = range.Rows - (table.HeaderRow ? 1 : 0);
                var at = args.Length == 0 || args[0].ValueKind == JsonValueKind.Null ? -1 : Int(args[0]);
                if (at != -1 && at != rowCount) throw Unsupported("当前支持在表尾追加行；中间插入尚未支持。");
                CellValue[][] rows;
                if (args.Length < 2 || args[1].ValueKind == JsonValueKind.Null) rows = [Enumerable.Repeat(CellValue.Blank, range.Columns).ToArray()];
                else if (member == "addRow")
                {
                    if (args[1].ValueKind != JsonValueKind.Array || args[1].GetArrayLength() != range.Columns) throw Invalid("追加行宽度必须等于表列数。");
                    rows = [args[1].EnumerateArray().Select(ReadScalar).ToArray()];
                }
                else
                {
                    if (args[1].ValueKind != JsonValueKind.Array) throw Invalid("追加行需要二维数组。");
                    if ((long)args[1].GetArrayLength() * range.Columns > MaxWriteCells) throw new ModelFailure("BudgetExceeded", "追加超过 20000 格。");
                    rows = Matrix(args[1], args[1].GetArrayLength(), range.Columns).Select(row => row.Select(ReadScalar).ToArray()).ToArray();
                }
                if (rows.Length == 0) return null;
                var appended = new CellRectangle(range.Bottom + 1, range.Left, checked(range.Bottom + rows.Length), range.Right);
                WorkbookLimits.ValidatePosition(appended.Bottom, appended.Right); Spend(checked((int)appended.Count), false);
                foreach (var address in Cells(sheet.Id, appended))
                    if (sheet.GetCell(address.Row, address.Column) is { } existing && (existing.Content.Kind == CellContentKind.Formula || existing.Value.Kind != CellValueKind.Blank) ||
                        sheet.Annotations.Links.ContainsKey(new(address.Row, address.Column)) || sheet.Annotations.Notes.ContainsKey(new(address.Row, address.Column)))
                        throw Unsupported("表尾目标区域含数据或附件，不能隐式覆盖；请先整理目标区域。");
                Write(sheet.Id, range with { Bottom = appended.Bottom }, new ResizeTableCommand(sheet.Id, table.Id, appended.Bottom));
                if (args.Length == 2 && args[1].ValueKind != JsonValueKind.Null)
                { Spend((int)appended.Count, true); Model.SetValues(sheet.Id, appended, rows.Select(row => (IReadOnlyList<CellValue>)row).ToArray()); }
                return null;
            case "clearFilters": case "reapplyFilters":
                Arity(args, 0); return FilterApi(sheet, table, data, "AutoFilter", member == "clearFilters" ? "clearCriteria" : "reapply", args);
            case "convertToRange": case "delete":
                Arity(args, 0); Write(sheet.Id, range, new RemoveScriptTableCommand(sheet.Id, table.Id, member == "delete"));
                return member == "convertToRange" ? RangeHandle(sheet.Id, range) : null;
            default: throw Unsupported(name);
        }
    }
    private object? TableSortApi(WorksheetSnapshot sheet, WorksheetTable table, string member, JsonElement[] args)
    {
        if (member == "apply") { ApplySort(sheet, table.Range with { Bottom = table.Range.Bottom - (table.TotalsRow ? 1 : 0) }, args, true, table.HeaderRow); return null; }
        Arity(args, 0);
        if (member == "clear") { Write(sheet.Id, table.Range, new ClearSortStateCommand(sheet.Id, table.Id)); return null; }
        if (member == "reapply")
        {
            if (table.Sort is { } sort) Write(sheet.Id, table.Range, new SortRangeCommand(new(sheet.Id, new(sort.Range.Top, sort.Range.Left), new(sort.Range.Bottom, sort.Range.Right), sort.Keys)));
            return null;
        }
        Spend(1, false);
        return member switch
        {
            "getFields" => table.Sort?.Keys.Select(k => new { key = k.Column - table.Range.Left, ascending = !k.Descending, sortOn = (int)k.Basis, color = Rgb(k.PreferredColor) }).ToArray() ?? [],
            "getMatchCase" => table.Sort?.Keys.FirstOrDefault()?.CaseSensitive ?? false,
            "getMethod" => table.Sort?.SortMethod == "stroke" ? 1 : 0, _ => throw Unsupported(member)
        };
    }
    private void ApplySort(WorksheetSnapshot sheet, CellRectangle range, JsonElement[] args, bool table, bool tableHeader = true)
    {
        Arity(args, 1, table ? 3 : 5);
        var matchCase = args.Length > 1 && args[1].ValueKind != JsonValueKind.Null && Bool(args[1]);
        var header = table ? tableHeader : args.Length > 2 && args[2].ValueKind != JsonValueKind.Null && Bool(args[2]);
        if (!table && args.Length > 3 && args[3].ValueKind != JsonValueKind.Null && Int(args[3]) != 0) throw Unsupported("只支持按行排序。");
        var methodIndex = table ? 2 : 4;
        if (args.Length > methodIndex && args[methodIndex].ValueKind != JsonValueKind.Null) throw Unsupported("不提供拼音或笔画排序；省略 method 使用序数文本排序。");
        if (args[0].ValueKind != JsonValueKind.Array || args[0].GetArrayLength() is < 1 or > 32) throw Invalid("排序需要 1–32 个键。");
        var keys = args[0].EnumerateArray().Select(field =>
        {
            if (field.ValueKind != JsonValueKind.Object || field.EnumerateObject().Any(p => p.Name is not ("key" or "ascending" or "sortOn" or "color" or "dataOption"))) throw Unsupported("排序键包含尚未支持的选项。");
            if (field.TryGetProperty("dataOption", out var option) && Int(option) != 0) throw Unsupported("textAsNumber 排序尚未实现。");
            var basis = field.TryGetProperty("sortOn", out var on) ? Int(on) : 0;
            if (basis is < 0 or > 2) throw Unsupported("图标排序尚未实现。");
            var index = Int(field.GetProperty("key")); if (index < 0 || index > range.Right - range.Left) throw Invalid("排序键超出范围。");
            return new SortKey(range.Left + index, field.TryGetProperty("ascending", out var ascending) && !Bool(ascending), matchCase, (SortBasis)basis,
                basis == 0 ? null : Argb(field.GetProperty("color")));
        }).ToArray();
        Spend(checked((int)range.Count), false);
        Write(sheet.Id, range, new SortRangeCommand(new(sheet.Id, new(range.Top, range.Left), new(range.Bottom, range.Right), keys, header)));
    }
}
