using System.Collections.Immutable;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Runtime;

internal sealed partial class ScriptModelAdapter
{
    private object? FilterApi(WorksheetSnapshot sheet, WorksheetTable? table, JsonElement data, string type, string member, JsonElement[] args)
    {
        var filter = table is null ? sheet.Data.Filter : table.Filter;
        var range = filter?.Range ?? (table is null ? (CellRectangle?)null : table.Range with { Bottom = table.Range.Bottom - (table.TotalsRow ? 1 : 0) });
        int? column = null;
        if (type == "Filter")
        {
            if (table is null) throw Invalid("列筛选缺少表身份。");
            var columnId = Int(data.GetProperty("columnId")); var index = table.Columns.Select(c => c.Id).ToList().IndexOf(columnId);
            if (index < 0) throw new ModelFailure("InvalidReference", "表列句柄已失效。");
            column = table.Range.Left + index;
        }
        void Commit(WorksheetAutoFilter? next)
        {
            var affected = next?.Range ?? range ?? ScalarRange;
            Spend(checked((int)affected.Count), false);
            Write(sheet.Id, affected, table is null ? new SetAutoFilterCommand(sheet.Id, next) : new SetTableFilterCommand(sheet.Id, table.Id, next));
        }
        if (member.StartsWith("get", StringComparison.Ordinal))
        {
            Arity(args, 0); Spend(1, false);
            return member switch
            {
                "getEnabled" => filter is not null,
                "getIsDataFiltered" => filter?.Columns.Values.Any(c => c.HasCriteria) == true,
                "getRange" => filter is null ? null : RangeHandle(sheet.Id, filter.Range),
                "getCriteria" when column is { } c => Criteria(filter?.Columns.GetValueOrDefault(c)),
                "getCriteria" => filter is null ? Array.Empty<object>() : Enumerable.Range(filter.Range.Left, filter.Range.Right - filter.Range.Left + 1).Select(c => Criteria(filter.Columns.GetValueOrDefault(c))).ToArray(),
                _ => throw Unsupported(member)
            };
        }
        if (member == "remove") { Arity(args, 0); Commit(null); return null; }
        if (member == "reapply")
        {
            Arity(args, 0); if (filter is not null)
            {
                if (filter.Range.Count > MaxReadCells) throw new ModelFailure("BudgetExceeded", "筛选重应用超过预算。");
                Spend((int)filter.Range.Count, false); Model.PrepareValueRead(sheet.Id, filter.Range, cancellation);
                _ = WorksheetFilterEngine.Evaluate(Model.ValueView, sheet.Id, cancellationToken: cancellation);
            }
            return null;
        }
        if (member is "clearCriteria" or "clear" or "clearColumnCriteria")
        {
            Arity(args, member == "clearColumnCriteria" ? 1 : 0);
            if (filter is null) return null;
            if (member == "clearColumnCriteria") column = filter.Range.Left + Int(args[0]);
            if (column is { } c && (c < filter.Range.Left || c > filter.Range.Right)) throw Invalid("筛选列越界。");
            Commit(new(filter.Range, column is { } selected ? filter.Columns.Remove(selected) : [])); return null;
        }
        AutoFilterColumn? criteria;
        if (type == "AutoFilter")
        {
            Arity(args, 1, 3); var address = Address(args[0], sheet.Id);
            if (table is not null && address != range) throw Invalid("表筛选范围必须等于表头及数据区。");
            if (filter is not null && filter.Range != address) filter = null;
            range = address;
            if (args.Length == 1) { Commit(new(address, filter?.Columns)); return null; }
            if (args.Length != 3) throw Invalid("设置列筛选需要 columnIndex 和 criteria。");
            column = address.Left + Int(args[1]); criteria = ReadCriteria(args[2]);
        }
        else
        {
            if (member == "apply") { Arity(args, 1); criteria = ReadCriteria(args[0]); }
            else if (member == "applyValuesFilter") { Arity(args, 1); criteria = ValuesCriteria(args[0]); }
            else if (member == "applyCustomFilter")
            {
                Arity(args, 1, 3); var first = Condition(Text(args[0]));
                var second = args.Length > 1 && args[1].ValueKind != JsonValueKind.Null ? Condition(Text(args[1])) : null;
                var op = args.Length > 2 && args[2].ValueKind != JsonValueKind.Null ? Int(args[2]) : 0;
                if (op is < 0 or > 1) throw Invalid("筛选运算符无效。");
                criteria = new(first: first, second: second, useAnd: op == 0);
            }
            else throw Unsupported(member);
        }
        if (range is not { } target || column is not { } selectedColumn || selectedColumn < target.Left || selectedColumn > target.Right) throw Invalid("筛选列越界。");
        var columns = filter?.Columns ?? ImmutableDictionary<int, AutoFilterColumn>.Empty;
        Commit(new(target, criteria is null ? columns.Remove(selectedColumn) : columns.SetItem(selectedColumn, criteria))); return null;
    }
    private static AutoFilterColumn? ReadCriteria(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Object) throw Invalid("筛选条件需要对象。");
        if (!value.EnumerateObject().Any()) return null;
        if (value.EnumerateObject().Any(p => p.Name is not ("filterOn" or "values" or "criterion1" or "criterion2" or "operator"))) throw Unsupported("筛选条件包含未支持的选项。");
        var on = Int(value.GetProperty("filterOn"));
        if (on == 5) return ValuesCriteria(value.GetProperty("values"));
        if (on != 9) throw Unsupported("仅支持值列表和自定义条件筛选。");
        var first = Condition(Text(value.GetProperty("criterion1")));
        var second = value.TryGetProperty("criterion2", out var other) ? Condition(Text(other)) : null;
        var op = value.TryGetProperty("operator", out var oper) ? Int(oper) : 0;
        if (op is < 0 or > 1) throw Invalid("筛选运算符无效。");
        return new(first: first, second: second, useAnd: op == 0);
    }
    private static AutoFilterColumn ValuesCriteria(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Array || value.GetArrayLength() > 100_000) throw Invalid("筛选值需要有界数组。");
        if (value.EnumerateArray().Any(v => v.ValueKind != JsonValueKind.String)) throw Unsupported("日期分组条件尚未实现；请使用显示文本。");
        var labels = value.EnumerateArray().Select(Text).ToArray();
        return new(labels.Where(s => s.Length != 0), labels.Contains(""));
    }
    private static AutoFilterCondition Condition(string text)
    {
        foreach (var (prefix, comparison) in new[] { (">=", AutoFilterComparison.GreaterThanOrEqual), ("<=", AutoFilterComparison.LessThanOrEqual),
            ("<>", AutoFilterComparison.NotEqual), (">", AutoFilterComparison.GreaterThan), ("<", AutoFilterComparison.LessThan), ("=", AutoFilterComparison.Equal) })
            if (text.StartsWith(prefix, StringComparison.Ordinal)) return new(comparison, text[prefix.Length..]);
        return new(AutoFilterComparison.Equal, text);
    }
    private static object Criteria(AutoFilterColumn? column)
    {
        if (column?.Color is not null) throw Unsupported(MessageText.Get("颜色筛选条件尚未暴露给脚本 API。请通过桌面筛选入口编辑。"));
        if (column?.Top is not null) throw Unsupported(MessageText.Get("前后项筛选条件尚未暴露给脚本 API。请通过桌面筛选入口编辑。"));
        if (column?.Values is { } values) return new { filterOn = 5, values = column.IncludeBlank ? values.Add("").ToArray() : values.ToArray() };
        if (column?.First is not { } first) return new { };
        static string Expression(AutoFilterCondition c) => (c.Comparison switch { AutoFilterComparison.Equal => "=", AutoFilterComparison.NotEqual => "<>",
            AutoFilterComparison.LessThan => "<", AutoFilterComparison.LessThanOrEqual => "<=", AutoFilterComparison.GreaterThan => ">", _ => ">=" }) + c.Operand;
        var result = new Dictionary<string, object> { ["filterOn"] = 9, ["criterion1"] = Expression(first), ["operator"] = column.UseAnd ? 0 : 1 };
        if (column.Second is { } second) result["criterion2"] = Expression(second);
        return result;
    }
}
