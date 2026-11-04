using QuickExcel.Domain;
using System.Collections.Immutable;

namespace QuickExcel.Core;

internal static class DataStructure
{
    internal static WorksheetData Axis(WorksheetData data, SheetAxis axis, Func<int, int?> position,
        Func<int, int, (int Start, int End)?> interval)
    {
        CellRectangle? Range(CellRectangle range)
        {
            var mapped = interval(axis == SheetAxis.Row ? range.Top : range.Left, axis == SheetAxis.Row ? range.Bottom : range.Right);
            if (mapped is null) return null;
            var result = axis == SheetAxis.Row ? range with { Top = mapped.Value.Start, Bottom = mapped.Value.End } : range with { Left = mapped.Value.Start, Right = mapped.Value.End };
            WorkbookLimits.ValidatePosition(result.Top, result.Left); WorkbookLimits.ValidatePosition(result.Bottom, result.Right);
            return result;
        }
        WorksheetAutoFilter? filter = null; WorksheetSortState? sort = null;
        if (data.Filter is { } oldFilter && Range(oldFilter.Range) is { } filterRange)
        {
            if (axis == SheetAxis.Row && position(oldFilter.Range.Top) != filterRange.Top)
                throw new NotSupportedException("操作会删除或移动筛选表头，请先清除筛选；操作未应用。");
            filter = new(filterRange, oldFilter.Columns.Select(p => new KeyValuePair<int?, AutoFilterColumn>(axis == SheetAxis.Column ? position(p.Key) : p.Key, p.Value))
                .Where(p => p.Key.HasValue).Select(p => new KeyValuePair<int, AutoFilterColumn>(p.Key!.Value, p.Value)));
        }
        if (data.Sort is { } oldSort && Range(oldSort.Range) is { } sortRange)
        {
            var keys = oldSort.Keys.Select(k => (Key: k, Column: axis == SheetAxis.Column ? position(k.Column) : k.Column))
                .Where(p => p.Column.HasValue).Select(p => new SortKey(p.Column!.Value, p.Key.Descending, p.Key.CaseSensitive, p.Key.Basis, p.Key.PreferredColor, p.Key.CustomOrder)).ToArray();
            if (keys.Length > 0) sort = new(sortRange, keys, oldSort.SortMethod);
        }
        var tables = new List<WorksheetTable>();
        foreach (var table in data.Tables)
        {
            if (Range(table.Range) is not { } range) continue;
            if (axis == SheetAxis.Row && (table.HeaderRow && position(table.Range.Top) != range.Top || table.TotalsRow && position(table.Range.Bottom) != range.Bottom))
                throw new NotSupportedException("插删不能移除表头或汇总行，请先调整表定义。");
            if (axis == SheetAxis.Row && table.TotalsRow && range.Bottom - range.Top < (table.HeaderRow ? 2 : 1))
                throw new NotSupportedException("含汇总行的表须至少保留一行数据；本次删除未应用。");
            var empty = axis == SheetAxis.Row && table.HeaderRow && !table.TotalsRow && range.Top == range.Bottom;
            if (empty) range = range with { Bottom = range.Bottom + 1 }; // Excel keeps an empty insertion row.
            var columns = table.Columns;
            if (axis == SheetAxis.Column)
            {
                var mapped = table.Columns.Select((c, i) => (Column: c, Position: position(table.Range.Left + i))).Where(p => p.Position.HasValue).ToDictionary(p => p.Position!.Value, p => p.Column);
                var names = table.Columns.Select(c => c.Name).ToHashSet(StringComparer.OrdinalIgnoreCase); var id = table.Columns.Max(c => c.Id);
                var next = new List<TableColumn>();
                for (var c = range.Left; c <= range.Right; c++)
                {
                    if (mapped.TryGetValue(c, out var existing)) { next.Add(existing); continue; }
                    var n = 1; while (names.Contains("列" + n)) n++; var name = "列" + n; names.Add(name); next.Add(new(checked(++id), name));
                }
                columns = next.ToImmutableArray();
            }
            var states = Axis(new(table.Filter, table.Sort), axis, position, interval);
            var tableFilter = empty && states.Filter is { } remaining ? new WorksheetAutoFilter(range, remaining.Columns) : states.Filter;
            tables.Add(new(table.Id, table.Name, range, columns, table.HeaderRow, table.TotalsRow, table.TotalsRowShown, table.Style, tableFilter, empty ? null : states.Sort,
                empty || table.InsertRow && range.Bottom - range.Top == table.Range.Bottom - table.Range.Top));
        }
        var conditional = data.ConditionalFormats.Select(r => (Rule: r, Ranges: r.Ranges.Select(Range).OfType<CellRectangle>().ToArray()))
            .Where(p => p.Ranges.Length > 0).Select(p => p.Rule.WithRanges(p.Ranges)).ToImmutableArray();
        return data with { Filter = filter, Sort = sort, Tables = tables.ToImmutableArray(), ConditionalFormats = conditional };
    }
}
