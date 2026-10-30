using QuickExcel.Domain;

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
        return new(filter, sort);
    }
}
