using QuickExcel.Domain;

namespace QuickExcel.Presentation;

/// <summary>Bounds whole-axis cell operations before enumeration; explicit rectangles keep their extent.</summary>
public static class SelectionPlanning
{
    public static bool IsWholeColumn(CellRectangle range) => range.Top == 0 && range.Bottom == WorkbookLimits.MaxRows - 1;
    public static bool IsWholeRow(CellRectangle range) => range.Left == 0 && range.Right == WorkbookLimits.MaxColumns - 1;

    public static CellRectangle BoundToUsedCells(WorksheetSnapshot sheet, CellRectangle range)
    {
        var columns = IsWholeColumn(range); var rows = IsWholeRow(range);
        if (!columns && !rows) return range;
        var bottom = 0; var right = 0;
        foreach (var position in sheet.Cells.Keys.Concat(sheet.Annotations.Links.Keys).Concat(sheet.Annotations.Notes.Keys))
        {
            if (!range.Contains(position.Row, position.Column)) continue;
            bottom = Math.Max(bottom, position.Row); right = Math.Max(right, position.Column);
        }
        return range with { Bottom = columns ? bottom : range.Bottom, Right = rows ? right : range.Right };
    }
}
