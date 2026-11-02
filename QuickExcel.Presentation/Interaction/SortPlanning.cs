using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

/// <summary>A resolved record rectangle plus whether its first row is a header that must stay in place.</summary>
public sealed record SortBlock(CellRectangle Range, bool HasHeader);

/// <summary>
/// WPF-free rule set shared by the ribbon, the header panel and the multi-level dialog. It answers
/// “which records does this sort move?” without touching the document or the grid.
/// </summary>
public static class SortPlanning
{
    public const int MaxCells = SortRangeCommand.MaxCells;

    /// <summary>
    /// A one-cell or whole-column selection cannot describe records, so the contiguous block around
    /// the active cell is used instead. A multi-row selection keeps its rows and grows sideways.
    /// </summary>
    public static SortBlock Resolve(WorksheetSnapshot sheet, CellRectangle? selection, CellPosition active)
    {
        var row = Math.Clamp(active.Row, 0, WorkbookLimits.MaxRows - 1);
        var column = Math.Clamp(active.Column, 0, WorkbookLimits.MaxColumns - 1);
        var block = selection is { } selected && IsRecordSelection(selected, row, column)
            ? SortRangeCommand.ExpandColumns(sheet, selected, hasHeader: false)
            : CurrentRegion(sheet, row, column);
        if (block.Count > MaxCells)
            throw new InvalidOperationException($"数据块为 {block.Rows:N0} 行 × {block.Columns:N0} 列，超过排序上限 {MaxCells:N0} 格；请选择较小的数据区域。");
        return new SortBlock(block, LooksLikeHeader(sheet, block));
    }

    /// <summary>Exact rows of an explicit, bounded multi-row selection that contains the active cell.</summary>
    public static bool IsRecordSelection(CellRectangle selection, int row, int column) =>
        selection.Contains(row, column) && selection.Rows >= 2 && selection.Count <= MaxCells;

    /// <summary>
    /// The filled block connected to the active cell. A row stays connected while any filled cell
    /// sits inside the current column span, so a blank cell inside a sparse column does not split
    /// the block; a completely blank row or column does.
    /// </summary>
    public static CellRectangle CurrentRegion(WorksheetSnapshot sheet, int row, int column)
    {
        if (!HasRecordCell(sheet.GetCell(row, column))) return new CellRectangle(row, column, row, column);
        var rowColumns = new Dictionary<int, HashSet<int>>();
        var columnRows = new Dictionary<int, HashSet<int>>();
        foreach (var (position, cell) in sheet.Cells)
        {
            if (!HasRecordCell(cell)) continue;
            if (!rowColumns.TryGetValue(position.Row, out var columns)) rowColumns[position.Row] = columns = [];
            columns.Add(position.Column);
            if (!columnRows.TryGetValue(position.Column, out var rows)) columnRows[position.Column] = rows = [];
            rows.Add(position.Row);
        }
        var top = row; var bottom = row; var left = column; var right = column;
        bool RowHasContent(int candidate) => rowColumns.TryGetValue(candidate, out var columns) && columns.Any(c => c >= left && c <= right);
        bool ColumnHasContent(int candidate) => columnRows.TryGetValue(candidate, out var rows) && rows.Any(r => r >= top && r <= bottom);
        bool extended;
        do
        {
            extended = false;
            while (top > 0 && RowHasContent(top - 1)) { top--; extended = true; }
            while (bottom < WorkbookLimits.MaxRows - 1 && RowHasContent(bottom + 1)) { bottom++; extended = true; }
            while (left > 0 && ColumnHasContent(left - 1)) { left--; extended = true; }
            while (right < WorkbookLimits.MaxColumns - 1 && ColumnHasContent(right + 1)) { right++; extended = true; }
            if ((long)(bottom - top + 1) * (right - left + 1) > MaxCells * 4L) break;
        }
        while (extended);
        return new CellRectangle(top, left, bottom, right);
    }

    /// <summary>
    /// Conservative Excel-like header rule: every non-blank header cell is text, and the first data
    /// row carries at least one typed value. All-text one-column blocks stay ambiguous and keep the
    /// first row unless the dialog explicitly marks it as a header.
    /// </summary>
    public static bool LooksLikeHeader(WorksheetSnapshot sheet, CellRectangle range)
    {
        if (range.Rows < 2 || range.Columns < 2) return false;
        var texts = 0;
        for (var column = range.Left; column <= range.Right; column++)
        {
            var value = sheet.GetCell(range.Top, column)?.Value ?? CellValue.Blank;
            if (value.Kind == CellValueKind.Blank) continue;
            if (value.Kind != CellValueKind.Text) return false;
            texts++;
        }
        if (texts < 2) return false;
        for (var column = range.Left; column <= range.Right; column++)
            if ((sheet.GetCell(range.Top + 1, column)?.Value ?? CellValue.Blank).Kind is
                CellValueKind.Number or CellValueKind.Boolean or CellValueKind.Error) return true;
        return false;
    }

    private static bool HasRecordCell(CellSnapshot? cell) =>
        cell is not null && (cell.Content.Kind == CellContentKind.Formula || cell.Value.Kind != CellValueKind.Blank);
}
