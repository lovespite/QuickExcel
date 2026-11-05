using QuickExcel.Domain;
using QuickExcel.GridGeometry;

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    private Guid _scrollSession;
    private readonly Dictionary<SheetId, GridScrollExtent> _scrollExtents = new();
    private GridScrollExtent _scrollExtent = new();
    private CellPosition? _pendingScrollCell;

    private void BindScrollExtent(Guid session)
    {
        if (_scrollSession != session) { _scrollExtents.Clear(); _scrollSession = session; }
        if (!_scrollExtents.TryGetValue(Sheet.Id, out var extent))
            _scrollExtents.Add(Sheet.Id, extent = new());
        _scrollExtent = extent;
    }

    private void IncludeUsedExtent()
    {
        _scrollExtent.Include(Sheet.UsedRowCount - 1, Sheet.UsedColumnCount - 1);
        foreach (var merge in _merges) _scrollExtent.Include(merge.Bottom, merge.Right);
        foreach (var table in Sheet.Data.Tables) _scrollExtent.Include(table.Range.Bottom, table.Range.Right);
        foreach (var comment in Sheet.Data.ThreadComments) _scrollExtent.Include(comment.Position.Row, comment.Position.Column);
        foreach (var anchor in Sheet.Images.Select(i => i.Anchor).Concat(Sheet.Data.Charts.Select(c => c.Anchor)))
        {
            var bounds = DrawingGeometry.RotatedBounds(anchor.Bounds(Sheet), anchor.Rotation);
            _scrollExtent.Include(DrawingEnd(bounds.Y + bounds.Height, true), DrawingEnd(bounds.X + bounds.Width, false));
        }
        // Drawing endpoints may be hidden, zero-sized or outside the sheet. Find
        // the last reachable cell without requiring a valid new drawing anchor.
        int DrawingEnd(double end, bool row)
        {
            var low = 0; var high = (row ? _rows.Count : _columns.Count) - 1;
            while (low < high)
            {
                var mid = low + (high - low) / 2;
                var start = row ? DrawingGeometry.RowStart(Sheet, mid + 1) : DrawingGeometry.ColumnStart(Sheet, mid + 1);
                if (start < end) low = mid + 1; else high = mid;
            }
            return low;
        }
    }

    private void RestoreScrollViewport()
    {
        if (_pendingScrollCell is not { } stored || ViewWidth <= 0 || ViewHeight <= 0) return;
        _pendingScrollCell = null;
        if (_rows.Cell(stored.Row) is { Frozen: false } row) _rows.ScrollTo(_rows.Offset + row.Start - _rows.FrozenSize, ViewHeight);
        if (_columns.Cell(stored.Column) is { Frozen: false } col) _columns.ScrollTo(_columns.Offset + col.Start - _columns.FrozenSize, ViewWidth);
        // Only restoring a saved view may include its visible cells. Scroll input
        // never does this, so reaching the boundary cannot continually grow it.
        if (_rows.Offset > 0) _scrollExtent.Include(_rows.Visible(ViewHeight).LastOrDefault().ModelIndex, 0);
        if (_columns.Offset > 0) _scrollExtent.Include(0, _columns.Visible(ViewWidth).LastOrDefault().ModelIndex);
    }

    private void IncludeSelectionExtent()
    {
        _scrollExtent.Include(_current.Row, _current.Column);
        foreach (var range in _selection.Ranges) _scrollExtent.IncludeSelection(range, _rows.Count, _columns.Count);
        UpdateScrollbars();
    }

    private void Scroll(double x, double y)
    {
        _rows.ScrollTo(y, ViewHeight, _scrollExtent.LastRow);
        _columns.ScrollTo(x, ViewWidth, _scrollExtent.LastColumn);
        UpdateScrollbars(); InvalidateContent(); PositionEditor(); ViewportChanged?.Invoke();
    }

    private void UpdateScrollbars()
    {
        _scrolling = true;
        try
        {
            _rows.ScrollTo(_rows.Offset, ViewHeight, _scrollExtent.LastRow);
            _columns.ScrollTo(_columns.Offset, ViewWidth, _scrollExtent.LastColumn);
            _vertical.Maximum = _rows.MaximumOffset(ViewHeight, _scrollExtent.LastRow);
            _vertical.ViewportSize = Math.Max(0, ViewHeight - _rows.FrozenSize); _vertical.Value = _rows.Offset;
            _horizontal.Maximum = _columns.MaximumOffset(ViewWidth, _scrollExtent.LastColumn);
            _horizontal.ViewportSize = Math.Max(0, ViewWidth - _columns.FrozenSize); _horizontal.Value = _columns.Offset;
        }
        finally { _scrolling = false; }
    }
}
