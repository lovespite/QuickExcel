// QuickExcel selection/hover projection. No model writes or cell enumeration.
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace FastWpfGrid;

partial class FastGridControl
{
    private bool _useSelectionOverlay;

    /// <summary>Opt-in translucent selection with original text colors, drawn separately from cell content.</summary>
    public bool UseSelectionOverlay
    {
        get => _useSelectionOverlay;
        set
        {
            if (_useSelectionOverlay == value) return;
            _useSelectionOverlay = value;
            if (!value) _scrollOffsetX = _scrollOffsetY = 0;
            if (_dragRenderingSubscribed || _dragTimer?.IsEnabled == true) StartDragScrolling();
            AdjustScrollbars();
            InvalidateAll();
        }
    }

    private void InvalidateSelection()
    {
        if (UseSelectionOverlay) selectionOverlay.Refresh();
        else InvalidateAll();
    }

    // Frozen and scrolling areas have different coordinate origins. Clip each
    // axis independently before taking the Cartesian product (at most four rects).
    private IEnumerable<(int First, int Last, double Min, double Max)> VisibleAxis(
        SeriesSizes sizes, int firstScroll, int visible, int header, int extent)
    {
        if (sizes.FrozenCount > 0 && header < extent)
            yield return (0, sizes.FrozenCount - 1, header, Math.Min(header + sizes.FrozenSize, extent));
        if (visible > 0 && header + sizes.FrozenSize < extent)
            yield return (firstScroll + sizes.FrozenCount, firstScroll + sizes.FrozenCount + visible - 1,
                header + sizes.FrozenSize, extent);
    }

    private IEnumerable<Rect> GetOverlayRectangles(GridCellRange range)
    {
        if (_drawBuffer == null) yield break;
        foreach (var rows in VisibleAxis(_rowSizes, FirstVisibleRowScrollIndex, VisibleRowCount, HeaderHeight, _drawBuffer.PixelHeight))
        foreach (var columns in VisibleAxis(_columnSizes, FirstVisibleColumnScrollIndex, VisibleColumnCount, HeaderWidth, _drawBuffer.PixelWidth))
        {
            var top = Math.Max(range.Top, rows.First); var bottom = Math.Min(range.Bottom, rows.Last);
            var left = Math.Max(range.Left, columns.First); var right = Math.Min(range.Right, columns.Last);
            if (top > bottom || left > right) continue;
            var x1 = Math.Max(GetColumnLeft(left), columns.Min);
            var y1 = Math.Max(GetRowTop(top), rows.Min);
            var x2 = Math.Min(GetColumnLeft(right) + _columnSizes.GetSizeByRealIndex(right), columns.Max);
            var y2 = Math.Min(GetRowTop(bottom) + _rowSizes.GetSizeByRealIndex(bottom), rows.Max);
            if (x2 > x1 && y2 > y1)
                yield return new Rect(x1 / DpiDetector.DpiXKoef, y1 / DpiDetector.DpiYKoef,
                    (x2 - x1) / DpiDetector.DpiXKoef, (y2 - y1) / DpiDetector.DpiYKoef);
        }
    }

    internal void RenderSelectionOverlay(DrawingContext drawing)
    {
        if (!UseSelectionOverlay || _drawBuffer == null) return;
        var color = IsLimitedSelection ? LimitedSelectedColor : SelectedColor;
        var fill = GetSolidBrush(Color.FromArgb(45, color.R, color.G, color.B));
        var pen = new Pen(GetSolidBrush(color), 1); pen.Freeze();
        var header = GetSolidBrush(Color.FromArgb(110, HeaderCurrentBackground.R, HeaderCurrentBackground.G, HeaderCurrentBackground.B));
        var hover = GetSolidBrush(Color.FromArgb(35, color.R, color.G, color.B));

        // Hover is visual-only as well. Avoid highlighting unrelated cells during a drag.
        if (!IsSelectionDragging && _mouseOverRow is { } hoverRow && _realColumnCount > 0)
            foreach (var rect in GetOverlayRectangles(new(hoverRow, 0, hoverRow, _realColumnCount - 1)))
                drawing.DrawRectangle(hover, null, rect);

        foreach (var range in _selectedCells.Ranges)
            foreach (var rect in GetOverlayRectangles(range)) drawing.DrawRectangle(fill, pen, rect);

        // Selected cells already share the range outline. Drawing the current cell again
        // adds a nested box during range selection (and doubles a single-cell border).
        if (_currentCell.IsCell && !_selectedCells.Contains(_currentCell))
            foreach (var rect in GetOverlayRectangles(GetMergedRealCellRange(_currentCell)))
                drawing.DrawRectangle(null, pen, rect);

        void RowHeader(int? row, Brush brush)
        {
            if (!row.HasValue || _realColumnCount == 0) return;
            foreach (var rect in GetOverlayRectangles(new(row.Value, 0, row.Value, _realColumnCount - 1)))
            { drawing.DrawRectangle(brush, null, new Rect(0, rect.Top, HeaderWidth / DpiDetector.DpiXKoef, rect.Height)); break; }
        }
        void ColumnHeader(int? column, Brush brush)
        {
            if (!column.HasValue || _realRowCount == 0) return;
            foreach (var rect in GetOverlayRectangles(new(0, column.Value, _realRowCount - 1, column.Value)))
            { drawing.DrawRectangle(brush, null, new Rect(rect.Left, 0, rect.Width, HeaderHeight / DpiDetector.DpiYKoef)); break; }
        }
        RowHeader(_currentCell.Row, header); ColumnHeader(_currentCell.Column, header);
        RowHeader(_mouseOverRowHeader, hover); ColumnHeader(_mouseOverColumnHeader, hover);
    }
}
