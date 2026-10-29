using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace FastWpfGrid;

partial class FastGridControl
{
    private bool _useSmoothScrolling;
    private double _scrollOffsetX, _scrollOffsetY;
    private long _lastDragScrollTimestamp;
    private bool _dragRenderingSubscribed;
    private TimeSpan? _lastDragRenderingTime;
    private int ScrollOffsetX => RetainedContentEnabled ? (int)Math.Round(_scrollOffsetX) : 0;
    private int ScrollOffsetY => RetainedContentEnabled ? (int)Math.Round(_scrollOffsetY) : 0;

    /// <summary>Opt-in pixel scrolling and retained content; requires the selection overlay.</summary>
    public bool UseSmoothScrolling
    {
        get => _useSmoothScrolling;
        set
        {
            if (_useSmoothScrolling == value) return;
            _useSmoothScrolling = value;
            _scrollOffsetX = _scrollOffsetY = 0; _lastDragScrollTimestamp = 0;
            _dragTimer.Interval = TimeSpan.FromMilliseconds(value ? 16 : 50);
            if (_dragRenderingSubscribed || _dragTimer.IsEnabled) StartDragScrolling();
            InvalidateAll();
        }
    }

    private void StartDragScrolling()
    {
        StopDragScrolling();
        if (RetainedContentEnabled)
        {
            CompositionTarget.Rendering += DragScrollRendering;
            _dragRenderingSubscribed = true;
        }
        // Rendering is throttled by WPF when occluded/minimized. A timer is a
        // fallback, not a second producer: both paths share the same time gate.
        _dragTimer.Start();
    }

    private void StopDragScrolling()
    {
        _dragTimer.Stop();
        if (_dragRenderingSubscribed) CompositionTarget.Rendering -= DragScrollRendering;
        _dragRenderingSubscribed = false; _lastDragScrollTimestamp = 0; _lastDragRenderingTime = null;
    }

    private void DragScrollRendering(object sender, EventArgs e)
    {
        if (!IsSelectionDragging || !IsLoaded) { StopDragScrolling(); return; }
        var time = ((RenderingEventArgs)e).RenderingTime;
        if (_lastDragRenderingTime == time) return;
        _lastDragRenderingTime = time;
        _dragTimer_Tick(sender, e);
    }

    private void InvalidateViewportPosition()
    {
        _isInvalidated = true; CheckInvalidation(); selectionOverlay.Refresh();
        NotifyViewportChanged();
    }

    private FastGridCellAddress GetDragEndpoint(Point point)
    {
        var cell = GetCellAddress(point);
        if (!IsSelectionDragging || !cell.IsEmpty || _drawBuffer == null || _realRowCount == 0 || _realColumnCount == 0 ||
            point.X >= 0 && point.Y >= 0 && point.X < _drawBuffer.PixelWidth && point.Y < _drawBuffer.PixelHeight) return cell;
        // A coalesced mouse move may jump straight outside and release before the
        // first scroll callback. Select the visible edge immediately, without scrolling.
        var x = Math.Clamp(point.X, Math.Min(HeaderWidth + FrozenWidth + 1, _drawBuffer.PixelWidth - 1), _drawBuffer.PixelWidth - 1);
        var y = Math.Clamp(point.Y, Math.Min(HeaderHeight + FrozenHeight + 1, _drawBuffer.PixelHeight - 1), _drawBuffer.PixelHeight - 1);
        var row = GetSeriesIndexOnPosition(y, HeaderHeight, _rowSizes, FirstVisibleRowScrollIndex);
        var column = GetSeriesIndexOnPosition(x, HeaderWidth, _columnSizes, FirstVisibleColumnScrollIndex);
        return new(row.HasValue ? Math.Clamp(row.Value, 0, _realRowCount - 1) : null,
            column.HasValue ? Math.Clamp(column.Value, 0, _realColumnCount - 1) : null);
    }

    // Time-based movement avoids speeding up with a faster timer or leaping after
    // a stalled dispatcher. The same method is used by deterministic regression probes.
    private void AdvanceDragAutoScroll(double elapsedSeconds)
    {
        if (!IsSelectionDragging || _drawBuffer == null || _realRowCount == 0 || _realColumnCount == 0) return;
        var seconds = Math.Clamp(elapsedSeconds, 0, 0.05);
        var pointer = _mouseCursorPoint;
        double Speed(bool before, bool after, double position, double extent, double dpi)
        {
            if (!before && !after) return 0;
            var outside = before ? -position : position - extent;
            var dipsPerSecond = 360 + Math.Clamp(outside / dpi, 0, 180) * 4;
            return (before ? -1 : 1) * dipsPerSecond * dpi;
        }
        var vx = Speed(_mouseIsBehindLeft, _mouseIsBehindRight, pointer?.X ?? (_mouseIsBehindLeft ? 0 : _drawBuffer.PixelWidth), _drawBuffer.PixelWidth, DpiDetector.DpiXKoef);
        var vy = Speed(_mouseIsBehindTop, _mouseIsBehindBottom, pointer?.Y ?? (_mouseIsBehindTop ? 0 : _drawBuffer.PixelHeight), _drawBuffer.PixelHeight, DpiDetector.DpiYKoef);
        bool Advance(SeriesSizes sizes, int extent, double delta, ref int first, ref double offset)
        {
            if (delta == 0 || sizes.ScrollCount == 0 || extent <= 0) return false;
            var old = sizes.GetPositionByScrollIndex(first) + offset;
            var end = sizes.GetPositionByScrollIndex(sizes.ScrollCount);
            var next = Math.Clamp(old + delta, 0, Math.Max(0, end - extent));
            if (Math.Abs(next - old) < 0.001) return false;
            first = Math.Clamp(sizes.GetScrollIndexOnPosition((int)Math.Floor(next)), 0, sizes.ScrollCount - 1);
            offset = next - sizes.GetPositionByScrollIndex(first);
            return true;
        }
        var movedY = Advance(_rowSizes, GridScrollAreaHeight, vy * seconds, ref FirstVisibleRowScrollIndex, ref _scrollOffsetY);
        var movedX = Advance(_columnSizes, GridScrollAreaWidth, vx * seconds, ref FirstVisibleColumnScrollIndex, ref _scrollOffsetX);
        if (!movedX && !movedY) return;
        using (CreateInvalidationContext())
        {
            // Re-hit-test the stationary pointer against the translated content,
            // including the partly visible trailing cell and frozen segments.
            var x = Math.Clamp(pointer?.X ?? HeaderWidth + FrozenWidth + 1, 0, _drawBuffer.PixelWidth - 1);
            var y = Math.Clamp(pointer?.Y ?? HeaderHeight + FrozenHeight + 1, 0, _drawBuffer.PixelHeight - 1);
            int? row = _mouseIsBehindBottom ? FirstVisibleRowScrollIndex + _rowSizes.FrozenCount + VisibleRowCount - 1 :
                _mouseIsBehindTop ? FirstVisibleRowScrollIndex + _rowSizes.FrozenCount :
                pointer.HasValue ? GetSeriesIndexOnPosition(y, HeaderHeight, _rowSizes, FirstVisibleRowScrollIndex) : _mouseMoveRow ?? _currentCell.Row;
            int? column = _mouseIsBehindRight ? FirstVisibleColumnScrollIndex + _columnSizes.FrozenCount + VisibleColumnCount - 1 :
                _mouseIsBehindLeft ? FirstVisibleColumnScrollIndex + _columnSizes.FrozenCount :
                pointer.HasValue ? GetSeriesIndexOnPosition(x, HeaderWidth, _columnSizes, FirstVisibleColumnScrollIndex) : _mouseMoveColumn ?? _currentCell.Column;
            if (row.HasValue) row = Math.Clamp(row.Value, 0, _realRowCount - 1);
            if (column.HasValue) column = Math.Clamp(column.Value, 0, _realColumnCount - 1);
            SetSelectedRectangle(_dragStartCell, new(row, column));
            AdjustScrollBarPositions();
            if (movedY) { if (IsTransposed) OnScrolledModelColumns(); else OnScrolledModelRows(); }
            if (movedX) { if (IsTransposed) OnScrolledModelRows(); else OnScrolledModelColumns(); }
            InvalidateViewportPosition();
        }
    }
}
