using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using FastWpfGrid;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    private Point _dragPoint;
    private SheetAxis? _dragAxis, _resizeAxis;
    private GridCellRange[] _dragBase = [];
    private bool _addSelection, _animationPending, _referenceGesture, _endNavigation;
    public event Action? ReferenceGestureEnded;
    private CellPosition _referenceAnchor;
    private TimeSpan? _lastDragFrame;
    private int _resizeIndex;
    private double _resizeStart, _resizeOriginal, _resizePixels;
    private long _gestureRevision;
    private GridCellRange? _fillSource, _fillPreview, _referenceRange;
    public event Action<SheetAxis, int, double, long>? DimensionChanged;
    public event Action<SheetAxis, int>? AutoFitRequested;
    public event Action<GridRange, GridRange, long>? FillRequested;
    public event Action<int>? FilterRequested;
    public Func<GridRange, bool>? ReferenceRequested { get; set; }
    public bool IsSelectionDragging => _dragging;
    public bool IsContextSelectionBlocked { get; private set; }
    public (double X, double Y) ScrollOffset => (_columns.Offset, _rows.Offset);
    public CellPosition? ScrollingCell
    {
        get
        {
            var row = _rows.Visible(ViewHeight).FirstOrDefault(r => !r.Frozen);
            var column = _columns.Visible(ViewWidth).FirstOrDefault(c => !c.Frozen);
            return row.Size > 0 && column.Size > 0 ? new(row.ModelIndex, column.ModelIndex) : null;
        }
    }
    public void ScrollTo(double x, double y) { if (Commit()) Scroll(x, y); }
    public Rect? GetCellBounds(CellPosition position) => CellBounds(position);
    public CellPosition? HitTestCell(Point point) => Hit(point);
    public void HighlightReference(GridRange? range)
    { _referenceRange = range is { } r ? new(r.Top, r.Left, r.Bottom, r.Right) : null; _overlay.InvalidateVisual(); }
    public void SelectRanges(IEnumerable<GridRange> ranges, CellPosition current)
    {
        var captured = ranges.Take(65).ToArray();
        if (!CanSelectRanges(captured, current)) return;
        if (!Commit() || !(BeforeSelection?.Invoke() ?? true)) return;
        ClearDrawingObjectSelection();
        _selection.Set(captured.Select(r => GridMergedSelection.Expand(new(r.Top, r.Left, r.Bottom, r.Right), _merges)));
        _current = current; _anchor = current; _overlay.InvalidateVisual(); PrepareEditor(false); SelectionChanged?.Invoke();
    }
    private bool CanSelectRanges(IReadOnlyList<GridRange> ranges, CellPosition current)
    {
        if (_snapshot is null || ranges.Count is 0 or > 64 || current.Row < 0 || current.Row >= _rows.Count || current.Column < 0 || current.Column >= _columns.Count) return false;
        if (ranges.Any(r => r.Top < 0 || r.Left < 0 || r.Bottom < r.Top || r.Right < r.Left || r.Bottom >= _rows.Count || r.Right >= _columns.Count)) return false;
        return ranges.Any(r => current.Row >= r.Top && current.Row <= r.Bottom && current.Column >= r.Left && current.Column <= r.Right) &&
            ProtectionSafety.CanSelect(_snapshot, Sheet, current.Row, current.Column);
    }
    private (SheetAxis Axis, int Index, double Size)? ResizeHit(Point point)
    {
        var x = point.X / _zoom; var y = point.Y / _zoom;
        if (y < HeaderHeight && x >= HeaderWidth)
            foreach (var cell in _columns.Visible(ViewWidth))
                if (Math.Abs(x - HeaderWidth - cell.Start - cell.Size) <= 4 / _zoom) return (SheetAxis.Column, cell.ModelIndex, cell.Size);
        if (x < HeaderWidth && y >= HeaderHeight)
            foreach (var cell in _rows.Visible(ViewHeight))
                if (Math.Abs(y - HeaderHeight - cell.Start - cell.Size) <= 4 / _zoom) return (SheetAxis.Row, cell.ModelIndex, cell.Size);
        return null;
    }
    private (SheetAxis Axis, int Index)? HeaderHit(Point point)
    {
        if (_snapshot is null) return null;
        var x = point.X / _zoom; var y = point.Y / _zoom;
        if (x >= 0 && x < HeaderWidth && y >= HeaderHeight && _rows.HitTest(y - HeaderHeight, ViewHeight) is var row && row >= 0)
            return (SheetAxis.Row, row);
        if (y >= 0 && y < HeaderHeight && x >= HeaderWidth && _columns.HitTest(x - HeaderWidth, ViewWidth) is var column && column >= 0)
            return (SheetAxis.Column, column);
        return null;
    }
    private Rect? FillHandle()
    {
        if (_selection.Ranges.Count != 1) return null;
        var range = _selection.Ranges[0];
        foreach (var fragment in Fragments(range))
            if (fragment.Clip.Contains(fragment.Bounds.BottomRight))
                return new Rect(fragment.Bounds.Right * _zoom - 4, fragment.Bounds.Bottom * _zoom - 4, 8, 8);
        return null;
    }
    private void PointerDown(object? sender, PointerPressedEventArgs e)
    {
        IsContextSelectionBlocked = false;
        if (_snapshot is null || IsComposing || _editing && CellEditor.IsPointerOver) return;
        var point = e.GetPosition(_canvas); var properties = e.GetCurrentPoint(_canvas).Properties;
        if (properties.IsRightButtonPressed)
        {
            if (!Commit() || !(BeforeSelection?.Invoke() ?? true)) { IsContextSelectionBlocked = true; return; }
            if (ObjectPointerDown(e, point)) return;
            if (HeaderHit(point) is { } header)
            {
                // Keep the clicked boundary even when the row/column intersects a merged cell.
                if (header.Axis == SheetAxis.Row)
                    SelectRange(new(header.Index, 0, header.Index, _columns.Count - 1), new(header.Index, _current.Column));
                else SelectRange(new(0, header.Index, _rows.Count - 1, header.Index), new(_current.Row, header.Index));
                IsContextSelectionBlocked = (header.Axis == SheetAxis.Row ? _current.Row : _current.Column) != header.Index;
            }
            else if (Hit(point) is { } target && !_selection.Contains(new(target.Row, target.Column))) Select(target);
            return;
        }
        if (!properties.IsLeftButtonPressed) return;
        if (Hit(point) is { } reference && ReferenceRequested?.Invoke(new(reference.Row, reference.Column, reference.Row, reference.Column)) == true)
        {
            _referenceGesture = _dragging = true; _referenceAnchor = reference; _dragPoint = point;
            e.Pointer.Capture(_canvas); e.Handled = true; RequestDragFrame(); return;
        }
        if (!Commit() || !(BeforeSelection?.Invoke() ?? true)) return;
        if (ObjectPointerDown(e, point)) return;
        _gestureRevision = _snapshot.Revision;
        if (Hit(point) is { } filterCell && FilterHeader(filterCell.Row, filterCell.Column) is not null &&
            CellBounds(filterCell) is { } filterBounds && point.X >= filterBounds.Right - 22 * _zoom)
        { Select(filterCell); FilterRequested?.Invoke(filterCell.Column); e.Handled = true; return; }
        if (ResizeHit(point) is { } resize)
        {
            if (e.ClickCount == 2) { AutoFitRequested?.Invoke(resize.Axis, resize.Index); e.Handled = true; return; }
            _resizeAxis = resize.Axis; _resizeIndex = resize.Index; _resizeOriginal = _resizePixels = resize.Size;
            _resizeStart = resize.Axis == SheetAxis.Row ? point.Y : point.X; e.Pointer.Capture(_canvas); e.Handled = true; return;
        }
        if (FillHandle() is { } handle && handle.Contains(point))
        {
            _fillSource = _fillPreview = _selection.Ranges[0]; _dragging = true; _dragPoint = point;
            CellEditor.IsVisible = false; e.Pointer.Capture(_canvas); e.Handled = true; RequestDragFrame(); return;
        }
        var row = _rows.HitTest(point.Y / _zoom - HeaderHeight, ViewHeight);
        var column = _columns.HitTest(point.X / _zoom - HeaderWidth, ViewWidth);
        if (point.X < HeaderWidth * _zoom && point.Y < HeaderHeight * _zoom)
        { SelectRange(new(0, 0, _rows.Count - 1, _columns.Count - 1), _current); e.Handled = true; return; }
        _dragAxis = point.X < HeaderWidth * _zoom ? SheetAxis.Row : point.Y < HeaderHeight * _zoom ? SheetAxis.Column : null;
        if (_dragAxis != SheetAxis.Column && row < 0 || _dragAxis != SheetAxis.Row && column < 0) return;
        var position = new CellPosition(row < 0 ? _current.Row : row, column < 0 ? _current.Column : column);
        if (_dragAxis is null) position = CellAnchor(position);
        if (!ProtectionSafety.CanSelect(_snapshot, Sheet, position.Row, position.Column)) return;
        _addSelection = e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta);
        _dragBase = _selection.Ranges.ToArray();
        if (!e.KeyModifiers.HasFlag(KeyModifiers.Shift)) _anchor = position;
        _current = position; _dragPoint = point;
        UpdateDragSelection(position, initial: true);
        CellClicked?.Invoke();
        if (e.ClickCount == 2 && _dragAxis is null) { BeginEdit(); e.Handled = true; return; }
        _dragging = true; CellEditor.IsVisible = false; e.Pointer.Capture(_canvas); e.Handled = true; RequestDragFrame();
    }
    private void UpdateDragSelection(CellPosition cell, bool initial = false)
    {
        if (_referenceGesture)
        { var reference = GridRange.Between(_referenceAnchor, cell); ReferenceRequested?.Invoke(reference); HighlightReference(reference); return; }
        if (_fillSource is { } source)
        {
            var rowDelta = cell.Row < source.Top ? cell.Row - source.Top : Math.Max(0, cell.Row - source.Bottom);
            var colDelta = cell.Column < source.Left ? cell.Column - source.Left : Math.Max(0, cell.Column - source.Right);
            _fillPreview = Math.Abs(rowDelta) >= Math.Abs(colDelta)
                ? source with { Top = Math.Min(cell.Row, source.Top), Bottom = Math.Max(cell.Row, source.Bottom) }
                : source with { Left = Math.Min(cell.Column, source.Left), Right = Math.Max(cell.Column, source.Right) };
            _overlay.InvalidateVisual(); return;
        }
        var r = GridRange.Between(_anchor, cell);
        var range = _dragAxis switch
        {
            SheetAxis.Row => new GridCellRange(r.Top, 0, r.Bottom, _columns.Count - 1),
            SheetAxis.Column => new GridCellRange(0, r.Left, _rows.Count - 1, r.Right),
            _ => GridMergedSelection.Expand(new(r.Top, r.Left, r.Bottom, r.Right), _merges)
        };
        if (_addSelection) { _selection.Set(_dragBase); if (initial) _selection.Toggle(range); else _selection.Add(range); }
        else _selection.Set(range);
        // Keep raw hit coordinates for range expansion and header boundaries, but
        // cell editing/navigation always belongs to the merged cell's anchor.
        _current = _dragAxis is null ? CellAnchor(cell) : cell; _overlay.InvalidateVisual(); SelectionChanged?.Invoke();
    }
    private void PointerMove(object? sender, PointerEventArgs e)
    {
        if (ObjectPointerMove(e)) return;
        _dragPoint = e.GetPosition(_canvas);
        if (_resizeAxis is { } axis)
        {
            _resizePixels = Math.Clamp(_resizeOriginal + ((axis == SheetAxis.Row ? _dragPoint.Y : _dragPoint.X) - _resizeStart) / _zoom, 1, 10000);
            (axis == SheetAxis.Row ? _rows : _columns).PreviewSize(_resizeIndex, (int)_resizePixels);
            InvalidateContent(); UpdateScrollbars(); PositionEditor(); e.Handled = true; return;
        }
        if (!_dragging) { UpdateHover(_dragPoint); return; }
        var point = new Point(Math.Clamp(_dragPoint.X, HeaderWidth * _zoom, Math.Max(HeaderWidth * _zoom, _canvas.Bounds.Width - 1)),
            Math.Clamp(_dragPoint.Y, HeaderHeight * _zoom, Math.Max(HeaderHeight * _zoom, _canvas.Bounds.Height - 1)));
        if (Hit(point) is { } cell && ((_dragAxis is null ? CellAnchor(cell) : cell) != _current || _fillSource is not null || _referenceGesture)) UpdateDragSelection(cell);
        RequestDragFrame();
    }
    private void RequestDragFrame()
    {
        if (_animationPending || !_dragging) return;
        if (TopLevel.GetTopLevel(this) is not { } top) return;
        _animationPending = true; top.RequestAnimationFrame(DragFrame);
    }
    private void DragFrame(TimeSpan now)
    {
        _animationPending = false; if (!_dragging || _disposed) { _lastDragFrame = null; return; }
        var elapsed = _lastDragFrame is { } last ? Math.Clamp((now - last).TotalSeconds, 0, .05) : 1d / 60;
        _lastDragFrame = now;
        static double Speed(double value, double min, double max) => value < min ? -Math.Clamp((min - value) * 8, 120, 2400) : value > max ? Math.Clamp((value - max) * 8, 120, 2400) : 0;
        var x = _dragAxis == SheetAxis.Row ? 0 : Speed(_dragPoint.X, (HeaderWidth + _columns.FrozenSize) * _zoom, _canvas.Bounds.Width);
        var y = _dragAxis == SheetAxis.Column ? 0 : Speed(_dragPoint.Y, (HeaderHeight + _rows.FrozenSize) * _zoom, _canvas.Bounds.Height);
        if (x != 0 || y != 0)
        {
            Scroll(_columns.Offset + x * elapsed / _zoom, _rows.Offset + y * elapsed / _zoom);
            var point = new Point(Math.Clamp(_dragPoint.X, HeaderWidth * _zoom, Math.Max(HeaderWidth * _zoom, _canvas.Bounds.Width - 1)),
                Math.Clamp(_dragPoint.Y, HeaderHeight * _zoom, Math.Max(HeaderHeight * _zoom, _canvas.Bounds.Height - 1)));
            if (Hit(point) is { } cell) UpdateDragSelection(cell);
            DragFrameApplied?.Invoke(System.Diagnostics.Stopwatch.GetTimestamp());
        }
        RequestDragFrame();
    }
    private void PointerUp(object? sender, PointerReleasedEventArgs e)
    {
        if (ObjectPointerUp(e)) return;
        if (_resizeAxis is { } axis)
        {
            _resizeAxis = null;
            DimensionChanged?.Invoke(axis, _resizeIndex, _resizePixels, _gestureRevision);
            Configure(_rows.Count, _columns.Count); InvalidateContent(); UpdateScrollbars();
        }
        if (_fillSource is { } source && _fillPreview is { } extent && source != extent)
            FillRequested?.Invoke(new(source.Top, source.Left, source.Bottom, source.Right), new(extent.Top, extent.Left, extent.Bottom, extent.Right), _gestureRevision);
        var reference = _referenceGesture;
        _dragging = _referenceGesture = false; _fillSource = _fillPreview = null; _lastDragFrame = null;
        e.Pointer.Capture(null); _overlay.InvalidateVisual();
        if (!reference) PrepareEditor(true); else ReferenceGestureEnded?.Invoke();
    }
    private void CancelGesture()
    {
        CancelObjectGesture();
        var reference = _referenceGesture;
        _dragging = _referenceGesture = false; _fillSource = _fillPreview = null; _lastDragFrame = null;
        if (reference) ReferenceGestureEnded?.Invoke();
        if (_resizeAxis is not null) { _resizeAxis = null; Configure(_rows.Count, _columns.Count); InvalidateContent(); UpdateScrollbars(); }
        _overlay.InvalidateVisual(); PositionEditor();
    }
    private async void HandleKey(object? sender, KeyEventArgs e)
    {
        if (e.Key == Key.Apps || e.Key == Key.F10 && e.KeyModifiers.HasFlag(KeyModifiers.Shift)) IsContextSelectionBlocked = false;
        if (_snapshot is null || IsComposing) return;
        if (ObjectKey(e)) return;
        if (e.Key == Key.Escape && (_dragging || _resizeAxis is not null)) { CancelGesture(); e.Handled = true; return; }
        if (e.Key == Key.F2) { e.Handled = BeginEdit(); return; }
        if (e.Key == Key.Escape && _editing) { Cancel(); e.Handled = true; return; }
        var control = e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta);
        if (e.KeyModifiers.HasFlag(KeyModifiers.Alt)) return;
        if (control && e.Key == Key.A && !_editing) { SelectRange(new(0, 0, _rows.Count - 1, _columns.Count - 1), _current); e.Handled = true; return; }
        if (_editing && e.Key is not (Key.Enter or Key.Tab)) return;
        var delta = e.Key switch { Key.Up => (-1, 0), Key.Down => (1, 0), Key.Left => (0, -1), Key.Right => (0, 1),
            Key.PageUp => (-Math.Max(1, _rows.Visible(ViewHeight).Count - 1), 0), Key.PageDown => (Math.Max(1, _rows.Visible(ViewHeight).Count - 1), 0),
            Key.Enter => (e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? -1 : 1, 0), Key.Tab => (0, e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? -1 : 1), _ => (0, 0) };
        if (delta == (0, 0) && e.Key is not (Key.Home or Key.End)) return;
        e.Handled = true; if (!await CommitAsync() || !(BeforeSelection?.Invoke() ?? true)) return;
        if (e.Key == Key.End && !control) { _endNavigation = !_endNavigation; return; }
        var edge = (control || _endNavigation) && e.Key is Key.Up or Key.Down or Key.Left or Key.Right;
        _endNavigation = false;
        CellPosition target;
        if (e.Key == Key.Home) target = new(control ? 0 : _current.Row, 0);
        else if (e.Key == Key.End)
        { var cells = Sheet.Cells.Keys; target = new(control && cells.Any() ? cells.Max(p => p.Row) : _current.Row, cells.Any() ? cells.Max(p => p.Column) : 0); }
        else if (edge)
            target = EntryNavigation.Edge(Sheet, _current, delta.Item1, delta.Item2, _rows.Count, _columns.Count, r => _rows.Cell(r) is null);
        else
        {
            var merge = CellRange(_current);
            if (e.Key == Key.Tab && delta.Item2 > 0 && merge.Right == _columns.Count - 1)
                target = new(_rows.Neighbor(_current.Row, 1), 0);
            else if (e.Key == Key.Tab && delta.Item2 < 0 && merge.Left == 0)
                target = new(_rows.Neighbor(_current.Row, -1), _columns.Count - 1);
            else target = new(_rows.Neighbor(delta.Item1 > 0 ? merge.Bottom : _current.Row, delta.Item1), _columns.Neighbor(delta.Item2 > 0 ? merge.Right : _current.Column, delta.Item2));
        }
        SelectCore(target, e.Key is not (Key.Enter or Key.Tab) && e.KeyModifiers.HasFlag(KeyModifiers.Shift), true);
    }
}
