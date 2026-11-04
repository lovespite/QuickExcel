using Avalonia;
using Avalonia.Input;
using Avalonia.Media;
using FastWpfGrid;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Grid.Avalonia;

public sealed record DrawingObjectChange(Guid DocumentId, SheetId SheetId, long Revision, DrawingObjectId Target, DrawingAnchor Original, DrawingAnchor Placement);

public sealed partial class WorkbookGrid
{
    private DrawingObjectId? _selectedObject;
    private DrawingAnchor? _objectOriginal, _objectPreview;
    private DrawingPlacement _objectPlacement;
    private DrawingPoint _objectStart, _objectOffset;
    private DrawingHandle _objectHandle;
    private IPointer? _objectPointer;
    private long _objectRevision;
    public DrawingObjectId? SelectedDrawingObject => _selectedObject;
    public void SelectDrawingObject(DrawingObjectId target)
    {
        if (_snapshot is null || ObjectAnchor(target) is null) throw new ArgumentOutOfRangeException(nameof(target));
        CancelObjectGesture(); _selectedObject = target; CellEditor.IsVisible = false; InvalidateContent(); DrawingObjectSelectionChanged?.Invoke();
    }
    public bool IsDrawingObjectDragging => _objectPointer is not null;
    public event Action<DrawingObjectChange>? DrawingObjectChanged;
    public event Action? DrawingObjectSelectionChanged;
    public bool CanEditDrawingObjects => _snapshot is not null && _snapshot.FileState?.FormulaReadOnly != true && Sheet.Data.Protection?.Objects != true && !IsComposing && !IsCommitPending;

    private DrawingAnchor? ObjectAnchor(DrawingObjectId id) => id.Index < 0 ? null : id.Kind == DrawingObjectKind.Image
        ? id.Index < Sheet.Images.Length ? Sheet.Images[id.Index].Anchor : null
        : id.Index < Sheet.Data.Charts.Length ? Sheet.Data.Charts[id.Index].Anchor : null;
    private DrawingAnchor DisplayAnchor(DrawingObjectId id, DrawingAnchor anchor) => _selectedObject == id && _objectPreview is { } preview ? preview : anchor;

    private IEnumerable<Fragment> ObjectFragments(DrawingAnchor anchor)
    {
        var bounds = anchor.Bounds(Sheet); if (bounds.Width <= 0 || bounds.Height <= 0) yield break;
        var extent = DrawingGeometry.RotatedBounds(bounds, anchor.Rotation);
        // Include the rotated corners when choosing frozen panes and viewport clips.
        extent = new(Math.Max(0, extent.X), Math.Max(0, extent.Y), extent.Width + Math.Min(0, extent.X), extent.Height + Math.Min(0, extent.Y));
        DrawingAnchor cells;
        try { cells = DrawingGeometry.FromBounds(Sheet, extent, DrawingAnchorKind.TwoCell); }
        catch (ArgumentException) { cells = DrawingGeometry.FromBounds(Sheet, bounds, DrawingAnchorKind.TwoCell); }
        var range = new GridCellRange(cells.Row, cells.Column, Math.Min(_rows.Count - 1, cells.EndRow), Math.Min(_columns.Count - 1, cells.EndColumn));
        foreach (var fragment in Fragments(range))
        {
            var rect = new Rect(fragment.Bounds.X + bounds.X - DrawingGeometry.ColumnStart(Sheet, range.Left),
                fragment.Bounds.Y + bounds.Y - DrawingGeometry.RowStart(Sheet, range.Top), bounds.Width, bounds.Height);
            var rotated = DrawingGeometry.RotatedBounds(new(rect.X, rect.Y, rect.Width, rect.Height), anchor.Rotation);
            if (new Rect(rotated.X, rotated.Y, rotated.Width, rotated.Height).Intersects(fragment.Clip)) yield return new(rect, fragment.Clip);
        }
    }
    private static DrawingPlacement Placement(Rect rect, double rotation) => new(new(rect.X, rect.Y, rect.Width, rect.Height), rotation);
    public Point? GetDrawingObjectHandle(DrawingObjectId target, DrawingHandle handle)
    {
        if (_snapshot is null || ObjectAnchor(target) is not { } anchor) return null;
        anchor = DisplayAnchor(target, anchor);
        foreach (var fragment in ObjectFragments(anchor))
        {
            var p = DrawingObjectInteraction.HandlePoint(Placement(fragment.Bounds, anchor.Rotation), handle, 28 / _zoom);
            if (fragment.Clip.Contains(new Point(p.X, p.Y))) return new(p.X * _zoom, p.Y * _zoom);
        }
        return null;
    }
    public void ClearDrawingObjectSelection()
    {
        CancelObjectGesture();
        if (_selectedObject is null) return;
        _selectedObject = null; InvalidateContent(); DrawingObjectSelectionChanged?.Invoke();
    }
    private (DrawingObjectId Id, DrawingHandle Handle, Fragment Fragment)? ObjectHit(Point screen, bool handles = true)
    {
        var point = new Point(screen.X / _zoom, screen.Y / _zoom);
        if (handles && _selectedObject is { } selected && ObjectAnchor(selected) is { } anchor && CanEditDrawingObjects)
            foreach (var fragment in ObjectFragments(DisplayAnchor(selected, anchor)))
            foreach (var handle in Enum.GetValues<DrawingHandle>().Where(h => h != DrawingHandle.Move))
            {
                var p = DrawingObjectInteraction.HandlePoint(Placement(fragment.Bounds, DisplayAnchor(selected, anchor).Rotation), handle, 28 / _zoom);
                if (fragment.Clip.Contains(point) && Math.Abs(point.X - p.X) <= 6 / _zoom && Math.Abs(point.Y - p.Y) <= 6 / _zoom) return (selected, handle, fragment);
            }
        foreach (var id in Enumerable.Range(0, Sheet.Data.Charts.Length).Reverse().Select(i => new DrawingObjectId(DrawingObjectKind.Chart, i))
            .Concat(Enumerable.Range(0, Sheet.Images.Length).Reverse().Select(i => new DrawingObjectId(DrawingObjectKind.Image, i))))
        {
            var placement = DisplayAnchor(id, ObjectAnchor(id)!);
            foreach (var fragment in ObjectFragments(placement))
                if (fragment.Clip.Contains(point) && DrawingObjectInteraction.Contains(Placement(fragment.Bounds, placement.Rotation), new(point.X, point.Y))) return (id, DrawingHandle.Move, fragment);
        }
        return null;
    }
    private bool ObjectPointerDown(PointerPressedEventArgs e, Point point)
    {
        if (ObjectHit(point) is not { } hit) { ClearDrawingObjectSelection(); return false; }
        _selectedObject = hit.Id; CellEditor.IsVisible = false; Focus(); InvalidateContent(); DrawingObjectSelectionChanged?.Invoke(); e.Handled = true;
        if (!CanEditDrawingObjects || !e.GetCurrentPoint(_canvas).Properties.IsLeftButtonPressed) return true;
        _objectOriginal = ObjectAnchor(hit.Id)!; _objectPreview = _objectOriginal; _objectRevision = _snapshot!.Revision; _objectHandle = hit.Handle;
        var bounds = _objectOriginal.Bounds(Sheet); _objectPlacement = new(bounds, _objectOriginal.Rotation);
        _objectOffset = new(bounds.X - hit.Fragment.Bounds.X, bounds.Y - hit.Fragment.Bounds.Y);
        _objectStart = new(point.X / _zoom + _objectOffset.X, point.Y / _zoom + _objectOffset.Y);
        _objectPointer = e.Pointer; e.Pointer.Capture(_canvas); return true;
    }
    private bool ObjectPointerMove(PointerEventArgs e)
    {
        if (_objectPointer is null) return false;
        if (_snapshot!.Revision != _objectRevision || !CanEditDrawingObjects) { CancelObjectGesture(); return true; }
        var p = e.GetPosition(_canvas); var current = new DrawingPoint(p.X / _zoom + _objectOffset.X, p.Y / _zoom + _objectOffset.Y);
        var placement = DrawingObjectInteraction.Update(_objectPlacement, _objectHandle, _objectStart, current, e.KeyModifiers.HasFlag(KeyModifiers.Shift));
        try { _objectPreview = DrawingGeometry.FromBounds(Sheet, placement.Bounds, _objectOriginal!.Kind, placement.Rotation); }
        catch (ArgumentException) { return true; } // Keep the last valid preview at the sheet boundary.
        InvalidateContent(); e.Handled = true; return true;
    }
    private bool ObjectPointerUp(PointerReleasedEventArgs e)
    {
        if (_objectPointer is null) return false;
        var original = _objectOriginal!; var placement = _objectPreview!; var target = _selectedObject!.Value; var revision = _objectRevision;
        CancelObjectGesture(); e.Handled = true;
        if (_snapshot!.Revision == revision && CanEditDrawingObjects && original != placement)
            DrawingObjectChanged?.Invoke(new(_snapshot.DocumentId, Sheet.Id, revision, target, original, placement));
        Focus(); return true;
    }
    private void CancelObjectGesture()
    {
        var pointer = _objectPointer; _objectPointer = null; _objectOriginal = _objectPreview = null;
        if (pointer is not null) { pointer.Capture(null); InvalidateContent(); }
    }
    private bool ObjectKey(KeyEventArgs e)
    {
        if (_selectedObject is not { } target) return false;
        if (e.Key == Key.Escape)
        { if (IsDrawingObjectDragging) CancelObjectGesture(); else { ClearDrawingObjectSelection(); PrepareEditor(true); } e.Handled = true; return true; }
        if (e.Key is not (Key.Left or Key.Right or Key.Up or Key.Down)) return false;
        e.Handled = true; if (!CanEditDrawingObjects || ObjectAnchor(target) is not { } original) return true;
        var step = e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? 10 : 1; var b = original.Bounds(Sheet);
        var x = e.Key == Key.Left ? -step : e.Key == Key.Right ? step : 0; var y = e.Key == Key.Up ? -step : e.Key == Key.Down ? step : 0;
        var placement = DrawingGeometry.FromBounds(Sheet, b with { X = Math.Max(0, b.X + x), Y = Math.Max(0, b.Y + y) }, original.Kind, original.Rotation);
        DrawingObjectChanged?.Invoke(new(_snapshot!.DocumentId, Sheet.Id, _snapshot.Revision, target, original, placement)); return true;
    }
    private void PaintDrawingSelection(DrawingContext context)
    {
        if (_selectedObject is not { } target || ObjectAnchor(target) is not { } anchor) return;
        anchor = DisplayAnchor(target, anchor); var pen = new Pen(Brushes.DodgerBlue, 1.5 / _zoom);
        foreach (var fragment in ObjectFragments(anchor))
        {
            using var clip = context.PushClip(fragment.Clip); var placement = Placement(fragment.Bounds, anchor.Rotation);
            Point Handle(DrawingHandle handle) { var p = DrawingObjectInteraction.HandlePoint(placement, handle, 28 / _zoom); return new(p.X, p.Y); }
            var corners = new[] { DrawingHandle.TopLeft, DrawingHandle.TopRight, DrawingHandle.BottomRight, DrawingHandle.BottomLeft };
            for (var i = 0; i < 4; i++) context.DrawLine(pen, Handle(corners[i]), Handle(corners[(i + 1) % 4]));
            if (!CanEditDrawingObjects) continue;
            context.DrawLine(pen, Handle(DrawingHandle.Top), Handle(DrawingHandle.Rotate));
            foreach (var handle in Enum.GetValues<DrawingHandle>().Where(h => h != DrawingHandle.Move))
            {
                var p = Handle(handle); var size = 8 / _zoom;
                if (handle == DrawingHandle.Rotate) context.DrawEllipse(Brushes.White, pen, p, size / 2, size / 2);
                else context.DrawRectangle(Brushes.White, pen, new Rect(p.X - size / 2, p.Y - size / 2, size, size));
            }
        }
    }
}
