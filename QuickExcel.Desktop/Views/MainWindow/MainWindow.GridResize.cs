using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private sealed record GridResize(SheetId SheetId, SheetAxis Axis, int Index, double OriginalPixels, Point Origin)
    {
        public double Pixels { get; set; } = OriginalPixels;
    }
    private GridResize? _gridResize;
    private DispatcherOperation? _resizePreview;

    private void InstallGridResizeHandlers()
    {
        SheetGrid.PreviewMouseLeftButtonDown += (_, e) =>
        {
            if (e.Handled || FindGridResize(e.GetPosition(_grid.Control)) is not { } hit) return;
            e.Handled = true;
            if (e.ClickCount == 2)
            {
                AutoFitGridAxis(hit.Axis, hit.Index);
                UpdateGridResizeCursor(e.GetPosition(_grid.Control));
                return;
            }
            if (BeginGridResize(hit.Axis, hit.Index, e.GetPosition(_grid.Control)))
            {
                UpdateGridResizeCursor(e.GetPosition(_grid.Control));
                if (!Mouse.Capture(SheetGrid, CaptureMode.Element)) CompleteGridResize(true);
            }
        };
        SheetGrid.PreviewMouseMove += (_, e) =>
        {
            if (_gridResize is { } resize)
            {
                var point = e.GetPosition(_grid.Control);
                PreviewGridResize(resize.Axis == SheetAxis.Row ? point.Y - resize.Origin.Y : point.X - resize.Origin.X);
                e.Handled = true;
            }
            UpdateGridResizeCursor(e.GetPosition(_grid.Control));
        };
        SheetGrid.PreviewMouseLeftButtonUp += (_, e) =>
        {
            if (_gridResize is not { } resize) return;
            var point = e.GetPosition(_grid.Control);
            PreviewGridResize(resize.Axis == SheetAxis.Row ? point.Y - resize.Origin.Y : point.X - resize.Origin.X);
            CompleteGridResize(false); UpdateGridResizeCursor(point); e.Handled = true;
        };
        SheetGrid.LostMouseCapture += (_, _) =>
        { if (_gridResize is not null && Mouse.Captured != SheetGrid) CompleteGridResize(true); };
        SheetGrid.MouseLeave += (_, _) => { if (_gridResize is null) UpdateGridResizeCursor(new(-1, -1)); };
        Deactivated += (_, _) => CompleteGridResize(true);
        PreviewKeyDown += (_, e) =>
        {
            if (_gridResize is not null && e.Key == Key.Escape)
            { CompleteGridResize(true); e.Handled = true; }
        };
    }

    private void UpdateGridResizeCursor(Point point)
    {
        var axis = _gridResize?.Axis ?? (!_busy && !_grid.Control.IsSelectionDragging ? FindGridResize(point)?.Axis : null);
        SheetGrid.Cursor = axis is null ? null : axis == SheetAxis.Row ? Cursors.SizeNS : Cursors.SizeWE;
        // The child grid sets its own Arrow cursor during hover. Force only the header
        // resize cursor through that child; leave editors and scrollbars their own cursors.
        SheetGrid.ForceCursor = axis is not null;
    }

    private (SheetAxis Axis, int Index)? FindGridResize(Point point)
    {
        if (_document is null) return null;
        var header = _grid.Control.HitTestResizeHeader(point);
        if (header.IsColumnHeader) return (SheetAxis.Column, header.Column!.Value);
        if (header.IsRowHeader) return (SheetAxis.Row, header.Row!.Value);
        return null;
    }

    private bool BeginGridResize(SheetAxis axis, int index, Point origin = default)
    {
        if (_gridResize is not null || _busy || !CommitEditors()) return false;
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        var pixels = axis == SheetAxis.Row ? Math.Clamp(sheet.Rows.GetValueOrDefault(index)?.Height ?? sheet.DefaultRowHeight ?? 31, 1, 10000) :
            sheet.Columns.GetValueOrDefault(index)?.Width is { } width ? Math.Clamp(width * 7 + 5, 1, 10000) : index == 0 ? 160 : 112;
        _gridResize = new(sheet.Id, axis, index, pixels, origin);
        return true;
    }

    private void AutoFitGridAxis(SheetAxis axis, int index)
    {
        if (!BeginGridResize(axis, index)) return;
        try
        {
            // Committing an editor may have changed the text or style being measured.
            _grid.RefreshDimensions(_document.Snapshot);
            _gridResize!.Pixels = _grid.MeasureAutoFit(axis, index);
            CompleteGridResize(false);
        }
        catch (Exception ex) { CompleteGridResize(true); ReportEditFailure(ex); }
    }

    private void PreviewGridResize(double displacement)
    {
        if (_gridResize is not { } resize || !double.IsFinite(displacement)) return;
        resize.Pixels = Math.Clamp(resize.OriginalPixels + displacement, resize.Axis == SheetAxis.Row ? 1 : 12, 10000);
        _resizePreview ??= Dispatcher.BeginInvoke(DispatcherPriority.Render, new Action(() =>
        {
            _resizePreview = null;
            if (_gridResize is { } current) _grid.ApplyDimensions(current.Axis, current.Index, current.Pixels);
        }));
    }

    private void CompleteGridResize(bool cancel)
    {
        if (_gridResize is not { } resize) return;
        _resizePreview?.Abort(); _resizePreview = null; _gridResize = null;
        if (Mouse.Captured == SheetGrid) Mouse.Capture(null);
        UpdateGridResizeCursor(new(-1, -1));
        if (_document.Snapshot.Sheets[_sheetIndex].Id != resize.SheetId) return;
        try
        {
            if (!cancel && Math.Abs(resize.Pixels - resize.OriginalPixels) > .01 && !_busy && CommitEditors())
                _editor.Execute(new SetAxisMetadataCommand(resize.SheetId, resize.Axis, resize.Index, 1,
                    resize.Axis == SheetAxis.Row ? resize.Pixels : (resize.Pixels - 5) / 7));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
        finally
        {
            _grid.RefreshDimensions(_document.Snapshot); _grid.ApplyDimensions(); UpdateChrome(); QueueRecalculation();
        }
    }
}
