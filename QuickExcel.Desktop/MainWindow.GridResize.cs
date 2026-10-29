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
            if (BeginGridResize(hit.Axis, hit.Index, e.GetPosition(_grid.Control)))
                Mouse.Capture(SheetGrid, CaptureMode.Element);
        };
        SheetGrid.PreviewMouseMove += (_, e) =>
        {
            if (_gridResize is { } resize)
            {
                var point = e.GetPosition(_grid.Control);
                PreviewGridResize(resize.Axis == SheetAxis.Row ? point.Y - resize.Origin.Y : point.X - resize.Origin.X);
                e.Handled = true;
            }
            else SheetGrid.Cursor = FindGridResize(e.GetPosition(_grid.Control)) is { } hit
                ? hit.Axis == SheetAxis.Row ? Cursors.SizeNS : Cursors.SizeWE : null;
        };
        SheetGrid.PreviewMouseLeftButtonUp += (_, e) =>
        {
            if (_gridResize is null) return;
            CompleteGridResize(false); Mouse.Capture(null); e.Handled = true;
        };
        SheetGrid.LostMouseCapture += (_, _) => { if (_gridResize is not null) CompleteGridResize(true); };
        PreviewKeyDown += (_, e) =>
        {
            if (_gridResize is not null && e.Key == Key.Escape)
            { CompleteGridResize(true); Mouse.Capture(null); e.Handled = true; }
        };
    }

    private (SheetAxis Axis, int Index)? FindGridResize(Point point)
    {
        if (_document is null) return null;
        var grid = _grid.Control;
        if (point.Y >= 0 && point.Y < grid.ColumnHeaderHeight)
        {
            var row = grid.GetVisibleModelRows().FirstOrDefault();
            foreach (var column in grid.GetVisibleModelColumns())
                if (grid.GetModelCellBounds(row, column) is { } bounds && Math.Abs(point.X - bounds.Right) <= 4)
                    return (SheetAxis.Column, column);
        }
        if (point.X >= 0 && point.X < grid.RowHeaderWidth)
        {
            var column = grid.GetVisibleModelColumns().FirstOrDefault();
            foreach (var row in grid.GetVisibleModelRows())
                if (grid.GetModelCellBounds(row, column) is { } bounds && Math.Abs(point.Y - bounds.Bottom) <= 4)
                    return (SheetAxis.Row, row);
        }
        return null;
    }

    private bool BeginGridResize(SheetAxis axis, int index, Point origin = default)
    {
        if (_busy || !CommitEditors()) return false;
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        var pixels = axis == SheetAxis.Row ? Math.Clamp(sheet.Rows.GetValueOrDefault(index)?.Height ?? sheet.DefaultRowHeight ?? 31, 1, 10000) :
            sheet.Columns.GetValueOrDefault(index)?.Width is { } width ? Math.Clamp(width * 7 + 5, 1, 10000) : index == 0 ? 160 : 112;
        _gridResize = new(sheet.Id, axis, index, pixels, origin);
        return true;
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
