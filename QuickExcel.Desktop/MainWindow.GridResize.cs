using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private sealed class GridResize(Thumb grip, SheetId sheetId, SheetAxis axis, int index,
        DataGridRow? row, DataGridColumn? column, double pixels)
    {
        public Thumb Grip { get; } = grip;
        public SheetId SheetId { get; } = sheetId;
        public SheetAxis Axis { get; } = axis;
        public int Index { get; } = index;
        public DataGridRow? Row { get; } = row;
        public DataGridColumn? Column { get; } = column;
        public double OriginalPixels { get; } = pixels;
        public double Pixels { get; set; } = pixels;
    }

    private GridResize? _gridResize;
    private DispatcherOperation? _resizePreview;

    private void InstallGridResizeHandlers(DataGrid grid)
    {
        GridResizeInput.Enable(grid);
        grid.AddHandler(Thumb.DragStartedEvent, new DragStartedEventHandler(GridResize_Started), true);
        grid.AddHandler(Thumb.DragDeltaEvent, new DragDeltaEventHandler(GridResize_Delta), true);
        grid.AddHandler(Thumb.DragCompletedEvent, new DragCompletedEventHandler(GridResize_Completed), true);
    }

    private GridResize? FindGridResize(DataGrid grid, Thumb grip)
    {
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        if (grip.TemplatedParent is DataGridColumnHeader header && grid.CanUserResizeColumns)
        {
            var column = grip.Name switch
            {
                "PART_RightHeaderGripper" => header.Column,
                "PART_LeftHeaderGripper" => grid.Columns.Where(c => c.DisplayIndex < header.DisplayIndex && c.Visibility == Visibility.Visible)
                    .OrderByDescending(c => c.DisplayIndex).FirstOrDefault(),
                _ => null
            };
            if (column is { CanUserResize: true })
                return new(grip, sheet.Id, SheetAxis.Column, column.DisplayIndex, null, column, column.ActualWidth);
        }
        if (grip.TemplatedParent is DataGridRowHeader rowHeader && grid.CanUserResizeRows &&
            ItemsControl.ContainerFromElement(grid, rowHeader) is DataGridRow headerRow)
        {
            DataGridRow? row = headerRow;
            if (grip.Name == "PART_TopHeaderGripper")
            {
                var index = grid.ItemContainerGenerator.IndexFromContainer(row);
                row = index > 0 ? grid.ItemContainerGenerator.ContainerFromIndex(index - 1) as DataGridRow : null;
            }
            else if (grip.Name != "PART_BottomHeaderGripper") return null;
            if (row?.Item is SheetRow item && row.ActualHeight > 0 &&
                sheet.Rows.GetValueOrDefault(item.Index)?.Hidden != true && !_grid.IsFiltered(item.Index))
                return new(grip, sheet.Id, SheetAxis.Row, item.Index, row, null, row.ActualHeight);
        }
        return null;
    }

    private void GridResize_Started(object sender, DragStartedEventArgs e)
    {
        if (e.OriginalSource is not Thumb grip || grip.TemplatedParent is not (DataGridColumnHeader or DataGridRowHeader)) return;
        if (FindGridResize((DataGrid)sender, grip) is not { } resize || _busy || !CommitEditors())
        { grip.CancelDrag(); return; }
        _gridResize = resize;
    }

    private void GridResize_Delta(object sender, DragDeltaEventArgs e)
    {
        if (_gridResize is not { } resize || !ReferenceEquals(e.OriginalSource, resize.Grip)) return;
        var change = resize.Axis == SheetAxis.Row ? e.VerticalChange : e.HorizontalChange;
        if (!double.IsFinite(change)) return;
        resize.Pixels = Math.Clamp(resize.Pixels + change, resize.Axis == SheetAxis.Row ? 1 : 12, 400);
        // Merge queued pointer moves before layout; only the latest size reaches the view.
        _resizePreview ??= Dispatcher.BeginInvoke(DispatcherPriority.Render, new Action(ApplyResizePreview));
    }

    private void ApplyResizePreview()
    {
        _resizePreview = null;
        if (_gridResize is not { } resize) return;
        if (resize.Row is { } row)
        {
            row.Height = resize.Pixels;
            foreach (var grid in new[] { SheetGrid, FrozenGrid })
                if (grid.ItemContainerGenerator.ContainerFromItem(_rows[resize.Index]) is DataGridRow counterpart && !ReferenceEquals(row, counterpart))
                    counterpart.Height = resize.Pixels;
            _artwork?.PreviewRowHeight(resize.Index, resize.Pixels - resize.OriginalPixels);
            _frozenArtwork?.PreviewRowHeight(resize.Index, resize.Pixels - resize.OriginalPixels);
        }
        else
        {
            resize.Column!.Width = resize.Pixels;
            if (resize.Index < FrozenGrid.Columns.Count) FrozenGrid.Columns[resize.Index].Width = resize.Pixels;
        }
    }

    private void GridResize_Completed(object sender, DragCompletedEventArgs e)
    {
        if (_gridResize is not { } resize || !ReferenceEquals(e.OriginalSource, resize.Grip)) return;
        _resizePreview?.Abort();
        _resizePreview = null;
        _gridResize = null;
        if (_document.Snapshot.Sheets[_sheetIndex].Id != resize.SheetId) return;
        try
        {
            if (!e.Canceled && Math.Abs(resize.Pixels - resize.OriginalPixels) > 0.01 && !_busy && CommitEditors())
            {
                var size = resize.Axis == SheetAxis.Row ? resize.Pixels : (resize.Pixels - 5) / 7;
                _editor.Execute(new SetAxisMetadataCommand(resize.SheetId, resize.Axis, resize.Index, 1, size));
            }
        }
        catch (Exception ex) { ReportEditFailure(ex); }
        finally
        {
            // Also restores canceled/rejected previews, without adding undo entries.
            _grid.RefreshDimensions(_document.Snapshot);
            RefreshWorkspaceDimensions();
            UpdateChrome();
            // Only an actual pending cell/script draft needs calculation, not the size change.
            QueueRecalculation();
        }
    }
}
