using System.IO;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunGridResizeSmokeAsync(string folder, Action<bool, string> check)
    {
        async Task Idle() => await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
        DataGridRow Row(DataGrid grid, int index) => (DataGridRow)grid.ItemContainerGenerator.ContainerFromItem(_rows[index]);
        Thumb RowGrip(DataGrid grid, int index, string side = "Bottom")
        {
            grid.UpdateLayout();
            var header = VisualChild<DataGridRowHeader>(Row(grid, index))!;
            return (Thumb)header.Template.FindName($"PART_{side}HeaderGripper", header);
        }
        Thumb ColumnGrip(int index, string side = "Right")
        {
            SheetGrid.ScrollIntoView(_rows[0], SheetGrid.Columns[index]); SheetGrid.UpdateLayout();
            var presenter = VisualChild<DataGridColumnHeadersPresenter>(SheetGrid)!;
            var header = (DataGridColumnHeader)presenter.ItemContainerGenerator.ContainerFromIndex(index);
            return (Thumb)header.Template.FindName($"PART_{side}HeaderGripper", header);
        }
        void Start(Thumb grip) => grip.RaiseEvent(new DragStartedEventArgs(0, 0) { RoutedEvent = Thumb.DragStartedEvent });
        async Task Delta(Thumb grip, double x, double y)
        { grip.RaiseEvent(new DragDeltaEventArgs(x, y) { RoutedEvent = Thumb.DragDeltaEvent }); await Idle(); }
        void End(Thumb grip, bool cancel = false) => grip.RaiseEvent(new DragCompletedEventArgs(0, 0, cancel) { RoutedEvent = Thumb.DragCompletedEvent });
        static bool Near(double actual, double expected) => Math.Abs(actual - expected) < 0.01;

        var path = Path.Combine(folder, "resized.xlsq");
        InstallDocument(WorkbookDocument.CreateBlank(_services), path, "尺寸回归.xlsq");
        await Idle();
        GridSurface.SelectRange(new(2, 2, 3, 3), new(2, 2)); await Idle();
        var selection = GridSurface.Selection.ToArray();
        var revision = _document.Snapshot.Revision;
        var grip = ColumnGrip(0);
        check(grip.IsVisible && SheetGrid.CanUserResizeColumns, "Column resize grip is enabled and visible");
        Start(grip); await Delta(grip, 14, 0); await Delta(grip, 21, 0);
        check(Near(SheetGrid.Columns[0].ActualWidth, 195) && _document.Snapshot.Revision == revision, "Column drag previews without document edits");
        End(grip); await Idle();
        check(_document.Snapshot.Revision == revision + 1 && Near(_document.Snapshot.Sheets[0].Columns[0].Width!.Value, (195d - 5) / 7), "Column drag commits exactly one command in character units");
        check(GridSurface.Selection.SequenceEqual(selection) && _selectedRow == 2 && _selectedColumn == 2, "Header resizing preserves selection and active cell");
        check(_document.IsDirty && Title.StartsWith("*"), "Resize marks the workbook dirty");
        RefreshSheet(); await Idle();
        check(Near(SheetGrid.Columns[0].ActualWidth, 195), "Refresh retains resized column width");
        await UndoAsync(false); await Idle();
        check(Near(SheetGrid.Columns[0].ActualWidth, 160) && !_document.Session.CanUndo, "One undo restores original column width with no extra history");
        await UndoAsync(true); await Idle();
        check(Near(SheetGrid.Columns[0].ActualWidth, 195), "Redo restores resized column width");

        revision = _document.Snapshot.Revision;
        grip = RowGrip(SheetGrid, 0); Start(grip); await Delta(grip, 0, 13); await Delta(grip, 0, 18);
        check(Near(Row(SheetGrid, 0).ActualHeight, 62) && _document.Snapshot.Revision == revision, "Row drag previews full row height without document edits");
        check(Near(VisualChild<DataGridCellsPresenter>(Row(SheetGrid, 0))!.ActualHeight, 62), "Row contents stretch with the resized row");
        End(grip); await Idle();
        check(_document.Snapshot.Revision == revision + 1 && _document.Snapshot.Sheets[0].Rows[0].Height == 62, "Row drag commits exactly one command");
        SheetGrid.ScrollIntoView(_rows[80], SheetGrid.Columns[0]); await Idle();
        SheetGrid.ScrollIntoView(_rows[0], SheetGrid.Columns[0]); await Idle();
        check(Near(Row(SheetGrid, 0).ActualHeight, 62) && Near(Row(SheetGrid, 1).ActualHeight, 31), "Recycling retains resized row without leaking height to neighbors");
        await UndoAsync(false); await Idle();
        check(Near(Row(SheetGrid, 0).ActualHeight, 31), "One undo restores original row height");
        await UndoAsync(true); await Idle();
        check(Near(Row(SheetGrid, 0).ActualHeight, 62), "Redo restores resized row height");

        revision = _document.Snapshot.Revision;
        grip = RowGrip(SheetGrid, 0); Start(grip); await Delta(grip, 0, 20); End(grip, true); await Idle();
        check(_document.Snapshot.Revision == revision && Near(Row(SheetGrid, 0).ActualHeight, 62), "Canceled row drag restores layout without history");
        grip = ColumnGrip(0); Start(grip); await Delta(grip, 40, 0); End(grip, true); await Idle();
        check(_document.Snapshot.Revision == revision && Near(SheetGrid.Columns[0].ActualWidth, 195), "Canceled column drag restores layout without history");
        grip = ColumnGrip(0); Start(grip); End(grip); await Idle();
        check(_document.Snapshot.Revision == revision, "Clicking a resize grip without movement creates no history");
        _editState.StartComposition();
        try
        {
            grip = ColumnGrip(0); Start(grip);
            check(_gridResize is null && _document.Snapshot.Revision == revision, "Composition gate rejects header resize without committing input");
            End(grip, true);
        }
        finally { _editState.EndComposition(); }

        grip = ColumnGrip(1, "Left"); Start(grip); await Delta(grip, 7, 0); End(grip); await Idle();
        check(Near(SheetGrid.Columns[0].ActualWidth, 202) && Near(SheetGrid.Columns[1].ActualWidth, 112), "Left column grip resizes the preceding column");
        grip = RowGrip(SheetGrid, 1, "Top"); Start(grip); await Delta(grip, 0, -7); End(grip); await Idle();
        check(Near(Row(SheetGrid, 0).ActualHeight, 55) && Near(Row(SheetGrid, 1).ActualHeight, 31), "Top row grip resizes the preceding row");

        GridSurface.Select(new(1, 1)); Freeze_Click(new MenuItem(), new()); await Idle();
        grip = RowGrip(FrozenGrid, 0); Start(grip); await Delta(grip, 0, 14); End(grip); await Idle();
        check(Near(Row(SheetGrid, 0).ActualHeight, 69) && Near(Row(FrozenGrid, 0).ActualHeight, 69), "Frozen row resize persists and synchronizes both panes");
        grip = ColumnGrip(0); Start(grip); await Delta(grip, 14, 0); End(grip); await Idle();
        check(Near(FrozenGrid.Columns[0].ActualWidth, 216) && CurrentView.FrozenRows == 1 && CurrentView.FrozenColumns == 1, "Column resize synchronizes frozen pane and preserves freeze state");
        check(await SaveAsync(false), "Resized workbook saves successfully");
        await OpenPathAsync(path); await Idle();
        check(Near(Row(SheetGrid, 0).ActualHeight, 69) && Near(SheetGrid.Columns[0].ActualWidth, 216), "Save and reopen preserve dragged row height and column width");

        GridSurface.Select(new(2, 2)); FormulaBar.Text = "resize draft";
        grip = RowGrip(SheetGrid, 0); Start(grip); await Delta(grip, 0, 10); End(grip); await Idle();
        check(_document.GetCellInput(0, 2, 2) == "resize draft", "Resizing commits a pending draft to its original cell");
        await UndoAsync(false); await Idle();
        check(Near(Row(SheetGrid, 0).ActualHeight, 69) && _document.GetCellInput(0, 2, 2) == "resize draft", "Undo resize preserves the separately committed cell draft");
        await UndoAsync(false); await Idle();
        check(_document.GetCellInput(0, 2, 2) == "", "The preceding undo restores the cell draft independently");

        var sheetId = _document.Snapshot.Sheets[0].Id;
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(sheetId, SheetAxis.Column, 1, 1, hidden: true)); await Idle();
        grip = ColumnGrip(2, "Left"); Start(grip); await Delta(grip, 7, 0); End(grip); await Idle();
        check(Near(SheetGrid.Columns[0].ActualWidth, 223) && _document.Snapshot.Sheets[0].Columns[1].Hidden, "Left grip skips hidden columns and preserves their metadata");
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(sheetId, SheetAxis.Column, 1, 1, hidden: false)); await Idle();

        ApplyValues([(0, 0, "Header"), (1, 0, "keep"), (2, 0, "hide")]);
        ApplyFilter(new(0, 0, 2, 0), 0, "keep"); await Idle();
        grip = RowGrip(SheetGrid, 1); Start(grip); await Delta(grip, 0, 10); End(grip); await Idle();
        check(CurrentView.FilterRange is not null && Near(Row(SheetGrid, 1).ActualHeight, 41) && Row(SheetGrid, 2).Height == 0,
            "Visible row resize preserves filters and hidden row geometry");
        ClearFilter_Click(this, new()); await Idle();

        grip = ColumnGrip(1); Start(grip); await Delta(grip, -1000, 0); End(grip); await Idle();
        check(Near(SheetGrid.Columns[1].ActualWidth, 12) && _document.Snapshot.Sheets[0].Columns[1].Width == 1, "Narrow columns round-trip below the former 55 pixel display floor");
        GridSurface.Select(new(1, 1)); Freeze_Click(new MenuItem(), new()); await Idle();
        check(Near(FrozenGrid.Columns[1].ActualWidth, 12), "Frozen pane uses the same minimum column width");
        Freeze_Click(new MenuItem { Tag = "clear" }, new()); await Idle();
        grip = ColumnGrip(1); Start(grip); await Delta(grip, 1000, 0); End(grip); await Idle();
        check(Near(SheetGrid.Columns[1].ActualWidth, 400), "Column resizing clamps to supported maximum");
        grip = RowGrip(SheetGrid, 0); Start(grip); await Delta(grip, 0, 1000); End(grip); await Idle();
        check(Near(Row(SheetGrid, 0).ActualHeight, 400), "Row resizing clamps to supported maximum");
        grip = RowGrip(SheetGrid, 0); Start(grip); await Delta(grip, 0, -1000); End(grip); await Idle();
        check(Near(Row(SheetGrid, 0).ActualHeight, 1), "Row resizing clamps to supported minimum");

        InstallDocument(WorkbookDocument.CreateBlank(_services), null, "尺寸重绘回归.xlsq");
        ApplyValues([(0, 0, "=1+2")]); await RecalculateAsync();
        GridSurface.Select(new(2, 2)); Freeze_Click(new MenuItem(), new()); await Idle();
        var frozenColumn = FrozenGrid.Columns[0];
        var frozenItems = FrozenGrid.ItemsSource;
        var contentRefreshes = 0; var frozenResets = 0; var calculations = 0; var disabled = 0; var widthWrites = 0;
        void RowChanged(object? sender, PropertyChangedEventArgs e) => contentRefreshes++;
        void ColumnsChanged(object? sender, NotifyCollectionChangedEventArgs e) { if (e.Action == NotifyCollectionChangedAction.Reset) frozenResets++; }
        void DocumentChanged(object? sender, DocumentChange e) { if (e.CalculationOnly) calculations++; }
        void EnabledChanged(object sender, DependencyPropertyChangedEventArgs e) { if (Equals(e.NewValue, false)) disabled++; }
        void WidthChanged(object? sender, EventArgs e) => widthWrites++;
        var widthProperty = DependencyPropertyDescriptor.FromProperty(DataGridColumn.WidthProperty, typeof(DataGridColumn));
        _rows[0].PropertyChanged += RowChanged;
        ((INotifyCollectionChanged)FrozenGrid.Columns).CollectionChanged += ColumnsChanged;
        _document.Session.Changed += DocumentChanged;
        Workspace.IsEnabledChanged += EnabledChanged;
        widthProperty.AddValueChanged(SheetGrid.Columns[0], WidthChanged);
        _suppressAutoCalculation = false;
        try
        {
            grip = ColumnGrip(0); Start(grip);
            for (var i = 0; i < 8; i++) grip.RaiseEvent(new DragDeltaEventArgs(2, 0) { RoutedEvent = Thumb.DragDeltaEvent });
            check(widthWrites == 0, "Native column resize is suppressed and queued pointer moves do not repeatedly write widths");
            await Idle();
            check(widthWrites == 1 && Near(SheetGrid.Columns[0].ActualWidth, 176), "A burst of pointer moves produces one layout preview with the final width");
            check(Near(FrozenGrid.Columns[0].ActualWidth, 176), "Frozen column widths follow the live preview");
            End(grip); await Idle();
            check(calculations == 0 && disabled == 0 && !_calculationPending, "Pure resize never starts calculation or disables the workspace");
            check(ReferenceEquals(frozenColumn, FrozenGrid.Columns[0]) && ReferenceEquals(frozenItems, FrozenGrid.ItemsSource) && frozenResets == 0,
                "Resize retains frozen column and item identities without rebuilding the pane");
            check(contentRefreshes == 0, "Resize does not invalidate cell text and style bindings");
            var cache = _document.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation!;
            check(cache.State == CalculationState.Fresh && cache.InputRevision == _document.Snapshot.Revision, "Resize retains valid formula results at the current revision");

            grip = RowGrip(SheetGrid, 0); Start(grip);
            grip.RaiseEvent(new DragDeltaEventArgs(0, 10) { RoutedEvent = Thumb.DragDeltaEvent });
            check(VisualChild<DataGridCellsPresenter>(Row(SheetGrid, 0))!.ReadLocalValue(HeightProperty) == DependencyProperty.UnsetValue,
                "Native row resize does not compete with the row container height");
            await Idle();
            check(Near(Row(SheetGrid, 0).ActualHeight, 41) && Near(Row(FrozenGrid, 0).ActualHeight, 41), "Row preview updates both panes without resetting them");
            End(grip); await Idle();
            check(calculations == 0 && disabled == 0 && frozenResets == 0, "Row completion also avoids recalculation and pane rebuilds");
            revision = _document.Snapshot.Revision;
            grip = RowGrip(SheetGrid, 0); Start(grip);
            grip.RaiseEvent(new DragDeltaEventArgs(0, 20) { RoutedEvent = Thumb.DragDeltaEvent });
            End(grip, true); await Idle();
            check(Near(Row(SheetGrid, 0).ActualHeight, 41) && _document.Snapshot.Revision == revision, "Cancel discards an unrendered preview without a late layout jump");
        }
        finally
        {
            _suppressAutoCalculation = true;
            _rows[0].PropertyChanged -= RowChanged;
            ((INotifyCollectionChanged)FrozenGrid.Columns).CollectionChanged -= ColumnsChanged;
            _document.Session.Changed -= DocumentChanged;
            Workspace.IsEnabledChanged -= EnabledChanged;
            widthProperty.RemoveValueChanged(SheetGrid.Columns[0], WidthChanged);
        }
    }
}
