using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using FastWpfGrid;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunMergedSelectionSmokeAsync(string folder)
    {
        var cases = new List<string>();
        void Check(bool condition, string message)
        { if (!condition) throw new Exception("Merged selection: " + message); cases.Add(message); }
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(_grid.Control);
        }
        InstallDocument(WorkbookDocument.CreateBlank(_services), null, "合并选区回归.xlsq");
        var id = _document.Snapshot.Sheets[0].Id;
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Column, 0, 11, 10));
        var merge = new GridCellRange(0, 5, 6, 7);
        GridSurface.SelectRange(new(merge.Top, merge.Left, merge.Bottom, merge.Right), new(6, 7));
        ExecuteWorkbookCommand(new MergeCellsCommand(id, new(0, 5, 6, 7)));
        await Idle();
        var grid = _grid.Control;
        var overlay = (FrameworkElement)grid.FindName("selectionOverlay");
        void WholeMerge(string context)
        {
            Check(_grid.Current == new CellPosition(0, 5) && _selectedRow == 0 && _selectedColumn == 5,
                context + ": active address belongs to F1");
            Check(grid.GetSelectedModelRanges().SequenceEqual([merge]), context + ": selects F1:H7");
        }
        IEnumerable<GeometryDrawing> Shapes(Drawing drawing) => drawing is DrawingGroup group
            ? group.Children.SelectMany(Shapes) : drawing is GeometryDrawing geometry ? [geometry] : [];
        void CheckOutline(string context)
        {
            var expected = grid.GetModelRangeFragments(merge).Select(f => Rect.Intersect(f.Bounds, f.Clip)).ToArray();
            var outlined = Shapes(VisualTreeHelper.GetDrawing(overlay)).Where(d => d.Pen is not null)
                .Select(d => d.Geometry.Bounds).ToArray();
            Check(expected.Length > 0 && outlined.Length == expected.Length &&
                outlined.All(r => expected.Any(e => Math.Abs(e.Left - r.Left) < 1 && Math.Abs(e.Top - r.Top) < 1 &&
                    Math.Abs(e.Width - r.Width) < 1 && Math.Abs(e.Height - r.Height) < 1)),
                context + ": one outline per visible merge fragment, with no nested or duplicate border");
        }
        WholeMerge("Fresh merge"); CheckOutline("Fresh merge");
        foreach (var position in new[] { new CellPosition(0, 5), new CellPosition(3, 6), new CellPosition(6, 7) })
        {
            GridSurface.Select(position); await Idle(); WholeMerge("Programmatic " + position); CheckOutline("Programmatic " + position);
        }
        // Exercise the same hit-test, current-cell and add/remove paths used by mouse clicks.
        T Bind<T>(string name) where T : Delegate => typeof(FastGridControl)
            .GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic)!.CreateDelegate<T>(grid);
        var clear = Bind<Action>("ClearSelectedCells");
        var setCurrent = Bind<Action<FastGridCellAddress>>("SetCurrentCell");
        var add = Bind<Action<FastGridCellAddress>>("AddSelectedCell");
        var remove = Bind<Action<FastGridCellAddress>>("RemoveSelectedCell");
        var changed = Bind<Action<bool>>("OnChangeSelectedCells");
        var drag = Bind<Action<FastGridCellAddress, FastGridCellAddress>>("SetSelectedRectangle");
        var hitBounds = grid.GetModelCellBounds(4, 6)!.Value;
        var toReal = Bind<Func<FastGridCellAddress, FastGridCellAddress>>("ModelToReal");
        var hit = toReal(grid.HitTestModel(new(hitBounds.Left + 10, hitBounds.Top + 10)));
        clear(); setCurrent(hit); add(hit); changed(true); await Idle();
        WholeMerge("Interior mouse path"); CheckOutline("Interior mouse path");
        grid.CurrentModelCell = new(4, 6); await Idle(); WholeMerge("Native current-cell setter");
        remove(hit); changed(true); await Idle();
        Check(grid.SelectedCellCount == 0, "Ctrl removal removes the entire merge");
        add(hit); changed(true); await Idle(); WholeMerge("Ctrl addition");
        drag(new(7, 4), new(3, 6)); await Idle();
        Check(grid.GetSelectedModelRanges().SequenceEqual([new GridCellRange(0, 4, 7, 7)]), "Drag expands to include the complete intersected merge");
        GridSurface.Select(new(3, 6)); await Idle();
        GridSurface.BeginEdit(); _grid.Editor.Text = "merged value";
        var full = grid.GetModelRangeFragments(merge)[0].Bounds;
        Check(Math.Abs(_grid.Editor.Width - full.Width) < 1 && Math.Abs(_grid.Editor.Height - full.Height) < 1,
            "Native editor spans the merged bounds");
        Check(CommitEditors() && _document.GetCellInput(0, 0, 5) == "merged value" && _document.GetCellInput(0, 3, 6) == "",
            "Interior selection edits only the merge owner");
        await PasteSelectionAsync("pasted value", false); await Idle();
        Check(_document.GetCellInput(0, 0, 5) == "pasted value" && _document.Snapshot.Sheets[0].Cells.Count == 1,
            "Scalar paste into one merged cell writes only its owner");
        await UndoAsync(false); await Idle();
        Check(_document.GetCellInput(0, 0, 5) == "merged value", "Merged scalar paste remains one undoable transaction");
        foreach (var (key, modifiers, destination) in new[]
        {
            (Key.Tab, ModifierKeys.None, new CellPosition(0, 8)),
            (Key.Enter, ModifierKeys.None, new CellPosition(7, 5)),
            (Key.Right, ModifierKeys.None, new CellPosition(0, 8)),
            (Key.Down, ModifierKeys.None, new CellPosition(7, 5)),
            (Key.Tab, ModifierKeys.Shift, new CellPosition(0, 4))
        })
        {
            GridSurface.Select(new(0, 5)); await NavigateGridAsync(key, modifiers); await Idle();
            Check(_grid.Current == destination, $"{modifiers}+{key} leaves the merged cell");
        }
        GridSurface.Select(new(0, 5)); await Idle();
        CheckOutline("Before scrolling"); CaptureGridEvidence(Path.Combine(folder, "merged-selection.png"));
        SetSmokeFreeze(2, 6); RefreshSheet();
        GridSurface.Select(new(3, 6)); await Idle(); WholeMerge("Frozen panes"); CheckOutline("Frozen panes");
        grid.ScrollModelIntoView(new(12, 10)); await Idle(); CheckOutline("Scrolled frozen panes");
        SetSmokeFreeze(0, 0); RefreshSheet();
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Row, 2, 1, hidden: true));
        GridSurface.Select(new(3, 6)); await Idle(); CheckOutline("Hidden interior row");
        Check(grid.SelectedCellCount == 18, "Hidden row does not enter visible merged selection enumeration");
        await UndoAsync(false); await Idle();
        GridSurface.Select(new(0, 5)); await Idle();
        ExecuteWorkbookCommand(new MergeCellsCommand(id, new(0, 5, 6, 7), unmerge: true));
        GridSurface.Select(new(3, 6)); await Idle();
        Check(grid.SelectedCellCount == 1 && _grid.Current == new CellPosition(3, 6), "Unmerge restores ordinary selection");
        await UndoAsync(false); GridSurface.Select(new(3, 6)); await Idle(); WholeMerge("Undo unmerge"); CheckOutline("Undo unmerge");
        File.WriteAllLines(Path.Combine(folder, "merged-selection-cases.txt"), cases);
    }
}
