using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunScrollRangeChecksAsync(string directory, Action<bool, string> check)
    {
        var first = new WorksheetSnapshot(SheetId.New(), "Scroll");
        var second = new WorksheetSnapshot(SheetId.New(), "Other");
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [first, second]);
        var path = Path.Combine(directory, "scroll-range.xlsq");
        new TypedWorkbookStorage().Save(path, snapshot); await OpenPathAsync(path);
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        var layout = (global::Avalonia.Controls.Grid)_grid.Content!;
        var canvas = (Canvas)layout.Children[0];
        var vertical = (ScrollBar)layout.Children[1]; var horizontal = (ScrollBar)layout.Children[2];
        (double X, double Y) Maximum() => (horizontal.Maximum, vertical.Maximum);
        using var pointer = new Pointer(128, PointerType.Mouse, true);
        void Wheel(Vector delta, KeyModifiers modifiers = KeyModifiers.None)
        {
            var point = canvas.TranslatePoint(new(200, 120), this)!.Value;
            canvas.RaiseEvent(new PointerWheelEventArgs(canvas, pointer, this, point, 1,
                new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.Other), modifiers, delta));
        }
        void Bottom()
        {
            horizontal.Value = horizontal.Maximum; vertical.Value = vertical.Maximum;
        }

        check(Maximum() == (0, 0), "Empty sheet scrollbars cover only the viewport");
        _grid.ScrollTo(double.MaxValue, double.MaxValue); Wheel(new(-100, -100)); Wheel(new(0, -100), KeyModifiers.Shift);
        check(Maximum() == (0, 0) && _grid.ScrollOffset == (0, 0), "Wheel, touchpad deltas and direct scrolling cannot expand an empty range");
        _grid.SelectRange(new(0, 0, WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1), new(0, 0));
        check(Maximum() == (0, 0), "Select-all does not expose the worksheet limits");
        _grid.SelectRange(new(90, 0, 100, WorkbookLimits.MaxColumns - 1), new(90, 0));
        check(vertical.Maximum > 0 && horizontal.Maximum == 0, "Full-row selection only expands vertical scrolling");
        _grid.SelectRange(new(0, 40, WorkbookLimits.MaxRows - 1, 50), new(0, 40));
        var selected = Maximum(); check(selected.X > 0 && selected.Y > 0, "Full-column selection expands horizontal scrolling");
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        var verticalThumb = vertical.GetVisualDescendants().OfType<Thumb>().Single();
        var horizontalThumb = horizontal.GetVisualDescendants().OfType<Thumb>().Single();
        var thickness = (verticalThumb.Bounds.Width, horizontalThumb.Bounds.Height);
        check(vertical.IsExpanded && horizontal.IsExpanded && thickness.Width > 2 && thickness.Height > 2,
            "Main grid scrollbars show expanded thumbs without hover");
        await Task.Delay(TimeSpan.FromMilliseconds(Math.Max(vertical.HideDelay.TotalMilliseconds, horizontal.HideDelay.TotalMilliseconds) + 100));
        check(vertical.IsExpanded && horizontal.IsExpanded && thickness == (verticalThumb.Bounds.Width, horizontalThumb.Bounds.Height),
            "Idle timeout leaves both main grid scrollbar thumbs at full thickness");
        _grid.Select(new(0, 0));
        check(Maximum() == selected && !Document.IsDirty, "Returning selection preserves the range without changing the document");
        for (var i = 0; i < 5; i++) { Bottom(); Wheel(new(-100, -100)); }
        check(Maximum() == selected && _grid.ScrollOffset == selected, "Repeated scrollbar and wheel end stops never extend the range");

        _grid.SelectRanges([new(0, 0, 0, 0), new(200, 80, 220, 90)], new(0, 0));
        var disjoint = Maximum();
        check(disjoint.X > selected.X && disjoint.Y > selected.Y, "Offscreen disjoint range expands both axes even with A1 active");
        _workspace.Editor.SetCell(0, 400, 140, "distant"); Synchronize();
        var edited = Maximum();
        check(edited.X > disjoint.X && edited.Y > disjoint.Y && _grid.Current == new CellPosition(0, 0), "Editing outside the selection expands the used range");
        await History(false);
        check(Maximum() == edited, "Undoing a distant edit does not shrink the range");
        _workspace.Editor.SetCell(0, 500, 160, "clear me"); Synchronize(); var beforeClear = Maximum();
        _workspace.Editor.SetCell(0, 500, 160, ""); Synchronize();
        check(Maximum() == beforeClear, "Clearing the last used cell does not shrink the range");
        _grid.Select(new(0, 0)); _grid.ScrollTo(0, 0);
        Document.Session.UpdateNavigation(second.Id, new(0, 0)); Synchronize();
        check(Maximum() == (0, 0), "Another sheet has an independent scroll range");
        Document.Session.UpdateNavigation(first.Id, new(0, 0)); Synchronize();
        check(Maximum() == beforeClear, "Switching back retains the session scroll range");
        // Reopen the same file/document identity in a fresh DocumentSession.
        await OpenPathAsync(path);
        check(Maximum() == (0, 0), "Reopening the same document resets transient range growth");

        _grid.Select(new(100, 30));
        check(_grid.GetCellBounds(new(100, 30)) is not null && _grid.ScrollOffset.X > 0 && _grid.ScrollOffset.Y > 0,
            "Cell navigation beyond the bounds grows the range and reveals its target");
        _grid.CellEditor.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Down });
        check(_grid.Current.Row == 101 && _grid.GetCellBounds(_grid.Current) is not null, "Keyboard navigation crosses the current bottom edge");
        var beforeDrag = Maximum();
        _grid.IsHitTestVisible = false;
        try { _grid.BeginBenchmarkDrag(new(canvas.Bounds.Width + 100, canvas.Bounds.Height + 100)); await Task.Delay(200); }
        finally { _grid.CancelInputGesture(); _grid.IsHitTestVisible = true; }
        check(Maximum().X > beforeDrag.X && Maximum().Y > beforeDrag.Y, "Selection drag outside the viewport grows both bounds");
        var afterDrag = Maximum(); Bottom(); Wheel(new(-100, -100));
        check(Maximum() == afterDrag, "Scrolling after a drag uses the new fixed bounds");

        _grid.ScrollTo(0, 0); _grid.Select(new(0, 0));
        await FreezeAsync(2, 2); await ZoomAsync(150);
        Bottom(); var frozen = Maximum();
        check(_grid.ScrollOffset == frozen && frozen.X > 0 && frozen.Y > 0, "Frozen panes and zoom retain reachable bounded endpoints");
        Wheel(new(-100, -100)); check(Maximum() == frozen, "Frozen and zoomed wheel scrolling cannot grow the range");

        // File used extents, merges and saved viewport are also valid initial bounds.
        var stored = new WorksheetSnapshot(SheetId.New(), "Stored", new Dictionary<CellPosition, CellSnapshot>
        { [new(200, 60)] = new(CellInput.Parse("used")) }, mergedCells: ["A1:CA250"],
            view: new WorksheetView { TopLeftCell = new(300, 100) });
        var storedPath = Path.Combine(directory, "scroll-stored.xlsq");
        var storedSnapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [stored]);
        new TypedWorkbookStorage().Save(storedPath, storedSnapshot);
        await OpenPathAsync(storedPath);
        check(_grid.ScrollingCell == new CellPosition(300, 100), "Opening restores a saved viewport beyond the used range");
        var restored = Maximum(); Bottom(); Wheel(new(-100, -100));
        check(Maximum() == restored, "Restored viewport does not cause continuing scroll growth");
        using var unarranged = new QuickExcel.Grid.Avalonia.WorkbookGrid();
        unarranged.Bind(storedSnapshot, 0, WorkbookLimits.MaxRows, WorkbookLimits.MaxColumns);
        var host = new Window { Width = 900, Height = 600, ShowInTaskbar = false, Content = unarranged };
        try
        {
            host.Show(this); await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            check(unarranged.ScrollingCell == new CellPosition(300, 100), "Binding before first layout preserves the saved viewport");
        }
        finally { host.Close(); }
    }
}
