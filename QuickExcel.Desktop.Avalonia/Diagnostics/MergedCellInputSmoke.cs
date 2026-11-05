using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunMergedCellInputChecksAsync(Action<bool, string> check)
    {
        var canvas = (Canvas)((global::Avalonia.Controls.Grid)_grid.Content!).Children[0];
        using var pointer = new Pointer(127, PointerType.Mouse, true);
        ulong timestamp = 0;
        Point Local(Point point) => canvas.TranslatePoint(point, this)!.Value;
        void Down(Point point, KeyModifiers modifiers = KeyModifiers.None, int clicks = 1, bool right = false) =>
            canvas.RaiseEvent(new PointerPressedEventArgs(canvas, pointer, this, Local(point), ++timestamp,
                new PointerPointProperties(right ? RawInputModifiers.RightMouseButton : RawInputModifiers.LeftMouseButton,
                    right ? PointerUpdateKind.RightButtonPressed : PointerUpdateKind.LeftButtonPressed), modifiers, clicks));
        void Move(Point point) => canvas.RaiseEvent(new PointerEventArgs(InputElement.PointerMovedEvent, canvas, pointer, this, Local(point), ++timestamp,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other), KeyModifiers.None));
        void Up(Point point, bool right = false) => canvas.RaiseEvent(new PointerReleasedEventArgs(canvas, pointer, this, Local(point), ++timestamp,
            new PointerPointProperties(RawInputModifiers.None, right ? PointerUpdateKind.RightButtonReleased : PointerUpdateKind.LeftButtonReleased),
            KeyModifiers.None, right ? MouseButton.Right : MouseButton.Left));
        void CheckAnchor(string label) => check(_grid.Current == new CellPosition(0, 0) &&
            _grid.CellEditor.Text == "合并标题" && _formula.Text == "合并标题", label);

        // These points cover all four fragments of A1:C2 across both frozen axes.
        foreach (var point in new[] { new Point(75, 40), new Point(235, 40), new Point(75, 68), new Point(235, 68) })
        {
            _grid.Select(new(3, 3)); _grid.ScrollTo(40, 15);
            Down(point); Up(point);
            CheckAnchor($"Merged pointer click at {point} selects and edits the anchor");
            check(_grid.SelectionRanges.Single() == new GridRange(0, 0, 1, 2), "Pointer selection contains the entire merge");
        }
        var child = new Point(235, 68);
        _grid.Select(new(3, 3)); _grid.ScrollTo(40, 15);
        Down(child, right: true); Up(child, right: true);
        CheckAnchor("Right-click on a merged child selects its anchor");

        _grid.Select(new(3, 3)); _grid.ScrollTo(40, 15);
        Down(child, KeyModifiers.Shift); Up(child);
        CheckAnchor("Shift-click on a merged child keeps the anchor as current");
        check(_grid.SelectionRanges.Single() == new GridRange(0, 0, 3, 3), "Shift-click expands through the entire merge");

        _grid.Select(new(3, 3)); _grid.ScrollTo(40, 15);
        Down(child, KeyModifiers.Control); Up(child);
        CheckAnchor("Control-click on a merged child keeps the anchor as current");
        check(_grid.SelectionRanges.Count == 2, "Control-click retains the previous disjoint selection");

        _grid.Select(new(3, 3)); _grid.ScrollTo(40, 15);
        var start = _grid.GetCellBounds(new(3, 3))!.Value.Center;
        Down(start); Move(child);
        var changes = 0;
        void SelectionChanged() => changes++;
        _grid.SelectionChanged += SelectionChanged;
        try { Move(new(75, 68)); Move(child); }
        finally { _grid.SelectionChanged -= SelectionChanged; }
        check(changes == 0, "Moving within the same merged selection does not repeat selection notifications");
        Up(child);
        CheckAnchor("Dragging into a merged child keeps the anchor as current");
        check(_grid.SelectionRanges.Single() == new GridRange(0, 0, 3, 3), "Dragging includes the whole merge and original cell");

        Down(child, clicks: 2); Up(child);
        check(_grid.IsEditing && _grid.Current == new CellPosition(0, 0), "Double-click edits the merged anchor");
        _grid.CellEditor.Text = "合并\t标题";
        check(await _grid.CommitAsync() && Document.GetCellInput(0, 0, 0) == "合并\t标题" &&
            Document.Snapshot.Sheets[0].GetCell(1, 1) is null, "Merged edit commits original tab text only to the anchor");
        await History(false);
        check(Document.GetCellInput(0, 0, 0) == "合并标题", "Merged edit is one undo entry");

        await ZoomAsync(150); _grid.ScrollTo(40, 15);
        var zoomed = new Point(child.X * 1.5, child.Y * 1.5);
        Down(zoomed); Up(zoomed);
        CheckAnchor("Scrolled frozen merged click resolves its anchor at 150 percent zoom");
        await History(false); _grid.ScrollTo(40, 15);

        // Header operations intentionally retain the clicked row/column boundary.
        var rowHeader = new Point(20, 68);
        Down(rowHeader, right: true); Up(rowHeader, right: true);
        check(_grid.Current.Row == 1 && !_grid.IsContextSelectionBlocked, "Merged row header preserves its clicked boundary");
        var columnHeader = new Point(235, 18);
        Down(columnHeader, right: true); Up(columnHeader, right: true);
        check(_grid.Current.Column == 1 && !_grid.IsContextSelectionBlocked, "Merged column header preserves its clicked boundary");

        var id = Document.Snapshot.Sheets[0].Id;
        await ExecuteAsync(new SetCellProtectionCommand(id, new(0, 0, 0, 0), false));
        _grid.Select(new(3, 3));
        await ExecuteAsync(new SetProtectionCommand(id, new() { SelectLockedCells = true }));
        _grid.ScrollTo(40, 15); Down(child); Up(child);
        CheckAnchor("Merged click uses the anchor's selection protection instead of the covered child's style");
        await History(false); await History(false);
        _grid.Select(new(0, 0));

        foreach (var (first, last) in new[]
        {
            (new CellPosition(0, 0), new CellPosition(1, 2)),
            (new CellPosition(1, 0), new CellPosition(0, 2))
        })
        {
            await MergeAsync(remove: true);
            var from = _grid.GetCellBounds(first)!.Value.Center;
            var to = _grid.GetCellBounds(last)!.Value.Center;
            Down(from); Move(to); Up(to);
            check(_grid.Current == last, "The drag endpoint is active before merging");
            await MergeAsync(remove: false);
            CheckAnchor($"Merging a drag from {first} to {last} immediately targets the anchor");
            check(_address.Text == "A1" && Document.Snapshot.GetSheet(id).View.ActiveCell == new CellPosition(0, 0) &&
                _grid.SelectionRanges.Single() == new GridRange(0, 0, 1, 2), "Merge updates the active address and retains the full selection");
        }
    }
}
