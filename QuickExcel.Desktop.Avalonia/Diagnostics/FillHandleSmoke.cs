using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunFillHandleChecksAsync(Action<bool, string> check)
    {
        var canvas = (Canvas)((global::Avalonia.Controls.Grid)_grid.Content!).Children[0];
        using var pointer = new Pointer(137, PointerType.Mouse, true); ulong timestamp = 0;
        Point Local(Point point) => canvas.TranslatePoint(point, this)!.Value;
        void Down(Point point, int clicks = 1) => canvas.RaiseEvent(new PointerPressedEventArgs(canvas, pointer, this, Local(point), ++timestamp,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed), KeyModifiers.None, clicks));
        void Move(Point point) => canvas.RaiseEvent(new PointerEventArgs(InputElement.PointerMovedEvent, canvas, pointer, this, Local(point), ++timestamp,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other), KeyModifiers.None));
        void Up(Point point, KeyModifiers modifiers = KeyModifiers.None) => canvas.RaiseEvent(new PointerReleasedEventArgs(canvas, pointer, this, Local(point), ++timestamp,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased), modifiers, MouseButton.Left));
        async Task Settle()
        {
            if (_operation is { } operation) await operation;
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        }
        var id = Document.Snapshot.ActiveSheetId;
        await ExecuteAsync(new SetCellsCommand([
            new(new(id, 5, 3), CellInput.Parse("2")), new(new(id, 6, 3), CellInput.Parse("5")),
            new(new(id, 5, 4), CellInput.Parse("ID007")), new(new(id, 5, 5), CellInput.Parse("=D6+1")),
            new(new(id, 6, 4), CellInput.Parse("ID008")), new(new(id, 7, 4), CellInput.Parse("ID009")),
            new(new(id, 8, 4), CellInput.Parse("ID010")), new(new(id, 9, 5), CellInput.Parse("keep"))]));
        void Select(GridRange range) { _grid.SelectRange(range, new(range.Top, range.Left)); _grid.ScrollTo(0, 0); }
        Point Handle(int row, int column) => _grid.GetCellBounds(new(row, column))!.Value.BottomRight;
        Select(new(5, 3, 6, 3)); var start = Handle(6, 3); var end = _grid.GetCellBounds(new(8, 3))!.Value.Center;
        var before = Document.Snapshot;
        Down(start); Move(end);
        check(ReferenceEquals(before, Document.Snapshot), "Fill preview never writes the document");
        Up(end); await Settle();
        check(Document.GetCellValue(0, 8, 3) == "11" && _grid.SelectionRanges.Single() == new GridRange(5, 3, 8, 3), "Fill handle infers numeric step and selects the result");
        await History(false); check(Document.GetCellInput(0, 7, 3) == "", "Automatic fill is one undo entry");
        Select(new(5, 3, 6, 3)); Down(start); Move(end); Up(end, KeyModifiers.Control); await Settle();
        check(Document.GetCellValue(0, 7, 3) == "2" && Document.GetCellValue(0, 8, 3) == "5", "Ctrl on release forces repeated copy");
        await History(false);

        Select(new(5, 4, 5, 4)); start = Handle(5, 4); end = _grid.GetCellBounds(new(5, 6))!.Value.Center;
        Down(start); Move(end); Up(end); await Settle();
        check(Document.GetCellInput(0, 5, 6) == "ID009", "Horizontal fill extends a zero-padded text ID");
        await History(false);

        Select(new(5, 5, 5, 5)); start = Handle(5, 5);
        Down(start); Up(start); Down(start, 2); Up(start); await Settle();
        check(Document.GetCellInput(0, 8, 5) == "=(D9+1)" && Document.GetCellInput(0, 9, 5) == "keep", "Double-click fills formulas along adjacent data without overwriting the next cell");
        await History(false);

        Select(new(5, 3, 6, 3)); start = Handle(6, 3); end = _grid.GetCellBounds(new(8, 3))!.Value.Center;
        before = Document.Snapshot; Down(start); Move(end);
        _grid.CellEditor.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape });
        Up(end); await Settle();
        check(ReferenceEquals(before, Document.Snapshot), "Escape cancels autofill without changing history");

        // Exercise the same edge-scroll animation used by a captured fill, without claiming physical mouse input.
        Select(new(5, 3, 6, 3)); start = Handle(6, 3); Down(start);
        Move(new(canvas.Bounds.Width / 2, canvas.Bounds.Height + 80));
        for (var i = 0; i < 20 && _grid.ScrollOffset.Y == 0; i++) await Task.Delay(20);
        check(_grid.ScrollOffset.Y > 0 && ReferenceEquals(before, Document.Snapshot), "Fill handle scrolls beyond the viewport while keeping preview read-only");
        _grid.CancelInputGesture(); Up(start); _grid.ScrollTo(0, 0);
        await History(false); // Remove only this check's seed edit.
    }
}
