using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunFormulaReferenceChecksAsync(Action<bool, string> check)
    {
        // These gestures are routed below, not injected through the OS pointer.
        // Keep a stationary real mouse from setting CellEditor.IsPointerOver and intercepting them.
        var hitTestVisible = _grid.IsHitTestVisible;
        _grid.IsHitTestVisible = false;
        try { await RunFormulaReferenceGesturesAsync(check); }
        finally { _grid.IsHitTestVisible = hitTestVisible; }
    }

    private async Task RunFormulaReferenceGesturesAsync(Action<bool, string> check)
    {
        var canvas = (Canvas)((global::Avalonia.Controls.Grid)_grid.Content!).Children[0];
        using var pointer = new Pointer(126, PointerType.Mouse, true);
        ulong timestamp = 0;
        Point PointFor(CellPosition cell) => canvas.TranslatePoint(_grid.GetCellBounds(cell)!.Value.Center, this)!.Value;
        void Move(CellPosition cell)
        {
            canvas.RaiseEvent(new PointerEventArgs(InputElement.PointerMovedEvent, canvas, pointer, this, PointFor(cell), ++timestamp,
                new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other), KeyModifiers.None));
        }
        void Down(CellPosition cell)
        {
            canvas.RaiseEvent(new PointerPressedEventArgs(canvas, pointer, this, PointFor(cell), ++timestamp,
                new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed), KeyModifiers.None, 1));
        }
        void Up(CellPosition cell) => canvas.RaiseEvent(new PointerReleasedEventArgs(canvas, pointer, this, PointFor(cell), ++timestamp,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased), KeyModifiers.None, MouseButton.Left));

        _grid.ScrollTo(0, 0);
        _grid.Select(new(4, 3));
        _formula.Focus(); _formula.Text = "=SUM()";
        _formula.CaretIndex = _formula.SelectionStart = _formula.SelectionEnd = 5;
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        var before = Document.Snapshot;
        Down(new(5, 0));
        check(_formula.Text == "=SUM(A6)" && _formulaGesture is not null && _grid.Current == new CellPosition(4, 3),
            "Formula-bar pointer press retains its draft after grid focus changes");
        Move(new(7, 1));
        Up(new(7, 1));
        check(_formula.Text == "=SUM(A6:B8)" && _formula.IsKeyboardFocusWithin && _formulaGesture is null && !_grid.IsSelectionDragging,
            "Formula-bar pointer drag replaces one reference and restores text focus on release");
        check(ReferenceEquals(before, Document.Snapshot), "Formula reference picking never commits an unfinished draft");
        _formula.CaretIndex = _formula.SelectionStart = _formula.SelectionEnd = _formula.Text!.Length;
        _formula.Text += "+";
        _formula.CaretIndex = _formula.SelectionStart = _formula.SelectionEnd = _formula.Text.Length;
        Down(new(5, 1)); Up(new(5, 1));
        check(_formula.Text == "=SUM(A6:B8)+B6" && _grid.Current == new CellPosition(4, 3),
            "Typing another operand after a drag allows another pointer reference");
        check(await CommitFormulaAsync() && Document.GetCellInput(0, 4, 3) == "=SUM(A6:B8)+B6",
            "Pointer-picked formula commits only to its original cell");
        await History(false);
        check(Document.GetCellInput(0, 4, 3) == "", "All pointer-picked references commit as one undo entry");

        _grid.Select(new(4, 4)); _grid.BeginEdit();
        _grid.CellEditor.Text = "=SUM()";
        _grid.CellEditor.CaretIndex = _grid.CellEditor.SelectionStart = _grid.CellEditor.SelectionEnd = 5;
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        before = Document.Snapshot;
        Down(new(5, 0));
        Move(new(7, 1)); Up(new(7, 1));
        check(_grid.CellEditor.Text == "=SUM(A6:B8)" && _grid.IsEditing && _grid.CellEditor.IsKeyboardFocusWithin &&
            _grid.Current == new CellPosition(4, 4) && ReferenceEquals(before, Document.Snapshot),
            $"In-cell pointer drag preserves the editor, original owner and uncommitted draft (text={_grid.CellEditor.Text}, editing={_grid.IsEditing}, focus={_grid.CellEditor.IsKeyboardFocusWithin}, current={_grid.Current}, sameSnapshot={ReferenceEquals(before, Document.Snapshot)}, pointerOver={_grid.CellEditor.IsPointerOver})");
        _grid.Cancel();
        check(ReferenceEquals(before, Document.Snapshot), "Cancelling a pointer-picked formula leaves the workbook unchanged");

        _grid.Select(new(4, 3)); _formula.Focus(); _formula.Text = "plain text";
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        Down(new(5, 0)); Up(new(5, 0));
        check(_grid.Current == new CellPosition(5, 0) && Document.GetCellInput(0, 4, 3) == "plain text" && _formulaGesture is null,
            "Ordinary text still commits and navigates on a grid click");
        await History(false);
    }
}
