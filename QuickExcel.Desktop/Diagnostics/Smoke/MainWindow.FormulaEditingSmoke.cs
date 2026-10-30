using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunFormulaEditingSmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>(); _suppressAutoCalculation = true;
        void Check(bool ok, string message) { if (!ok) throw new Exception(message); checks.Add(message); }
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        async Task KeyAsync(Key key)
        {
            var source = Keyboard.FocusedElement as UIElement ?? SheetGrid;
            var args = new KeyEventArgs(Keyboard.PrimaryDevice, PresentationSource.FromVisual(this), 0, key) { RoutedEvent = Keyboard.PreviewKeyDownEvent };
            source.RaiseEvent(args); Check(args.Handled, "Routed " + key + " reaches the formula/grid handler"); await _gridInput;
        }
        var sheet = new WorksheetSnapshot(SheetId.New(), "Entry", [
            new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(1)))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromNumber(2)))),
            new(new(2, 0), new(CellContent.FromLiteral(CellValue.FromNumber(3)))),
            new(new(4, 0), new(CellContent.FromLiteral(CellValue.FromNumber(5))))]);
        try
        {
            InstallDocument(new WorkbookDocument(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]), _services), null, "Formula entry smoke"); await Idle();
            GridSurface.Select(new(3, 0)); var revision = _document.Snapshot.Revision; BeginAutoSum();
            Check(_grid.Editor.Text == "=SUM(A1:A3)" && _document.Snapshot.Revision == revision, "AutoSum proposes the contiguous block above without committing");
            await KeyAsync(Key.Enter);
            Check(_document.Snapshot.Sheets[0].GetCell(3, 0)!.Value.NumberValue == 6, "Enter commits and calculates the suggested SUM");
            await UndoAsync(false); Check(_document.Snapshot.Sheets[0].GetCell(3, 0) is null && !_document.CanUndo, "AutoSum is one undoable formula edit");
            await UndoAsync(true); Check(_document.Snapshot.Sheets[0].GetCell(3, 0)!.Value.NumberValue == 6, "Redo restores the complete AutoSum formula and calculation");
            await UndoAsync(false);
            revision = _document.Snapshot.Revision;
            GridSurface.Select(new(0, 2)); GridSurface.BeginEdit(); _grid.Editor.Text = "=av"; _grid.Editor.CaretIndex = 3; await Idle();
            Check(_formulaHelp.IsOpen && _completion?.Names.Contains("AVERAGE") == true, "Actual completion popup offers registered functions");
            await KeyAsync(Key.Tab);
            Check(_grid.Editor.Text == "=AVERAGE(" && _grid.IsEditing && _document.Snapshot.Revision == revision, "Tab accepts completion without navigating or committing");
            _grid.Editor.Text = "=su(A1)"; _grid.Editor.CaretIndex = 3; await KeyAsync(Key.Tab);
            Check(_grid.Editor.Text == "=SUM(A1)" && _grid.Editor.CaretIndex == 5, "Completion reuses an existing parenthesis and places the caret inside its arguments");
            _grid.Editor.Text = "=IF(A1>0,SUM(A1,"; _grid.Editor.CaretIndex = _grid.Editor.Text.Length; await Idle();
            Check(_formulaHint.Text.Contains("SUM(") && _formulaHint.Text.Contains("第 2"), "Actual popup displays the innermost function and argument index");
            CaptureGridEvidence(Path.Combine(folder, "formula-hint.png"));
            var visual = (FrameworkElement)_formulaHelp.Child; visual.UpdateLayout();
            var bitmap = new RenderTargetBitmap((int)Math.Ceiling(visual.ActualWidth), (int)Math.Ceiling(visual.ActualHeight), 96, 96, PixelFormats.Pbgra32); bitmap.Render(visual);
            var png = new PngBitmapEncoder(); png.Frames.Add(BitmapFrame.Create(bitmap));
            using (var stream = File.Create(Path.Combine(folder, "formula-hint-popup.png"))) png.Save(stream);
            await KeyAsync(Key.Escape); Check(_grid.IsEditing && !_formulaHelp.IsOpen, "First Escape dismisses formula help while keeping the draft");
            await KeyAsync(Key.Escape); Check(!_grid.IsEditing && _document.Snapshot.Revision == revision, "Second Escape cancels the draft without history");
            GridSurface.Select(new(0, 2)); GridSurface.BeginEdit(); _grid.Editor.Text = "=SUM("; _grid.Editor.CaretIndex = 5;
            Check(BeginReferenceGesture(new(0, 0)), "Grid formula draft starts a reference gesture");
            UpdateReferenceGesture(new(2, 0)); EndReferenceGesture();
            Check(_grid.Editor.Text == "=SUM(A1:A3" && _grid.Current == new CellPosition(0, 2) && _document.Snapshot.Revision == revision,
                "Reference dragging changes only the original cell draft and leaves model selection unchanged");
            _grid.Editor.Select(_grid.Editor.Text.Length, 0); _grid.Editor.SelectedText = ")"; await KeyAsync(Key.Enter);
            Check(_document.Snapshot.Sheets[0].GetCell(0, 2)!.Value.NumberValue == 6 && _document.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue == 1,
                "Pointed formula commits to C1, leaving the referenced cells unchanged");
            GridSurface.Select(new(1, 2)); FormulaBar.Focus(); FormulaBar.Text = "=SUM("; FormulaBar.CaretIndex = 5;
            revision = _document.Snapshot.Revision;
            Check(BeginReferenceGesture(new(0, 0)), "Formula bar starts a reference gesture"); UpdateReferenceGesture(new(1, 0)); EndReferenceGesture();
            FormulaBar.Select(FormulaBar.Text.Length, 0); FormulaBar.SelectedText = ")";
            Check(_formulaRow == 1 && _formulaColumn == 2 && _document.Snapshot.Revision == revision, "Formula-bar draft keeps its original address through reference picking");
            await ApplyFormulaBarAsync(); Check(_document.Snapshot.Sheets[0].GetCell(1, 2)!.Value.NumberValue == 3, "Formula-bar apply uses the ordinary undoable commit path");
            GridSurface.Select(new(2, 2)); GridSurface.BeginEdit(); _grid.Editor.Text = "=A1"; _grid.Editor.CaretIndex = 3;
            await KeyAsync(Key.F4); Check(_grid.Editor.Text == "=$A$1", "F4 changes the active draft reference to absolute");
            _grid.Editor.Select(_grid.Editor.Text.Length, 0); _grid.Editor.SelectedText = "+$A2"; await KeyAsync(Key.Enter);
            Check(_document.Snapshot.Sheets[0].GetCell(2, 2)!.Value.NumberValue == 3, "F4 output calculates through the standard evaluator");
            GridSurface.Select(new(8, 2)); GridSurface.BeginEdit(); _grid.Editor.Text = "=A1"; _grid.Editor.CaretIndex = 3; revision = _document.Snapshot.Revision;
            _editState.StartComposition(); BeginAutoSum();
            Check(!HandleFormulaAssistKey(_grid.Editor, Key.F4, ModifierKeys.None) && !BeginReferenceGesture(new(1, 0)) && _grid.Editor.Text == "=A1" && _document.Snapshot.Revision == revision,
                "IME composition gates AutoSum, F4 and reference picking"); _editState.EndComposition(); GridSurface.Cancel(); CloseFormulaHelp();
            GridSurface.Select(new(0, 0)); await HandleGridKeyAsync(Key.Down, ModifierKeys.Control, false);
            Check(_grid.Current == new CellPosition(2, 0), "Ctrl+Down stops at the end of a contiguous data block");
            await HandleGridKeyAsync(Key.Down, ModifierKeys.Control, false); Check(_grid.Current == new CellPosition(4, 0), "Ctrl+Down crosses a blank gap to the next nonempty cell");
            await HandleGridKeyAsync(Key.Up, ModifierKeys.Control, false); Check(_grid.Current == new CellPosition(2, 0), "Ctrl+Up crosses the gap in reverse");
            await HandleGridKeyAsync(Key.Home, ModifierKeys.Control, false); await HandleGridKeyAsync(Key.Down, ModifierKeys.Control | ModifierKeys.Shift, false);
            Check(_grid.SelectionRanges.Single() == new GridRange(0, 0, 2, 0), "Ctrl+Shift+Down extends from the stable original anchor");
            await HandleGridKeyAsync(Key.End, ModifierKeys.Control, false); Check(_grid.Current == new CellPosition(4, 2), "Ctrl+End goes to the stored used-area corner");
            await HandleGridKeyAsync(Key.Home, ModifierKeys.Control, false); await HandleGridKeyAsync(Key.End, ModifierKeys.None, false);
            Check(_endNavigation && _grid.Current == new CellPosition(0, 0), "End enters navigation mode without moving or editing");
            await HandleGridKeyAsync(Key.Down, ModifierKeys.None, false); Check(!_endNavigation && _grid.Current == new CellPosition(2, 0), "End then Down follows the same data boundary");
            GridSurface.Select(new(4, 0)); await HandleGridKeyAsync(Key.Down, ModifierKeys.Control | ModifierKeys.Shift, false);
            Check(_grid.Current == new CellPosition(WorkbookLimits.MaxRows - 1, 0) && _grid.SelectionRanges.Single() == new GridRange(4, 0, WorkbookLimits.MaxRows - 1, 0),
                "An empty tail expands the virtual grid to its declared row limit while retaining the shift anchor");
            await HandleGridKeyAsync(Key.Home, ModifierKeys.Control, false);
            Check(_document.Snapshot.Revision == revision, "Navigation does not create document history");
            var path = Path.Combine(folder, "entry.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, _document.Snapshot);
            Check(storage.Open(path).Sheets[0].GetCell(0, 2)!.Value.NumberValue == 6, "Assisted formulas survive normal XLSX save and reopen");
            File.WriteAllLines(Path.Combine(folder, "formula-entry-result.txt"), new[] { $"PASS: {checks.Count} checks" }.Concat(checks));
        }
        finally { _suppressAutoCalculation = false; EndReferenceGesture(); CloseFormulaHelp(); }
    }
}
