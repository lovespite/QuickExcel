using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunFillSmokeOnlyAsync(string folder)
    {
        Directory.CreateDirectory(folder); _suppressAutoCalculation = true;
        var cases = new List<string>();
        try
        {
            await RunFillSmokeAsync(folder, (ok, message) => { if (!ok) throw new Exception(message); cases.Add(message); });
            File.WriteAllLines(Path.Combine(folder, "fill-smoke-result.txt"), new[] { $"PASS: {cases.Count} checks" }.Concat(cases));
        }
        finally { _suppressAutoCalculation = false; }
    }
    private async Task RunFillSmokeAsync(string folder, Action<bool, string> check)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Fill", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(1), 1))),
            new(new(0, 1), new(CellContent.FromFormula("=A1+$F$1"))), new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromNumber(2)))),
            new(new(1, 1), new(CellContent.FromFormula("=A2+$F$1"))), new(new(0, 5), new(CellContent.FromLiteral(CellValue.FromNumber(10))))]);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new() { FontName = "Arial", FontSize = 11 }, new() { Bold = true, FontSize = 14 }]);
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        try
        {
            InstallDocument(new WorkbookDocument(snapshot, _services), null, "Fill smoke"); RibbonTabs.SelectedIndex = 0; await Idle();
            GridSurface.SelectRange(new(0, 0, 4, 1), new(0, 0)); await HandleGridKeyAsync(Key.D, ModifierKeys.Control, false);
            check(_document.Snapshot.Sheets[0].GetCell(4, 0)!.Value.NumberValue == 1 && _document.Snapshot.Sheets[0].GetCell(4, 1)!.Content.FormulaText == "=(A5+$F$1)",
                "Ctrl+D copies first-row formats and translates formulas across the selected rectangle");
            await UndoAsync(false);
            check(_document.Snapshot.Sheets[0].GetCell(1, 0)!.Value.NumberValue == 2 && _document.Snapshot.Sheets[0].GetCell(4, 1) is null && !_document.CanUndo,
                "Fill Down is one undoable transaction");
            GridSurface.SelectRange(new(0, 0, 1, 3), new(0, 0)); await HandleGridKeyAsync(Key.R, ModifierKeys.Control, false);
            check(_document.Snapshot.Sheets[0].GetCell(0, 3)!.Value.NumberValue == 1 && _document.Snapshot.Sheets[0].GetCell(1, 3)!.Value.NumberValue == 2,
                "Ctrl+R copies the first column without inventing a numeric series");
            await UndoAsync(false);
            GridSurface.Select(new(2, 1)); await HandleGridKeyAsync(Key.D, ModifierKeys.Control, false);
            check(_document.Snapshot.Sheets[0].GetCell(2, 1)!.Content.FormulaText == "=(A3+$F$1)" && GridRange.Rectangle(GridSurface.Selection) == new GridRange(2, 1, 2, 1),
                "Single-cell Ctrl+D copies the cell above and retains the original selection");
            await UndoAsync(false);
            GridSurface.Select(new(1, 2)); await HandleGridKeyAsync(Key.R, ModifierKeys.Control, false);
            check(_document.Snapshot.Sheets[0].GetCell(1, 2)!.Content.FormulaText == "=(B2+$F$1)", "Single-cell Ctrl+R copies the left neighbor with relative references");
            await UndoAsync(false);
            GridSurface.SelectRange(new(1, 0, 1, 1), new(1, 0)); await HandleGridKeyAsync(Key.D, ModifierKeys.Control, false);
            check(_document.Snapshot.Sheets[0].GetCell(1, 0)!.Value.NumberValue == 1 && GridRange.Rectangle(GridSurface.Selection) == new GridRange(1, 0, 1, 1),
                "A one-row Ctrl+D selection copies the complete row above");
            await UndoAsync(false);
            GridSurface.SelectRange(new(0, 1, 1, 1), new(0, 1)); await HandleGridKeyAsync(Key.R, ModifierKeys.Control, false);
            check(_document.Snapshot.Sheets[0].GetCell(0, 1)!.Value.NumberValue == 1 && _document.Snapshot.Sheets[0].GetCell(1, 1)!.Value.NumberValue == 2,
                "A one-column Ctrl+R selection copies the complete column to its left");
            await UndoAsync(false);
            GridSurface.Select(new(0, 0)); var edgeRevision = _document.Snapshot.Revision;
            await HandleGridKeyAsync(Key.D, ModifierKeys.Control, false); await HandleGridKeyAsync(Key.R, ModifierKeys.Control, false);
            check(_document.Snapshot.Revision == edgeRevision && !_document.CanUndo, "Fill at the top or left boundary adds no history");
            SetSmokeFreeze(1, 1); RefreshSheet(true); GridSurface.Select(new(1, 1)); await Idle();
            check(GridSurface.BeginEdit() && _grid.Editor.Clip.Bounds.Width > 0 && _grid.Editor.Clip.Bounds.Height > 0,
                "The editor uses the visible scrolling-pane clip when both axes are frozen");
            GridSurface.Cancel(); GridSurface.SelectRange(new(0, 0, 1, 1), new(0, 0)); await Idle();
            File.WriteAllText(Path.Combine(folder, "fill-handle-state.txt"), $"visibility={_grid.FillHandle.Visibility}; visible={_grid.FillHandle.IsVisible}; size={_grid.FillHandle.ActualWidth}x{_grid.FillHandle.ActualHeight}; " +
                $"editing={_grid.IsEditing}; composing={_grid.IsComposing}; dragging={_grid.Control.IsSelectionDragging}; count={_grid.Control.SelectedCellCount}; ranges={string.Join(";", _grid.SelectionRanges)}; " +
                $"fragments={string.Join(";", _grid.Control.GetModelRangeFragments(new(1, 1, 1, 1)))}");
            CaptureGridEvidence(Path.Combine(folder, "fill-handle.png"));
            check(_grid.FillHandle.IsVisible && _grid.FillHandle.ActualWidth >= 6, "A visible fill handle appears at the selected rectangle corner across frozen panes");
            _grid.FillHandle.RaiseEvent(new DragStartedEventArgs(0, 0) { RoutedEvent = Thumb.DragStartedEvent });
            var targetBounds = _grid.Control.GetModelCellBounds(5, 1)!.Value;
            var hit = _grid.Control.HitTestModel(new(targetBounds.Left + targetBounds.Width / 2, targetBounds.Top + targetBounds.Height / 2));
            _grid.PreviewFillAt(new(hit.Row!.Value, hit.Column!.Value)); var revision = _document.Snapshot.Revision; await Idle();
            CaptureGridEvidence(Path.Combine(folder, "fill-preview.png"));
            check(_document.Snapshot.Revision == revision && _document.Snapshot.Sheets[0].GetCell(5, 1) is null, "Fill preview is view-only and frozen viewport hit-testing maps model addresses");
            _grid.FillHandle.RaiseEvent(new DragCompletedEventArgs(0, 0, false) { RoutedEvent = Thumb.DragCompletedEvent }); await _gridInput;
            check(_document.Snapshot.Sheets[0].GetCell(5, 0)!.Value.NumberValue == 2 && _document.Snapshot.Sheets[0].GetCell(5, 1)!.Content.FormulaText == "=(A6+$F$1)",
                "The fill-handle completion handler repeats the captured pattern and copies formulas");
            await UndoAsync(false); GridSurface.SelectRange(new(0, 0, 1, 1), new(0, 0));
            check(_grid.BeginFillGesture(), "Another fill gesture can begin after undo"); _grid.PreviewFillAt(new(7, 1)); revision = _document.Snapshot.Revision;
            await HandleGridKeyAsync(Key.Escape, ModifierKeys.None, false); _grid.CompleteFillGesture(false);
            check(_document.Snapshot.Revision == revision && _document.Snapshot.Sheets[0].GetCell(7, 1) is null, "Escape discards a fill preview without adding history");
            GridSurface.SelectRange(new(0, 0, 1, 1), new(0, 0)); _grid.BeginFillGesture(); _grid.PreviewFillAt(new(7, 1));
            _editor.SetCell(0, 9, 9, "intervening"); revision = _document.Snapshot.Revision;
            _grid.CompleteFillGesture(false); try { await _gridInput; } catch (InvalidOperationException) { }
            check(_document.Snapshot.Revision == revision && _document.Snapshot.Sheets[0].GetCell(7, 1) is null, "An intervening edit expires the fill gesture before any write");
            _editor.SetCell(0, 8, 0, "31"); GridSurface.SelectRange(new(8, 0, 11, 0), new(8, 0));
            _smokeSeriesDriver = dialog =>
            {
                dialog.KindPicker.SelectedIndex = (int)CellSeriesKind.Months; dialog.StepInput.Text = "0.5";
                dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                check(dialog.IsVisible, "The actual series dialog rejects fractional date steps without closing");
                dialog.StepInput.Text = "1"; dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            Fill_Click(new MenuItem { Tag = "series" }, new()); await _gridInput; _smokeSeriesDriver = null;
            check(_document.Snapshot.Sheets[0].GetCell(9, 0)!.Value.NumberValue == 60 && _document.Snapshot.Sheets[0].GetCell(10, 0)!.Value.NumberValue == 91,
                "The series dialog fills month ends including Excel's fictitious leap day");
            GridSurface.Select(new(0, 0)); Fill_Click(new MenuItem { Tag = "clear-formats" }, new()); await _gridInput;
            check(_document.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue == 1 && _document.Snapshot.Sheets[0].GetCell(0, 0)!.Content.StyleId == 0,
                "Clear Formats uses the workbook default while retaining cell content");
            await UndoAsync(false); await HandleGridKeyAsync(Key.Delete, ModifierKeys.None, false);
            check(_document.Snapshot.Sheets[0].GetCell(0, 0)!.Value.Kind == CellValueKind.Blank && _document.Snapshot.Sheets[0].GetCell(0, 0)!.Content.StyleId == 1,
                "Clear Contents remains separate and keeps formatting");
            await UndoAsync(false);
            _editor.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 3, 1, hidden: true)); RefreshSheet(true);
            revision = _document.Snapshot.Revision;
            try { await ApplyFillAsync(new(0, 0, 0, 0), new(0, 0, 4, 0)); throw new Exception("Hidden fill accepted."); } catch (InvalidOperationException) { }
            check(_document.Snapshot.Revision == revision && _document.Snapshot.Sheets[0].GetCell(4, 0) is null, "Fill refuses hidden rows atomically instead of remapping the target");
            CurrentView.FilterRange = new(0, 0, 2, 0); CurrentView.FilterColumn = 0; CurrentView.FilterText = "1";
            try { await ApplyFillAsync(new(0, 0, 0, 0), new(0, 0, 2, 0)); throw new Exception("Filtered fill accepted."); } catch (InvalidOperationException) { }
            check(_document.Snapshot.Revision == revision, "Fill refuses a target containing filtered rows"); CurrentView.FilterRange = null;
            GridSurface.SelectRange(new(0, 0, 1, 1), new(0, 0)); _editState.StartComposition(); await FillSelectionAsync(false); _editState.EndComposition();
            check(_document.Snapshot.Revision == revision && _document.Snapshot.Sheets[0].GetCell(1, 0)!.Value.NumberValue == 2, "IME composition gates keyboard fill");
            var path = Path.Combine(folder, "filled.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, _document.Snapshot);
            check(storage.Open(path).Sheets[0].GetCell(10, 0)!.Value.NumberValue == 91, "Filled dates survive normal XLSX save and reopen");
        }
        finally { _smokeSeriesDriver = null; _grid.CancelFillGesture(); }
    }
}
