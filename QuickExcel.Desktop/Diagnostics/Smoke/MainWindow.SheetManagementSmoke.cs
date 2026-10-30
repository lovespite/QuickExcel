using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunSheetManagementSmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>(); _suppressAutoCalculation = true;
        void Check(bool ok, string message) { if (!ok) throw new Exception(message); checks.Add(message); }
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        var source = new WorksheetSnapshot(SheetId.New(), "Source", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(10)))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromNumber(20)))), new(new(0, 1), new(CellContent.FromFormula("=Rate")))],
            images: [new(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aA1sAAAAASUVORK5CYII="), "png", 3, 3, 2, 2)]);
        var hidden = new WorksheetSnapshot(SheetId.New(), "Hidden", visibility: SheetVisibility.Hidden);
        var other = new WorksheetSnapshot(SheetId.New(), "Other", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(100))))]);
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source, hidden, other], definedNames:
            [new("Rate", "=Source!$A$1"), new("Rate", "=Source!$A$2", source.Id)]);
        try
        {
            InstallDocument(new WorkbookDocument(snapshot, _services), null, "Sheet management smoke"); await Idle();
            _smokeTextDecision = "副本"; SheetManage_Click(new MenuItem { Tag = "copy" }, new()); await RecalculateAsync(); await Idle();
            var copy = _document.Snapshot.Sheets.Single(s => s.Name == "副本");
            Check(_document.Snapshot.Sheets.Length == 4 && _document.Snapshot.ActiveSheetId == copy.Id && _document.Snapshot.Sheets[_sheetIndex].Id == copy.Id,
                "Copy menu creates and activates the new stable sheet identity in one transaction");
            Check(copy.Images.Length == 1 && copy.GetCell(0, 1)!.Value.NumberValue == 20 && _document.Snapshot.DefinedNames.Any(n => n.Scope == copy.Id && n.Name == "Rate"),
                "The copied sheet keeps its picture and local-name formula meaning");
            await UndoAsync(false);
            Check(_document.Snapshot.Sheets.Length == 3 && _document.Snapshot.Sheets[_sheetIndex].Id == source.Id && _document.Snapshot.ActiveSheetId == source.Id && !_document.CanUndo,
                "One undo removes the copy and restores the original active sheet");
            await UndoAsync(true);
            Check(_document.Snapshot.Sheets[_sheetIndex].Id == copy.Id && _document.Snapshot.ActiveSheetId == copy.Id, "Redo restores the copy and its active tab consistently");
            SheetManage_Click(new MenuItem { Tag = "right" }, new()); await RecalculateAsync();
            Check(_document.Snapshot.Sheets.Select(s => s.Name).SequenceEqual(new[] { "Source", "Hidden", "Other", "副本" }) && _document.Snapshot.Sheets[_sheetIndex].Id == copy.Id,
                "Move Right crosses hidden tabs while retaining the active sheet identity");
            await UndoAsync(false);
            Check(_document.Snapshot.Sheets[1].Id == copy.Id && _document.Snapshot.Sheets[_sheetIndex].Id == copy.Id, "Undo reorder restores tab order without switching sheets");
            _smokeNameManagerDriver = dialog =>
            {
                dialog.NameInput.Text = "A1"; dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                Check(dialog.IsVisible, "Name manager rejects a cell address as a name without closing");
                dialog.NameInput.Text = "金额"; dialog.FormulaInput.Text = "='副本'!$A$1";
                dialog.UpdateLayout();
                var dpi = System.Windows.Media.VisualTreeHelper.GetDpi(dialog);
                var image = new System.Windows.Media.Imaging.RenderTargetBitmap((int)Math.Ceiling(dialog.ActualWidth * dpi.DpiScaleX),
                    (int)Math.Ceiling(dialog.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, System.Windows.Media.PixelFormats.Pbgra32);
                image.Render(dialog);
                var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder(); encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(image));
                using (var output = File.Create(Path.Combine(folder, "name-manager.png"))) encoder.Save(output);
                dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            NameManager_Click(this, new()); await _gridInput;
            Check(_document.Snapshot.DefinedNames.Any(n => n.Name == "金额" && n.Scope is null), "Actual name manager creates a Unicode workbook name");
            _editor.SetCell(_sheetIndex, 2, 1, "=金额"); _calculationPending = true; await RecalculateAsync();
            var revision = _document.Snapshot.Revision;
            _smokeNameManagerDriver = dialog =>
            {
                dialog.Names.SelectedItem = dialog.Names.Items.Cast<NameManagerDialog.NameChoice>().Single(n => n.Definition.Name == "金额");
                Check(!dialog.ScopePicker.IsEnabled, "An existing name keeps its original scope in the editor");
                dialog.NameInput.Text = "总额"; dialog.FormulaInput.Text = "=Other!$A$1"; dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            NameManager_Click(this, new()); await _gridInput;
            Check(_document.Snapshot.Revision == revision + 1 && _document.Snapshot.GetSheet(copy.Id).GetCell(2, 1)!.Value.NumberValue == 100 &&
                _document.Snapshot.GetSheet(copy.Id).GetCell(2, 1)!.Content.FormulaText == "=总额", "Rename and definition replacement update dependent cells in one transaction");
            await UndoAsync(false);
            Check(_document.Snapshot.GetSheet(copy.Id).GetCell(2, 1)!.Value.NumberValue == 10 && _document.Snapshot.DefinedNames.Any(n => n.Name == "金额"),
                "One undo restores both the old name and its original definition");
            await UndoAsync(true);
            _smokeNameManagerDriver = dialog =>
            {
                dialog.Names.SelectedItem = dialog.Names.Items.Cast<NameManagerDialog.NameChoice>().Single(n => n.Definition.Name == "总额");
                dialog.DeleteButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };
            NameManager_Click(this, new()); await _gridInput;
            Check(_document.Snapshot.GetSheet(copy.Id).GetCell(2, 1)!.Value.ErrorCode == CellErrorCode.Name, "Deleting a workbook name keeps the formula and exposes NAME error");
            await UndoAsync(false); Check(_document.Snapshot.GetSheet(copy.Id).GetCell(2, 1)!.Value.NumberValue == 100, "Undo name deletion restores calculation");
            revision = _document.Snapshot.Revision; _editState.StartComposition(); SheetManage_Click(new MenuItem { Tag = "copy" }, new()); _editState.EndComposition();
            Check(_document.Snapshot.Revision == revision, "IME composition gates sheet management");
            var path = Path.Combine(folder, "sheet-management.xlsx"); new TypedWorkbookStorage().Save(path, _document.Snapshot);
            var restored = new TypedWorkbookStorage().Open(path);
            Check(restored.Sheets.Length == 4 && restored.Sheets[1].Name == "副本" && restored.Sheets[1].Images.Length == 1 && restored.DefinedNames.Any(n => n.Name == "总额") &&
                restored.GetSheet(restored.ActiveSheetId).Name == "副本", "Normal XLSX save/reopen keeps copied images, scoped names, order and active sheet");
            await Idle(); CaptureGridEvidence(Path.Combine(folder, "sheet-management.png"));
            File.WriteAllLines(Path.Combine(folder, "sheet-management-result.txt"), new[] { $"PASS: {checks.Count} checks" }.Concat(checks));
        }
        finally { _smokeTextDecision = null; _smokeNameManagerDriver = null; _suppressAutoCalculation = false; }
    }
}
