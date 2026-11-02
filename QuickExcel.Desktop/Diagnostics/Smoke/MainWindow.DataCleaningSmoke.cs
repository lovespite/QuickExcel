using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunDataCleaningSmokeAsync(string folder, string? nativePath)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>(); _suppressAutoCalculation = true;
        void Check(bool value, string message) { if (!value) throw new Exception(message); checks.Add(message); }
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        async Task Commit(CleaningPreview preview, CleaningOperation operation, bool overwrite = false)
        {
            _smokeCleaningPreview = preview; _smokeCleaningOverwrite = overwrite;
            if (operation == CleaningOperation.Split) SplitText_Click(this, new());
            else if (operation == CleaningOperation.Duplicates) RemoveDuplicates_Click(this, new()); else ConvertNumbers_Click(this, new());
            await _workspaceInput; _smokeCleaningPreview = null; _smokeCleaningOverwrite = false;
            Check(StatusText.Text.StartsWith("清洗已应用"), "Handler commits the reviewed candidate: " + StatusText.Text); await RecalculateAsync(); await Idle();
        }
        try
        {
            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "数据清洗验收");
            ApplyValues([(0, 0, "Source"), (1, 0, "'001,2.50,2026-10-31"), (2, 0, "'002,3,2026-11-01"), (1, 3, "existing"),
                (5, 0, "Key"), (5, 1, "Record"), (6, 0, "red"), (6, 1, "1"), (7, 0, "RED"), (7, 1, "2"), (8, 0, "blue"), (8, 1, "3"),
                (6, 2, "c1"), (7, 2, "c2"), (8, 2, "c3"),
                (0, 6, "Number"), (1, 6, "'001"), (2, 6, "'12-"), (3, 6, "=1+1"), (0, 7, "'1234567890123456")]);
            await RecalculateAsync(); await Idle();
            var id = _document.Snapshot.Sheets[0].Id; var before = _document.Snapshot;
            var dialog = new CleaningDialog(this, before, id, new(1, 0, 2, 0), CleaningOperation.Split);
            _ = Dispatcher.BeginInvoke(new Action(async () =>
            {
                dialog.Destination.Text = "C2"; dialog.Types.Text = "文本,数字,YMD"; dialog.PreviewButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                Check(ReferenceEquals(before, _document.Snapshot) && dialog.Preview?.Changes.Length == 6, "Real preview computes typed changes without modifying the workbook");
                Check(dialog.Preview?.Overwrites.Length == 1 && !dialog.ApplyButton.IsEnabled, "Existing destination content requires explicit overwrite selection");
                dialog.Types.Text = "TEXT,NUMBER,YMD";
                Check(dialog.Preview is null && !dialog.ApplyButton.IsEnabled, "Changing a setting invalidates the previous preview");
                dialog.PreviewButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); dialog.Overwrite.IsChecked = true;
                await Dispatcher.InvokeAsync(dialog.UpdateLayout, DispatcherPriority.ApplicationIdle);
                var root = dialog; var dpi = VisualTreeHelper.GetDpi(root);
                var bitmap = new RenderTargetBitmap((int)(root.ActualWidth * dpi.DpiScaleX), (int)(root.ActualHeight * dpi.DpiScaleY), dpi.PixelsPerInchX, dpi.PixelsPerInchY, PixelFormats.Pbgra32);
                bitmap.Render(root); var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap));
                using (var file = File.Create(Path.Combine(folder, "cleaning-preview.png"))) encoder.Save(file);
                dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }), DispatcherPriority.ApplicationIdle);
            Check(dialog.ShowDialog() == true && dialog.Preview is not null, "Applying the dialog returns the reviewed candidate");
            await Commit(dialog.Preview!, CleaningOperation.Split, true);
            var sheet = _document.Snapshot.Sheets[0];
            Check(sheet.GetCell(1, 2)!.Value == CellValue.FromText("001") && sheet.GetCell(1, 3)!.Value == CellValue.FromNumber(2.5), "Split handler preserves the text ID and explicitly converts the numeric field");
            Check(sheet.GetCell(1, 4)!.Value == CellValue.FromNumber(46326) && _grid.DisplayText(1, 4) == "2026/10/31", "Explicit YMD conversion uses the workbook date system and visible date format");
            await UndoAsync(false); Check(_document.GetCellValue(0, 1, 3) == "existing" && _document.Snapshot.Sheets[0].GetCell(1, 2) is null, "One undo restores all overwritten and newly written cells");
            await UndoAsync(true); await Idle();
            var numeric = DataCleaning.ConvertTextNumbers(_document.Snapshot, id, new(1, 6, 3, 6));
            Check(numeric.Notices.Any(s => s.Contains("前导零")), "Numeric preview explains explicit leading-zero removal");
            await Commit(numeric, CleaningOperation.Numbers);
            Check(_document.Snapshot.Sheets[0].GetCell(1, 6)!.Value == CellValue.FromNumber(1) && _document.Snapshot.Sheets[0].GetCell(2, 6)!.Value == CellValue.FromNumber(-12),
                "Numeric handler converts the selected text values");
            Check(_document.Snapshot.Sheets[0].GetCell(3, 6)!.Content.FormulaText == "=1+1", "Numeric conversion retains existing formulas");
            var duplicates = DataCleaning.RemoveDuplicates(_document.Snapshot, id, new(5, 0, 8, 1), [0]);
            Check(duplicates.RemovedRows.SequenceEqual([7]), "Duplicate preview identifies the original removed row");
            var expanded = DataCleaning.RemoveDuplicates(_document.Snapshot, id, new(5, 0, 8, 1), [0], expandToWholeRows: true);
            Check(expanded.Result.Sheets[0].GetCell(7, 2)!.Value == CellValue.FromText("c3") && expanded.Result.Sheets[0].GetCell(8, 2) is null,
                "Expanded duplicate preview removes the whole row including adjacent columns outside the selection");
            await Commit(duplicates, CleaningOperation.Duplicates);
            Check(_document.GetCellValue(0, 7, 0) == "blue" && _document.GetCellValue(0, 7, 1) == "3" && _document.Snapshot.Sheets[0].GetCell(8, 0) is null,
                "Duplicate handler compacts whole selected records and clears the tail");
            var unchanged = _document.Snapshot;
            try { _ = DataCleaning.ConvertTextNumbers(unchanged, id, new(0, 7, 0, 7)); throw new Exception("Long ID conversion should fail."); }
            catch (ArgumentException) { Check(ReferenceEquals(unchanged, _document.Snapshot), "A precision-risk input cannot partially apply"); }
            var stale = DataCleaning.ConvertTextNumbers(_document.Snapshot, id, new(1, 7, 1, 7));
            _editor.SetCell(0, 4, 6, "changed");
            try { _editor.Execute(new ApplyCleaningCommand(stale)); throw new Exception("Stale preview was accepted."); }
            catch (InvalidOperationException) { Check(_document.GetCellValue(0, 4, 6) == "changed", "Stale preview is rejected without undoing newer edits"); }
            _editor.Execute(new SetAxisMetadataCommand(id, SheetAxis.Column, 2, 3, 24));
            await RecalculateAsync(); var path = Path.Combine(folder, "cleaning-ui.xlsx"); new TypedWorkbookStorage().Save(path, _document.Snapshot);
            await OpenPathAsync(path); await RecalculateAsync(); await Idle();
            Check(_document.Snapshot.Sheets[0].GetCell(1, 2)!.Value == CellValue.FromText("001") && _document.Snapshot.Sheets[0].GetCell(1, 4)!.Value == CellValue.FromNumber(46326),
                "Ordinary save and reopen retain explicit output types");
            RibbonTabs.SelectedIndex = 2; await Idle(); CaptureGridEvidence(Path.Combine(folder, "data-cleaning.png"));
            if (nativePath is not null) { await OpenPathAsync(nativePath); await RecalculateAsync(); Check(_document.Snapshot.Sheets.Length >= 9, "Native cleaning reference workbook opens and recalculates"); }
            File.WriteAllLines(Path.Combine(folder, "data-cleaning-result.txt"), new[] { $"PASS: {checks.Count} WPF checks" }.Concat(checks));
        }
        finally { _smokeCleaningPreview = null; _smokeCleaningOverwrite = false; _suppressAutoCalculation = false; }
    }
}
