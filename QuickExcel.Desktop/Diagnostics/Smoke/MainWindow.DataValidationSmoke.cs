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
    private async Task RunDataValidationSmokeAsync(string folder, string? nativePath)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>(); _suppressAutoCalculation = true;
        void Check(bool value, string message) { if (!value) throw new Exception(message); checks.Add(message); }
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        async Task Rule(WorksheetValidation rule, int row)
        {
            GridSurface.Select(new(row, 0)); _smokeValidationRule = rule; Validation_Click(this, new()); await _workspaceInput;
            _smokeValidationRule = null; await Idle();
        }
        try
        {
            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "数据验证验收");
            ApplyValues([(1, 0, "2"), (1, 1, "whole"), (2, 1, "warning"), (3, 1, "list"), (4, 1, "custom"), (5, 1, "date"),
                (0, 3, "'001"), (1, 3, "'002")]);
            var id = _document.Snapshot.Sheets[0].Id;
            await Rule(new([new(1, 0, 1, 0)], ValidationKind.Whole, "=1", "=5", inputTitle: "数量", inputMessage: "请输入 1 至 5 的整数"), 1);
            Check(_document.Snapshot.Sheets[0].Data.Validations.Length == 1 && ValidationHint.Text.Contains("1 至 5"), "Rule handler persists constraints and selection displays its input prompt");
            var revision = _document.Snapshot.Revision;
            FormulaBar.Text = "9"; Check(!CommitFormulaBar() && _formulaPending && _document.Snapshot.Revision == revision, "Invalid formula-bar entry retains draft and original document");
            Check(StatusText.Text.Contains("验证"), "Stop validation error reaches visible edit feedback");
            FormulaBar.Text = "4"; Check(CommitFormulaBar() && _document.GetCellValue(0, 1, 0) == "4", "Correcting the same draft commits a valid integer");
            var before = _document.Snapshot;
            try { _editor.SetCells(0, [(1, 0, "10"), (0, 4, "must not appear")]); throw new Exception("Invalid batch was accepted."); }
            catch (DataValidationException) { Check(ReferenceEquals(before, _document.Snapshot), "Batch input rejects all changes when one constrained cell fails"); }
            var clipboard = WorkbookClipboard.Capture(before, id, new(0, 3, 0, 3));
            try { _editor.Execute(new PasteCellsCommand(clipboard, id, new(1, 0, 1, 0), before.Revision, CellPasteMode.Values)); throw new Exception("Invalid paste was accepted."); }
            catch (DataValidationException) { Check(ReferenceEquals(before, _document.Snapshot), "Internal paste obeys the destination rule atomically"); }
            await UndoAsync(false); Check(_document.GetCellValue(0, 1, 0) == "2", "One undo restores the previous validated input");
            await Rule(new([new(2, 0, 2, 0)], ValidationKind.Whole, "=1", "=5", alert: ValidationAlert.Warning, errorMessage: "建议不超过 5"), 2);
            try { _editor.SetCell(0, 2, 0, "9"); throw new Exception("Warning lacked confirmation."); }
            catch (DataValidationException) { Check(_document.Snapshot.Sheets[0].GetCell(2, 0) is null, "Declining a warning leaves the cell unchanged"); }
            _smokeValidationContinue = true; _editor.SetCell(0, 2, 0, "9"); _smokeValidationContinue = false;
            Check(_document.GetCellValue(0, 2, 0) == "9", "Explicit warning confirmation commits the same transaction");
            await Rule(new([new(3, 0, 3, 0)], ValidationKind.List, "=$D$1:$D$2", inputTitle: "编号", inputMessage: "从列表选择，保留前导零"), 3);
            ValidationChoices_Opened(ValidationChoices, EventArgs.Empty); ValidationChoices.SelectedIndex = 0; await Idle();
            Check(_document.Snapshot.Sheets[0].GetCell(3, 0)?.Value == CellValue.FromText("001"), "Actual dropdown selection preserves a text ID and leading zeros");
            await Rule(new([new(4, 0, 4, 0)], ValidationKind.Custom, "=AND(ISNUMBER(A5),A5<10)"), 4);
            _editor.SetCell(0, 4, 0, "6"); Check(_document.GetCellValue(0, 4, 0) == "6", "Custom formula rule validates the candidate input");
            var dialog = new ValidationDialog(this, [new(5, 0, 5, 0)]);
            _ = Dispatcher.BeginInvoke(new Action(() =>
            {
                dialog.KindPicker.SelectedIndex = (int)ValidationKind.Date; dialog.First.Text = "2026-10-01"; dialog.Second.Text = "2026-10-31";
                dialog.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }), DispatcherPriority.ApplicationIdle);
            Check(dialog.ShowDialog() == true && dialog.Result?.Formula1 == "=DATE(2026,10,1)", "Real settings dialog converts ISO dates into workbook date-system formulas");
            await Rule(dialog.Result!, 5); _editor.SetCell(0, 5, 0, "2026-10-15");
            Check(_document.Snapshot.Sheets[0].GetCell(5, 0)?.Value == CellValue.FromNumber(46310), "ISO date entry under a date rule stores the workbook serial");
            Check(_document.Snapshot.Sheets[0].Data.Validations.Length == 5, "Date settings handler installs a fifth independent rule");
            GridSurface.Select(new(3, 0)); _smokeValidationContinue = true; RemoveValidation_Click(this, new()); await Idle(); _smokeValidationContinue = false;
            Check(_document.Snapshot.Sheets[0].Data.Validations.Length == 4 && _document.GetCellValue(0, 3, 0) == "001", "Clear validation handler removes the rule and retains existing content");
            await UndoAsync(false); Check(_document.Snapshot.Sheets[0].Data.Validations.Length == 5, "Undo restores the complete removed validation rule");
            await RecalculateAsync(); var path = Path.Combine(folder, "validation-ui.xlsx"); new TypedWorkbookStorage().Save(path, _document.Snapshot);
            await OpenPathAsync(path); await RecalculateAsync(); GridSurface.Select(new(3, 0)); await Idle();
            Check(_document.Snapshot.Sheets[0].Data.Validations.Length == 5 && ValidationChoices.Visibility == Visibility.Visible && ValidationHint.Text.Contains("编号"),
                "Normal desktop save/reopen restores rules, list control and input prompt");
            RibbonTabs.SelectedIndex = 2; await Idle(); CaptureGridEvidence(Path.Combine(folder, "data-validation.png"));
            if (nativePath is not null)
            {
                await OpenPathAsync(nativePath); GridSurface.Select(new(0, 0)); await Idle();
                Check(_document.Snapshot.Sheets[0].Data.Validations.Length > 100 && ValidationHint.Text.Contains("录入提示"),
                    "Native Excel validation workbook opens with modeled rules and visible prompts");
            }
            File.WriteAllLines(Path.Combine(folder, "data-validation-result.txt"), new[] { $"PASS: {checks.Count} WPF checks" }.Concat(checks));
        }
        finally { _smokeValidationRule = null; _smokeValidationContinue = false; _suppressAutoCalculation = false; }
    }
}
