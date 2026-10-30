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
    private async Task RunStructuredTablesSmokeAsync(string folder, string? nativePath)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>(); _suppressAutoCalculation = true;
        void Check(bool value, string message) { if (!value) throw new Exception(message); checks.Add(message); }
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        try
        {
            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "结构化引用验收");
            ApplyValues([(0, 0, "Qty"), (0, 1, "Price"), (0, 2, "Amount"), (1, 0, "2"), (1, 1, "10"), (2, 0, "3"), (2, 1, "20")]);
            RefreshSheet(); GridSurface.SelectRange(new(0, 0, 2, 2), new(0, 0));
            _smokeTextDecision = "Orders"; CreateTable_Click(this, new()); await _workspaceInput;
            GridSurface.Select(new(1, 2)); _smokeTextDecision = "=[@Qty]*[@Price]"; CalculatedTableColumn_Click(this, new()); await _workspaceInput; _smokeTextDecision = null;
            _document.SetCell(0, 0, 4, "=SUM(Orders[Amount])"); await RecalculateAsync(); await Idle();
            Check(_document.GetCellValue(0, 1, 2) == "20" && _document.GetCellValue(0, 2, 2) == "60", "Calculated column handler fills all rows and computes current-row references");
            Check(_document.GetCellValue(0, 0, 4) == "80", "Table aggregate outside the table recalculates from its calculated column");
            _document.SetCell(0, 1, 0, "4"); await RecalculateAsync();
            Check(_document.GetCellValue(0, 0, 4) == "100", "Ordinary body edit updates the calculated column and aggregate");
            GridSurface.Select(new(1, 1)); _smokeTextDecision = "Unit Price"; RenameTableColumn_Click(this, new()); await _workspaceInput; _smokeTextDecision = null; await RecalculateAsync();
            Check(_document.GetCellValue(0, 0, 1) == "Unit Price" && _document.GetCellValue(0, 0, 4) == "100", "Column rename handler updates header and all bound formulas");
            _smokeTextDecision = "Revenue"; RenameTable_Click(this, new()); await _workspaceInput; _smokeTextDecision = null; await RecalculateAsync();
            Check(_document.GetCellInput(0, 0, 4).Contains("Revenue"), "Table rename handler updates the external aggregate");
            await UndoAsync(false);
            Check(_document.Snapshot.Sheets[0].Data.Tables[0].Name == "Orders" && _document.GetCellInput(0, 0, 4).Contains("Orders"), "One undo restores table name and external references together");
            await UndoAsync(true);
            GridSurface.Select(new(1, 1)); _smokeTextDecision = "6"; ResizeTable_Click(this, new()); await _workspaceInput; _smokeTextDecision = null; await RecalculateAsync();
            Check(_document.Snapshot.Sheets[0].GetCell(5, 2)?.Content.Kind == CellContentKind.Formula && _document.GetCellValue(0, 0, 4) == "100", "Expanded rows inherit the calculated column and enter the aggregate range");
            await UndoAsync(false); Check(_document.Snapshot.Sheets[0].Data.Tables[0].Range.Bottom == 2 && _document.Snapshot.Sheets[0].GetCell(5, 2) is null, "Expansion undo removes generated formulas as one transaction");
            GridSurface.Select(new(2, 1)); Structure_Click(new MenuItem { Tag = "insertRow" }, new()); await RecalculateAsync();
            Check(_document.Snapshot.Sheets[0].Data.Tables[0].Range.Bottom == 3 && _document.Snapshot.Sheets[0].GetCell(2, 2)?.Content.Kind == CellContentKind.Formula,
                "Actual insert-row handler accepts a persisted table and fills its inserted calculated row");
            await UndoAsync(false);
            GridSurface.Select(new(1, 1)); _smokeTextDecision = "Qty"; var revision = _document.Snapshot.Revision;
            RenameTableColumn_Click(this, new()); await _workspaceInput; _smokeTextDecision = null;
            Check(_document.Snapshot.Revision == revision && _document.GetCellValue(0, 0, 1) == "Unit Price", "Duplicate column names are rejected atomically by the real handler");
            _smokeTextDecision = "Copy"; SheetManage_Click(new MenuItem { Tag = "copy" }, new()); _smokeTextDecision = null; await RecalculateAsync(); await Idle();
            Check(_document.Snapshot.Sheets.Length == 2 && _document.Snapshot.Sheets[1].Data.Tables[0].Name != "Revenue" && _document.GetCellValue(1, 0, 4) == "100",
                "Sheet copy handler remaps the table, calculated column and external aggregate on the copy");
            await UndoAsync(false); await RecalculateAsync();
            var path = Path.Combine(folder, "structured-ui.xlsx"); new TypedWorkbookStorage().Save(path, _document.Snapshot);
            await OpenPathAsync(path); await RecalculateAsync(); await Idle();
            Check(!_document.Snapshot.FileState!.FormulaReadOnly && _document.GetCellValue(0, 0, 4) == "100", "Normal save/reopen keeps supported structured formulas editable and recalculable");
            RibbonTabs.SelectedIndex = 2; GridSurface.Select(new(1, 2)); await Idle(); CaptureGridEvidence(Path.Combine(folder, "structured-table.png"));
            if (nativePath is not null)
            {
                await OpenPathAsync(nativePath); await RecalculateAsync(); await Idle();
                Check(!_document.Snapshot.FileState!.FormulaReadOnly && _document.GetCellValue(1, 0, 0) == "200", "Native Excel workbook imports and recalculates structured formulas without read-only fallback");
                Check(_document.GetCellValue(3, 2, 5) == "10", "A literal at-sign in an OOXML table column is distinguished from the current-row operator");
            }
            File.WriteAllLines(Path.Combine(folder, "structured-tables-result.txt"), new[] { $"PASS: {checks.Count} WPF checks" }.Concat(checks));
        }
        finally { _smokeTextDecision = null; _suppressAutoCalculation = false; }
    }
}
