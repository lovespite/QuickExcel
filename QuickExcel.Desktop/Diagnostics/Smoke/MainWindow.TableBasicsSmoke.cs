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
    private async Task RunTableBasicsSmokeAsync(string folder, string? nativePath)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>(); _suppressAutoCalculation = true;
        void Check(bool ok, string text) { if (!ok) throw new Exception(text); checks.Add(text); }
        async Task Idle() => await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
        try
        {
            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "表格验收");
            ApplyValues([(0, 0, "Group"), (0, 1, "Amount"), (1, 0, "B"), (1, 1, "20"), (2, 0, "A"), (2, 1, "10"), (3, 0, "A"), (3, 1, "30")]);
            RefreshSheet(); GridSurface.SelectRange(new(0, 0, 3, 1), new(0, 0));
            _smokeTextDecision = "Orders"; CreateTable_Click(this, new()); await _workspaceInput; _smokeTextDecision = null; await Idle();
            var sheet = _document.Snapshot.Sheets[0]; var table = sheet.Data.Tables.Single();
            Check(table.Name == "Orders" && _document.IsDirty, "Create Table handler stores the name and creates a document revision");
            Check(_grid.HeaderButtons.Count == 2 && ViewStatus.Text.Contains("Orders"), "Table header buttons and name are visible");
            Check(TablePresentation.GetStyle(_document.Snapshot, sheet, 0, 0).FillColor == "FF156082", "Named table header style reaches presentation");
            await ApplyFilterPopupResultAsync(0, new(FilterPopupAction.Apply, new(Values: System.Collections.Immutable.ImmutableHashSet.Create(CellValue.FromText("A"))))); await Idle();
            Check(CurrentView.TableId == table.Id && CurrentView.IsFilteredOut(_document.Snapshot, 0, 1) && _grid.Control.GetModelCellBounds(1, 0) is null,
                "Table filter panel updates its own conditions and real grid visibility");
            ClearFilter_Click(this, new()); Check(_document.Snapshot.Sheets[0].Data.Tables[0].Filter?.Columns.Count == 0 && _document.Snapshot.Sheets[0].Data.Filter is null,
                "Clear table filtering retains table and buttons without creating a worksheet AutoFilter");
            GridSurface.Select(new(1, 1)); Sort_Click(new MenuItem { Tag = "asc" }, new()); await _workspaceInput;
            Check(_document.GetCellInput(0, 1, 1) == "10" && _document.Snapshot.Sheets[0].Data.Tables[0].Sort is not null, "Actual sort handler sorts complete table records and persists table sort state");
            await UndoAsync(false); Check(_document.GetCellInput(0, 1, 1) == "20", "One undo restores table record order and sort metadata");
            GridSurface.Select(new(1, 0)); _smokeTextDecision = "6"; ResizeTable_Click(this, new()); await _workspaceInput; _smokeTextDecision = null;
            Check(_document.Snapshot.Sheets[0].Data.Tables[0].Range.Bottom == 5 && _document.GetCellInput(0, 3, 1) == "30", "Resize handler extends the table without moving cell content");
            await UndoAsync(false); Check(_document.Snapshot.Sheets[0].Data.Tables[0].Range.Bottom == 3, "Resize is one undoable transaction");
            var revision = _document.Snapshot.Revision;
            try { _document.SetCell(0, 0, 0, "renamed header"); throw new Exception("Header edit was accepted"); } catch (InvalidOperationException) { }
            Check(_document.Snapshot.Revision == revision && _document.GetCellInput(0, 0, 0) == "Group", "Protected header edits retain the original revision and content");
            _document.SetCell(0, 1, 1, "25"); RefreshSheet(); Check(_document.Snapshot.Sheets[0].Data.Tables[0].Name == "Orders", "Ordinary table body edit retains table identity");
            var path = Path.Combine(folder, "table-ui.xlsx"); new TypedWorkbookStorage().Save(path, _document.Snapshot);
            await OpenPathAsync(path); GridSurface.Select(new(0, 0)); await Idle();
            Check(CurrentView.TableId.HasValue && _grid.HeaderButtons.Count == 2 && _document.GetCellInput(0, 1, 1) == "25", "Normal desktop reopen restores table buttons and edited cell");
            RibbonTabs.SelectedIndex = 2; await Idle(); await _selectionStatisticsTask; await WaitForFastGridPaintAsync(_grid.Control);
            CaptureGridEvidence(Path.Combine(folder, "table-ui.png"));
            if (nativePath is not null)
            {
                await OpenPathAsync(nativePath); SheetPicker.SelectedIndex = 3; GridSurface.Select(new(2, 2)); await Idle();
                Check(_document.Snapshot.Sheets[3].Data.Tables[0].Name == "QE_bands" && _grid.HeaderButtons.Count == 4, $"Native Excel table displays all four table filter buttons (sheet={_sheetIndex}, headers={_grid.HeaderButtons.Count}, region={CurrentView.Region}, bounds={_grid.Control.GetModelCellBounds(1, 1)}, status={StatusText.Text})");
                await WaitForFastGridPaintAsync(_grid.Control); CaptureGridEvidence(Path.Combine(folder, "native-table-bands.png"));
                SheetPicker.SelectedIndex = 6; GridSurface.Select(new(2, 1)); await Idle();
                Check(_grid.Control.GetModelCellBounds(3, 1) is null && _grid.Control.GetModelCellBounds(4, 1) is not null, "Native Excel table filter controls actual grid row visibility");
                SheetPicker.SelectedIndex = 7; GridSurface.Select(new(7, 1)); await Idle();
                Check(ViewStatus.Text.Contains("汇总行") && _document.GetCellValue(7, 7, 2) == "150" && _document.Snapshot.Sheets[7].GetCell(7, 2)!.Value.Kind == CellValueKind.Text,
                    "Native totals row retains Excel's literal label text and its metadata in the workspace");
            }
            File.WriteAllLines(Path.Combine(folder, "table-basics-result.txt"), new[] { $"PASS: {checks.Count} WPF checks" }.Concat(checks));
        }
        finally { _suppressAutoCalculation = false; _smokeTextDecision = null; }
    }
}
