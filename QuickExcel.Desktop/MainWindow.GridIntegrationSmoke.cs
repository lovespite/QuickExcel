using System.IO;
using System.Windows.Controls;
using System.Windows.Threading;
using QuickExcel.Core;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunGridIntegrationSmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        if (ProductChildren<DataGrid>(this).Any() || !SheetGrid.Children.Contains(_grid.Element))
            throw new Exception("Main workspace must contain only the new grid surface.");
        _suppressAutoCalculation = true;
        try
        {
            await RunFilterScrollSmokeAsync(folder);
            await RunSheetSwitchSmokeAsync(folder);
            await RunSelectionOutlineSmokeAsync(folder);
            await RunMergedSelectionSmokeAsync(folder);
            InstallDocument(WorkbookDocument.CreateBlank(_services), Path.Combine(folder, "grid-editors.xlsq"), "编辑迁移.xlsq");
            await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
            GridSurface.Select(new(3, 2)); FormulaBar.Text = "11";
            if (!GridSurface.BeginEdit() || _grid.Editor.Text != "11") throw new Exception("Grid did not commit the older formula draft before loading raw input.");
            _grid.Editor.Text = "12";
            if (!await SaveAsync(false)) throw new Exception("Saving an active native editor failed.");
            using (var reopened = WorkbookDocument.Open(_filePath!, _services))
                if (reopened.GetCellInput(0, 3, 2) != "12") throw new Exception("Save omitted the active editor.");
            GridSurface.BeginEdit(); _grid.Editor.Text = "13";
            FormulaBar.Focus(); FormulaBar.Text = "=1+2";
            await ApplyFormulaBarAsync();
            if (_document.GetCellInput(0, 3, 2) != "=1+2" || _document.GetCellValue(0, 3, 2) != "3")
                throw new Exception("Older grid input overwrote the newer formula bar.");
            FormulaBar.Text = new string('x', 32768);
            var revision = _document.Snapshot.Revision;
            GridSurface.Select(new(4, 3));
            if (_grid.Current != new QuickExcel.Domain.CellPosition(3, 2) || _document.Snapshot.Revision != revision || !_formulaPending)
                throw new Exception("Rejected formula draft changed the selected address or history.");
            _editState.EndFormula(); UpdateFormulaBar();
            ApplyValues([(5, 0, "Header"), (6, 0, "keep")]); await RecalculateAsync();
            ApplyFilter(new(5, 0, 6, 0), 0, "keep");
            GridSurface.Select(new(6, 0)); GridSurface.BeginEdit(); _grid.Editor.Text = "hide";
            if (!CommitEditors()) throw new Exception("Filtered editor commit failed.");
            await RecalculateAsync();
            if (_grid.Current.Row != 5) throw new Exception("Filtering the edited row must return to the actual nonzero header.");
            _smokeDiscardDecision = System.Windows.MessageBoxResult.No;
            await RunDateDisplaySmokeAsync(folder);
            await RunCloseSmokeTestsAsync(folder);
            File.WriteAllText(Path.Combine(folder, "grid-integration-smoke-result.txt"),
                "PASS: filtered scroll ranges and bottom reachability, XLSX first-layout/scroll/sheet-switch rendering, merged selection geometry and navigation, main grid replacement, bidirectional editor ownership, rejected draft address, active-editor save, numeric/date fidelity and actual WPF close lifecycle.");
        }
        finally { _suppressAutoCalculation = false; _smokeDiscardDecision = null; }
    }
}
