using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunRecalculationCommandsSmokeAsync(Action<bool, string> check)
    {
        async Task Idle() => await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
        void Click(ButtonBase button) => button.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        var calculator = new RefreshSmokeCalculator(_services.Calculator);
        var document = WorkbookDocument.CreateBlank(new(_services.Storage, calculator));
        document.AddSheet("Other");
        document.SetCell(0, 0, 0, "2");
        document.SetCell(0, 0, 1, "=Other!A1+1");
        document.SetCell(1, 0, 0, $"='{document.Snapshot.Sheets[0].Name}'!A1*2");
        InstallDocument(document, null, "重算验收.xlsx");
        var data = RibbonTabs.Items.OfType<TabItem>().Single(t => t.Header as string == "数据");
        data.IsSelected = true; await Idle();
        check(RibbonRecalculateWorkbook.IsVisible && !_calculationPending,
            "Data ribbon exposes workbook recalculation even without pending edits");
        Click(RibbonRecalculateWorkbook); await _workspaceInput;
        check(calculator.LastForceFull && calculator.LastStatistics is { EvaluatedCount: 2, ReusedCount: 0 } &&
            _document.GetCellValue(0, 0, 1) == "5" && _document.GetCellValue(1, 0, 0) == "4" && !_document.ScriptsEnabled,
            "Workbook button fully recalculates cross-sheet formulas without enabling scripts");
        var revision = _document.Snapshot.Revision;
        var calls = calculator.Calls;
        Click(RibbonRecalculateWorkbook); await _workspaceInput;
        check(calculator.Calls == calls + 1 && calculator.LastStatistics is { FullRecalculation: true, EvaluatedCount: 2, ReusedCount: 0 } &&
            _document.Snapshot.Revision == revision,
            "Repeated manual recalculation evaluates fresh formulas without changing input revision");
        GridSurface.Select(new(0, 0)); FormulaBar.Focus(); FormulaBar.Text = "7";
        Click(RibbonRecalculateWorkbook); await _workspaceInput;
        check(_document.GetCellInput(0, 0, 0) == "7" && _document.GetCellValue(0, 0, 1) == "15",
            "Workbook button commits the active formula-bar draft before calculating");

        var reports = RibbonTabs.Items.OfType<TabItem>().Single(t => t.Header as string == "报表");
        reports.IsSelected = true; await Idle();
        var previewButton = ProductChildren<Button>((FrameworkElement)reports.Content).Single(b => b.Content as string == "打印预览 / PDF");
        var previews = 0;
        Exception? previewError = null;
        var observing = true;
        EventManager.RegisterClassHandler(typeof(Window), LoadedEvent, new RoutedEventHandler((sender, _) =>
        {
            if (!observing || sender is not Window preview || preview.Owner != this || !preview.Title.StartsWith("打印预览 ·")) return;
            previews++;
            Dispatcher.BeginInvoke(() =>
            {
                try
                {
                    var viewer = ProductChildren<DocumentViewer>(preview).Single();
                    check(viewer.Document.DocumentPaginator.PageCount > 0 && calculator.LastForceFull &&
                        _document.Snapshot.Sheets.SelectMany(s => s.Cells.Values).Where(c => c.Content.Kind == CellContentKind.Formula)
                            .All(c => c.Calculation?.State == CalculationState.Fresh),
                        "Print preview opens only after full-workbook calculation has completed");
                }
                catch (Exception ex) { previewError = ex; }
                finally { preview.Close(); }
            });
        }));
        try
        {
            // A newly installed document, like a newly opened file, has no pending UI calculation flag.
            InstallDocument(new WorkbookDocument(_document.Snapshot, new(_services.Storage, calculator)), null, "打印验收.xlsx");
            check(!_calculationPending, "Print fixture starts without pending UI edits");
            calls = calculator.Calls;
            Click(previewButton); await _workspaceInput;
            if (previewError is not null) throw previewError;
            check(previews == 1 && calculator.Calls == calls + 1 && calculator.LastStatistics is { FullRecalculation: true, EvaluatedCount: 2, ReusedCount: 0 },
                $"Preview forces one complete round even without pending edits (previews={previews}, calls={calculator.Calls - calls}, statistics={calculator.LastStatistics}, status={StatusText.Text})");
            GridSurface.Select(new(0, 0)); GridSurface.BeginEdit(); _grid.Editor.Text = "9";
            calls = calculator.Calls;
            Click(previewButton); await _workspaceInput;
            if (previewError is not null) throw previewError;
            check(previews == 2 && calculator.Calls == calls + 1 && _document.GetCellValue(0, 0, 1) == "19",
                "Preview commits the grid draft and recalculates exactly once before pagination");

            var gate = calculator.HoldNext();
            Click(previewButton); var pending = _workspaceInput;
            try
            {
                await gate.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
                Click(previewButton);
                StopCalculation_Click(this, new());
            }
            finally { gate.Release.TrySetResult(); }
            await pending;
            check(previews == 2 && !_busy && _calculationPending && _document.GetCellValue(0, 0, 1) == "19",
                "Cancelled calculation and a duplicate preview click never open a preview with old results");

            var failing = new SmokeFailingCalculator();
            InstallDocument(new WorkbookDocument(_document.Snapshot, new(_services.Storage, failing)), null, "失败验收.xlsx");
            Click(previewButton); await _workspaceInput;
            check(previews == 2 && failing.Attempts == 1 && _calculationPending,
                "Failed forced calculation cancels preview without retrying or printing cached values");

            InstallDocument(WorkbookDocument.CreateBlank(_services), null, "脚本未授权.xlsx");
            _editor.SetCell(0, 0, 0, "=JS.MISSING()");
            Click(previewButton); await _workspaceInput;
            check(previews == 2 && !_document.ScriptsEnabled && StatusText.Text.Contains("打印已取消"),
                "Blocked script formulas still prevent preview and never gain execution permission");
        }
        finally { observing = false; }
    }
}
