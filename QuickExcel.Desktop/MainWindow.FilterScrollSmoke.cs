using System.IO;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunFilterScrollSmokeAsync(string folder)
    {
        var cases = new List<string>();
        void Check(bool condition, string message)
        {
            if (!condition) throw new Exception("Filter scrolling: " + message);
            cases.Add(message);
            File.WriteAllLines(Path.Combine(folder, "filter-scroll-cases.txt"), cases);
        }
        var cells = Enumerable.Range(0, 241).ToDictionary(row => new CellPosition(row, 0),
            row => new CellSnapshot(CellContent.FromLiteral(CellValue.FromText(row == 0 ? "Header" : row % 20 == 0 ? "keep" : "hide"))));
        var sheet = new WorksheetSnapshot(SheetId.New(), "Filter scrolling", cells,
            rows: Enumerable.Range(1, 240).Select(row => new KeyValuePair<int, RowMetadata>(row, new(Height: 120))),
            defaultRowHeight: 12);
        InstallDocument(new WorkbookDocument(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]), _services), null, "筛选滚动回归.xlsq");
        var grid = _grid.Control;
        var bar = (ScrollBar)grid.FindName("vscroll");
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(grid);
        }
        void Scroll(double value)
        {
            bar.Value = value;
            bar.RaiseEvent(new ScrollEventArgs(ScrollEventType.ThumbPosition, bar.Value) { RoutedEvent = ScrollBar.ScrollEvent });
        }
        async Task Filter(string text)
        {
            await ApplyFilterPopupResultAsync(0, new(FilterPopupAction.Apply, new(First: new(FilterOperator.TextEquals, text))));
            await Idle();
        }
        EstablishDataRegion(new(0, 0, 240, 0));
        CurrentView.FrozenRows = 1; RefreshSheet(); await Idle();
        var revision = _document.Snapshot.Revision;
        var originalMaximum = bar.Maximum;
        Scroll(180); await Idle();
        await Filter("keep");
        Check(bar.Maximum > 0 && bar.Maximum < originalMaximum,
            "Filtering tall rows recomputes a positive, smaller scroll range");
        Check(bar.Value == grid.FirstVisibleRowScrollIndex && bar.ViewportSize > 0,
            "Filtering from a scrolled viewport synchronizes the thumb and visible row count");
        Scroll(bar.Maximum); await Idle();
        Check(grid.GetModelCellBounds(240, 0) is { } last && last.Bottom <= ((Image)grid.FindName("image")).ActualHeight + 1,
            "Dragging the filtered scrollbar to the bottom reveals the complete final data row");
        Check(grid.GetVisibleModelRows().Contains(_grid.RowCount - 1),
            "The filtered scrollbar also reaches rows outside the filter region");
        CaptureGridEvidence(Path.Combine(folder, "filter-scroll-bottom.png"));
        Scroll(0); await Idle();
        var content = (Image)grid.FindName("image");
        for (var i = 0; i < 20; i++)
            content.RaiseEvent(new MouseWheelEventArgs(Mouse.PrimaryDevice, Environment.TickCount, -120) { RoutedEvent = Mouse.MouseWheelEvent });
        await Idle();
        Check(grid.GetVisibleModelRows().Contains(240), "Mouse wheel scrolling also reaches the final filtered data row");
        await Filter("absent");
        Check(bar.Maximum == 0 && bar.Value == 0 && grid.FirstVisibleRowScrollIndex == 0 && bar.ViewportSize > 0,
            "An empty result resets the scroll position and keeps the remaining blank rows usable");
        Check(grid.GetModelCellBounds(0, 0) is not null, "The frozen filter header remains visible for an empty result");
        await Filter("keep");
        Check(bar.Maximum > 0, "Reapplying a nonempty filter restores scrolling without resizing the window");
        Scroll(bar.Maximum); await Idle();
        await ApplyFilterPopupResultAsync(0, new(FilterPopupAction.ClearAll, null)); await Idle();
        Check(bar.Maximum == originalMaximum && bar.Value == grid.FirstVisibleRowScrollIndex && bar.ViewportSize > 0,
            "Clearing the filter restores the original scroll range and thumb state");
        Scroll(bar.Maximum); await Idle();
        Check(grid.GetVisibleModelRows().Contains(_grid.RowCount - 1), "The restored scrollbar reaches the unfiltered bottom");
        Check(_document.Snapshot.Revision == revision, "Filtering and scrolling do not write document history");
    }
}
