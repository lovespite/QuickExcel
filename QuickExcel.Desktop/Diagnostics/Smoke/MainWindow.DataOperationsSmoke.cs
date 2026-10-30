using System.Collections.Immutable;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunDataOperationsSmokeAsync(string folder)
    {
        Directory.CreateDirectory(folder); var checks = new List<string>();
        void Check(bool condition, string name) { if (!condition) throw new Exception(name); checks.Add(name); File.AppendAllText(Path.Combine(folder, "progress.txt"), name + Environment.NewLine); }
        async Task Idle() => await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
        await VerifyFastGridRegressionAsync(Check);
        _suppressAutoCalculation = true;
        InstallDocument(WorkbookDocument.CreateBlank(_services), null, "数据操作验收.xlsq");
        ApplyValues([(2, 1, "部门"), (2, 2, "金额"), (2, 3, "日期"), (2, 4, "标签"), (2, 5, "公式"),
            (3, 1, "乙"), (3, 2, "30"), (3, 3, "46000"), (3, 4, "first"), (3, 5, "=C4*2"),
            (4, 1, "甲"), (4, 2, "10"), (4, 3, "46001"), (4, 4, "second"), (4, 5, "=C5*2"),
            (5, 1, "甲"), (5, 2, "20"), (5, 3, "46002"), (5, 4, "third"), (5, 5, "=C6*2"),
            (6, 1, "甲"), (6, 2, "20"), (6, 4, "fourth"), (8, 0, "outside")]);
        _editor.Execute(new SetAxisMetadataCommand(_document.Snapshot.Sheets[0].Id, SheetAxis.Row, 5, 1, hidden: true));
        RefreshSheet(true); await RecalculateAsync(); await Idle();
        GridSurface.SelectRange(new(2, 1, 6, 5), new(2, 1));
        Filter_Click(SheetGrid, new()); await Idle();
        Check(CurrentView.Region?.Range == new GridRange(2, 1, 6, 5), "Filter command explicitly establishes B3:F7 region");
        Check(_grid.HeaderButtons.Count == 5, "Five real header-cell filter buttons are realized");
        var headerButton = _grid.HeaderButtons.FirstOrDefault() ?? throw new Exception("Header filter button missing from visual tree.");
        var headerPoint = headerButton.TranslatePoint(new Point(headerButton.ActualWidth / 2, headerButton.ActualHeight / 2), this);
        var headerHit = InputHitTest(headerPoint) as DependencyObject;
        Check(headerButton.IsVisible && headerButton.ActualWidth > 0 && headerButton.ActualHeight > 0 && headerHit is not null &&
            (ReferenceEquals(headerHit, headerButton) || headerButton.IsAncestorOf(headerHit)), "Visible header filter button receives mouse hit testing");
        headerButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await Idle();
        Check(_filterPopup?.Search.IsVisible == true, "Clicking the actual header button opens the filter panel");
        _filterPopup!.CancelButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await _workspaceInput; await Idle();
        CaptureGridEvidence(Path.Combine(folder, "header-filter-buttons.png"));
        var region = CurrentView.Region!; var revision = _document.Snapshot.Revision;
        var candidates = DataFilterEngine.Candidates(_document.Snapshot, region, CurrentView.Filters, 1);
        var popup = new FilterPopup(SheetGrid, new Rect(100, 90, 40, 25), "部门", candidates, null);
        var pending = popup.ShowAsync(); await Idle();
        popup.Draft.Single(c => c.Key == CellValue.FromText("乙")).Selected = false;
        popup.Search.Text = "乙"; await Idle();
        Check(popup.Draft.Single(c => c.Key == CellValue.FromText("甲")).Selected, "Searching preserves selected values hidden from search results");
        popup.CancelButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        Check(await pending is null && CurrentView.Filters.Count == 0 && _document.Snapshot.Revision == revision, "Cancel discards filter draft and leaves document unchanged");
        popup = new(SheetGrid, new Rect(100, 90, 40, 25), "部门", candidates, null); pending = popup.ShowAsync(); await Idle();
        popup.Draft.Single(c => c.Key == CellValue.FromText("乙")).Selected = false;
        popup.ApplyButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        await ApplyFilterPopupResultAsync(1, (await pending)!);
        await ApplyFilterPopupResultAsync(2, new(FilterPopupAction.Apply, new(First: new(FilterOperator.GreaterOrEqual, "15"), Second: new(FilterOperator.LessOrEqual, "25"))));
        Check(CurrentView.Filters.Count == 2 && CurrentView.IsFilteredOut(_document.Snapshot, 0, 3) && CurrentView.IsFilteredOut(_document.Snapshot, 0, 4) && !CurrentView.IsFilteredOut(_document.Snapshot, 0, 6), "Department AND amount interval filters combine");
        Check(_document.Snapshot.Revision == revision, "Filter conditions do not create document history");
        await SortDataRegionAsync(region, [new(1), new(2, true)]);
        Check(_document.GetCellInput(0, 3, 4) == "first" && _document.GetCellInput(0, 4, 4) == "third" && _document.GetCellInput(0, 5, 4) == "fourth", "Two-level ordinal sort moves complete stable records including hidden rows");
        Check(_document.GetCellValue(0, 4, 5) == "40" && _document.GetCellInput(0, 8, 0) == "outside", "Sort shifts relative formulas and preserves outside cells");
        Check(CurrentView.IsFilteredOut(_document.Snapshot, 0, 3) && !CurrentView.IsFilteredOut(_document.Snapshot, 0, 4), "Sort re-evaluates filtered row addresses");
        Check(CurrentView.HeaderMarker(1) == "●↑1" && CurrentView.HeaderMarker(2) == "●↓2", "Header markers retain filter status and ordered sort direction together");
        await UndoAsync(false); Check(CurrentView.Region == region && CurrentView.IsFilteredOut(_document.Snapshot, 0, 4), "Undo retains region and recomputes filters");
        await UndoAsync(true); Check(_document.GetCellInput(0, 4, 4) == "third" && !CurrentView.IsFilteredOut(_document.Snapshot, 0, 4), "Redo recomputes filters on restored records");
        await ApplyFilterPopupResultAsync(1, new(FilterPopupAction.ClearColumn)); Check(CurrentView.Filters.Count == 1, "Clear column keeps the other column condition");
        await ApplyFilterPopupResultAsync(1, new(FilterPopupAction.ClearAll));
        Check(CurrentView.Filters.Count == 0 && _document.Snapshot.Sheets[0].Rows[5].Hidden, "Clear all filters preserves manual hidden metadata");
        var path = Path.Combine(folder, "sorted.xlsq"); _document.Save(path);
        using (var reopened = WorkbookDocument.Open(path, _services)) Check(reopened.GetCellInput(0, 4, 4) == "third", "Sorted records survive file save and reopen");
        CurrentView.FrozenRows = 3; CurrentView.FrozenColumns = 2; RefreshSheet(); await Idle();
        Check(_grid.HeaderButtons.Count > 0, "Header buttons exist in frozen reference rows");
        CaptureGridEvidence(Path.Combine(folder, "data-operations.png"));
        CurrentView.FrozenRows = CurrentView.FrozenColumns = 0;

        var surface = new FastGridSurface((p, text) => { try { _editor.SetCell(0, p.Row, p.Column, text); return true; } catch { return false; } }) { View = CurrentView };
        var window = new Window { Owner = this, Width = 1040, Height = 640, Content = surface.Element, Title = "FastWpfGrid validation", ShowInTaskbar = false };
        window.Show();
        try
        {
            surface.Bind(_document.Snapshot, 0, 100_000, 256); await Idle();
            Check(surface.ReadCount < 30_000, "Candidate bind reads visible cells without materializing 25.6M cells");
            Check(!surface.Control.AllowColumnResize && surface.Control.GetResizingColumn(new Point(160, 1)) is null,
                "Native unsaved resizing remains disabled; desktop owns dimension transactions");
            surface.SelectRange(new(0, 0, 99_999, 255), new(0, 0)); await Idle();
            Check(surface.Control.SelectedCellCount == 99_999L * 256 && surface.Control.SelectedRanges.Count == 1, "Huge candidate selection is one range and excludes manual hidden row");
            surface.SelectRange(new(5, 1, 6, 2), new(5, 1));
            Check(surface.Current == new CellPosition(6, 1) && surface.SelectionRanges.SequenceEqual([new GridRange(6, 1, 6, 2)]),
                "Candidate range beginning on a hidden row selects only requested visible addresses");
            surface.SelectRange(new(4, 1, 5, 2), new(5, 1));
            Check(surface.Current == new CellPosition(4, 1) && surface.SelectionRanges.SequenceEqual([new GridRange(4, 1, 4, 2)]),
                "Candidate range ending on a hidden row cannot turn into a header selection");
            surface.Select(new(9, 1)); Check(surface.BeginEdit(), "Candidate starts native raw-input editor");
            surface.Editor.Text = "中文草稿"; Check(surface.Commit(), "Candidate editor commits through controller");
            surface.Refresh(_document.Snapshot); Check(_document.GetCellInput(0, 9, 1) == "中文草稿", "Candidate commit changes the original domain address");
            surface.BeginEdit(); surface.Editor.Text = new string('x', 32768); var before = _document.Snapshot;
            Check(!surface.Commit(), "Invalid candidate draft is rejected"); surface.Select(new(10, 1));
            Check(ReferenceEquals(before, _document.Snapshot) && surface.Current == new CellPosition(9, 1) && surface.Editor.Text.Length == 32768, "Rejected draft retains original address text and history");
            surface.Cancel(); surface.View.FrozenRows = 3; surface.View.FrozenColumns = 2; surface.Refresh(_document.Snapshot); surface.Select(new(80, 20)); await Idle();
            Check(surface.Control.GetModelCellBounds(2, 1) is not null, "Candidate frozen header has stable model geometry after distant navigation");
            surface.Element.LayoutTransform = new ScaleTransform(1.25, 1.25); await Idle();
            Check(surface.Control.GetModelCellBounds(2, 1) is { Width: > 0 }, "Candidate geometry remains available at 125 percent zoom");
            await WaitForFastGridPaintAsync(surface.Control);
            SaveVisual(window, Path.Combine(folder, "fast-grid.png"));
        }
        finally { window.Close(); }
        GridSurface.Select(new(11, 1)); GridSurface.BeginEdit(); _grid.Editor.Text = new string('x', 32768);
        var rejected = _document.Snapshot.Revision;
        Check(!CommitEditors() && _grid.IsEditing && _grid.Editor.Text.Length == 32768 && rejected == _document.Snapshot.Revision,
            "Main workspace retains invalid drafts without history");
        _grid.Editor.Text = "主网格提交"; Check(CommitEditors() && _document.GetCellInput(0, 11, 1) == "主网格提交",
            "Main workspace commits to the original cell");
        File.WriteAllText(Path.Combine(folder, "data-operations-smoke.json"), JsonSerializer.Serialize(checks, new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllText(Path.Combine(folder, "data-operations-smoke-result.txt"), $"PASS: {checks.Count} data operations and candidate grid assertions. Native OS IME and cross-monitor DPI remain unverified.");
    }

    private static void SaveVisual(FrameworkElement visual, string path)
    {
        var bitmap = new RenderTargetBitmap((int)visual.ActualWidth, (int)visual.ActualHeight, 96, 96, PixelFormats.Pbgra32); bitmap.Render(visual);
        var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap)); using var stream = File.Create(path); encoder.Save(stream);
    }
}
