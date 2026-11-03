using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunRibbonFilterSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static T Named<T>(Control panel, string name) where T : Control => panel.GetVisualDescendants().OfType<T>().Single(c => c.Name == name);
        static void Click(Control panel, string name) => Named<Button>(panel, name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        async Task Capture(Control control, string name)
        {
            await Layout();
            using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        try
        {
            _workspace.New(); ShowBook(); Width = 950; Height = 900; await Layout();
            var ribbonHeight = _ribbon.Bounds.Height;
            foreach (var tab in _ribbon.Items.Cast<TabItem>())
            {
                _ribbon.SelectedItem = tab; await Layout();
                Check(Math.Abs(_ribbon.Bounds.Height - ribbonHeight) < .1, $"Ribbon height stays fixed on {tab.Header}");
                var scroll = (ScrollViewer)tab.Content!;
                Check(scroll.VerticalScrollBarVisibility == ScrollBarVisibility.Disabled && scroll.HorizontalScrollBarVisibility == ScrollBarVisibility.Auto,
                    $"Ribbon scrolls horizontally on {tab.Header}");
            }
            _ribbon.SelectedIndex = 0; await Layout();
            var firstScroll = (ScrollViewer)((TabItem)_ribbon.Items[0]!).Content!;
            Check(firstScroll.Extent.Width > firstScroll.Viewport.Width, "Narrow ribbon overflows horizontally without wrapping groups");
            firstScroll.Offset = new(firstScroll.Extent.Width - firstScroll.Viewport.Width, 0); await Layout();
            Check(firstScroll.Offset.X > 0, "Last ribbon groups remain reachable by horizontal scrolling");
            firstScroll.Offset = default;
            Width = 1400; await Layout(); Check(Math.Abs(_ribbon.Bounds.Height - ribbonHeight) < .1, "Resizing does not change ribbon height");
            foreach (var selector in new[] { _font, _fontSize })
            {
                for (var ancestor = selector.GetVisualParent(); ancestor is not null && ancestor != _ribbon; ancestor = ancestor.GetVisualParent())
                {
                    var origin = selector.TranslatePoint(default, ancestor)!.Value;
                    Check(origin.Y >= -.1 && origin.Y + selector.Bounds.Height <= ancestor.Bounds.Height + .1,
                        $"Font selector fits vertically inside {ancestor.GetType().Name}: y={origin.Y}, height={selector.Bounds.Height}, parentHeight={ancestor.Bounds.Height}");
                }
            }

            _workspace.Editor.SetCells(0, GridClipboard.Prepare("名称\t金额\n甲\t20\n乙\t10\n丙\t30", new(4, 0, 4, 0), false)); Synchronize();
            _grid.SelectRange(new(0, 1, WorkbookLimits.MaxRows - 1, 1), new(0, 1));
            await SortAsync(false);
            Check(Document.GetCellInput(0, 4, 0) == "名称" && Document.GetCellInput(0, 5, 0) == "乙", "Whole-column sort finds offset header and moves complete records");
            await History(false); Check(Document.GetCellInput(0, 5, 0) == "甲", "Whole-column sort undoes atomically");
            await ApplyStyleAsync(new(Bold: true));
            Check(Document.Snapshot.Styles[Document.Snapshot.Sheets[0].GetStyleId(7, 1)].Bold && Document.Snapshot.Sheets[0].Cells.Count < 30,
                "Whole-column formatting touches bounded used rows without a 20000-cell error");
            await History(false); await ClearAsync();
            Check(Document.GetCellInput(0, 5, 1) == "" && Document.GetCellInput(0, 5, 0) == "甲", "Whole-column clear affects selected data only");
            await History(false); await EstablishFilterAsync();
            Check(Document.Snapshot.Sheets[0].Data.Filter!.Range == new CellRectangle(4, 0, 7, 1), "Whole-column auto-filter bounds the actual record range");
            _grid.Select(new(4, 1)); await Capture(this, "ribbon-workspace");
            var before = Document.Snapshot.Revision;
            _filterDriver = async panel =>
            {
                await Layout();
                Check(_filterPopup is { IsOpen: true, IsLightDismissEnabled: true } && IsEnabled, "Filter is a light-dismiss floating menu with an enabled owner");
                Check(_filterPopup!.PlacementRect == _grid.GetCellBounds(new(4, 1)), "Filter is anchored to its column header");
                var list = Named<ListBox>(panel, "FilterCandidates");
                Check(list.ItemTemplate!.Build(null) is null, "Candidate template safely accepts a null data context");
                await Capture(panel, "filter-panel");
                Check(list.GetVisualDescendants().OfType<CheckBox>().Count() == 3, $"Candidate list realizes native checkboxes (items={list.ItemCount}, boxes={list.GetVisualDescendants().OfType<CheckBox>().Count()}, bounds={list.Bounds})");
                var search = Named<TextBox>(panel, "FilterSearch");
                search.Text = "missing"; await Layout();
                Check(list.ItemCount == 0 && !Named<CheckBox>(panel, "FilterSelectAll").IsEnabled, "Empty search results do not throw or enable select-all");
                search.Text = "20"; await Layout(); Named<CheckBox>(panel, "FilterSelectAll").IsChecked = false;
                search.Text = ""; await Layout();
                Check(list.GetVisualDescendants().OfType<CheckBox>().Count(c => c.IsChecked == true) == 2, "Search refresh preserves other candidate selections");
                Named<CheckBox>(panel, "FilterSelectAll").IsChecked = false; Click(panel, "FilterApply");
                Check(_filterPopup.IsOpen && Document.Snapshot.Revision == before && panel.GetVisualDescendants().OfType<TextBlock>().Any(t => t.IsVisible && t.Text == "请至少选择一个值或空白。"),
                    "Empty selection reports an inline error and leaves workbook unchanged");
                Named<ToggleButton>(panel, "FilterCustom").IsChecked = true; await Layout();
                Named<ComboBox>(panel, "FilterFirstOperator").SelectedIndex = (int)FilterOperator.Greater;
                var value = Named<TextBox>(panel, "FilterFirstValue"); value.Text = "invalid"; Click(panel, "FilterApply");
                Check(_filterPopup.IsOpen && Document.Snapshot.Revision == before, "Invalid numeric condition remains inside the floating panel");
                value.Text = "15"; Click(panel, "FilterApply");
            };
            await FilterAsync(1);
            Check(CurrentView.IsFilteredOut(Document.Snapshot, 0, 6) && !CurrentView.IsFilteredOut(Document.Snapshot, 0, 5), "Valid panel condition commits and filters the expected row");
            await History(false); Check(!CurrentView.IsFilteredOut(Document.Snapshot, 0, 6), "Filter application is one undo entry");
            for (var i = 0; i < 5; i++)
            {
                _filterDriver = async panel => { await Layout(); Named<TextBox>(panel, "FilterSearch").Text = "10"; panel.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Escape }); };
                await FilterAsync(1);
            }
            Check(_filterPopup is null, "Repeated open/search/Escape releases the popup");
            _filterDriver = async panel => { await Layout(); _filterPopup!.IsOpen = false; };
            await FilterAsync(1); Check(_filterPopup is null, "Light-dismiss completion releases the pending operation");
            // Candidate display can fail for unsupported number formats; clear and custom conditions stay available.
            _grid.SelectRange(new(5, 1, 7, 1), new(5, 1)); await ApplyStyleAsync(new(NumberFormatId: 0, NumberFormatCode: "[DBNum1]0"));
            _grid.Select(new(4, 1));
            _filterDriver = async panel =>
            {
                await Layout();
                Check(Named<ToggleButton>(panel, "FilterCustom").IsChecked == true && !Named<TextBox>(panel, "FilterSearch").IsEffectivelyEnabled,
                    "Candidate failures stay in the panel with custom conditions available");
                Check(Named<TextBox>(panel, "FilterFirstValue").IsFocused, "Condition mode receives keyboard focus when search is unavailable");
                Click(panel, "FilterCancel");
            };
            await FilterAsync(1); await History(false);
            // Exercise native virtualization and model-owned check states with many distinct values.
            _workspace.Editor.SetCells(0, Enumerable.Range(8, 300).Select(r => (r, 1, r.ToString()))); Synchronize();
            await ExecuteAsync(new SetAutoFilterCommand(Document.Snapshot.ActiveSheetId, new(new(4, 0, 307, 1))));
            _grid.Select(new(4, 1));
            _filterDriver = async panel =>
            {
                await Layout(); var list = Named<ListBox>(panel, "FilterCandidates"); var all = Named<CheckBox>(panel, "FilterSelectAll");
                all.IsChecked = false; list.ScrollIntoView(list.ItemCount - 1); await Layout();
                Check(list.GetVisualDescendants().OfType<CheckBox>().All(c => c.IsChecked == false), "Virtualized candidates retain deselection after scrolling");
                Named<TextBox>(panel, "FilterSearch").Text = "30"; await Layout(); all.IsChecked = true;
                Named<TextBox>(panel, "FilterSearch").Text = ""; await Layout(); Click(panel, "FilterCancel");
            };
            await FilterAsync(1);
            // A table filter must never move its total row, even through the menu's sort actions.
            var id = SheetId.New();
            var table = new WorksheetTable(1, "Sales", new(0, 0, 3, 1), [new(1, "名称"), new(2, "金额")], totalsRow: true, filter: new(new(0, 0, 2, 1)));
            var tableSheet = new WorksheetSnapshot(id, "表", new Dictionary<CellPosition, CellSnapshot>
            {
                [new(0, 0)] = new(CellInput.Parse("名称")), [new(0, 1)] = new(CellInput.Parse("金额")),
                [new(1, 0)] = new(CellInput.Parse("甲")), [new(1, 1)] = new(CellInput.Parse("20")),
                [new(2, 0)] = new(CellInput.Parse("乙")), [new(2, 1)] = new(CellInput.Parse("10")),
                [new(3, 0)] = new(CellInput.Parse("总计")), [new(3, 1)] = new(CellInput.Parse("30"))
            }, data: new() { Tables = [table] });
            var tablePath = Path.Combine(directory, "table.xlsq"); new TypedWorkbookStorage().Save(tablePath, new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [tableSheet]));
            await OpenPathAsync(tablePath); _grid.Select(new(0, 1));
            _filterDriver = async panel => { await Layout(); Click(panel, "FilterAscending"); }; await FilterAsync(1);
            Check(Document.GetCellInput(0, 1, 0) == "乙" && Document.GetCellInput(0, 0, 0) == "名称" && Document.GetCellInput(0, 3, 0) == "总计",
                "Table menu sorting preserves header and totals while moving complete records");
            await History(false); Check(Document.GetCellInput(0, 1, 0) == "甲", "Table menu sorting undoes atomically");
            _grid.SelectRange(new(0, 1, WorkbookLimits.MaxRows - 1, 1), new(100, 1)); await SortAsync(false);
            Check(Document.GetCellInput(0, 1, 0) == "乙" && Document.GetCellInput(0, 3, 0) == "总计", "Whole-column sort recovers table boundaries when the active cell is outside it");
            await History(false); _grid.Select(new(0, 1));
            _filterDriver = async panel => { await Layout(); _workspace.New(); Synchronize(); };
            await FilterAsync(1); Check(_filterPopup is null && Document.Snapshot.Sheets[0].Data.Filter is null, "Document replacement invalidates an open filter without applying stale choices");
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
