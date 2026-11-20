using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
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
        async Task Capture(Control control, string name, double scale = 1)
        {
            var background = _ribbon.Background;
            try
            {
                // Isolated ribbon captures need the same opaque background as the containing window.
                if (ReferenceEquals(control, _ribbon)) _ribbon.Background = Background;
                await Layout();
                using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width * scale), (int)Math.Ceiling(control.Bounds.Height * scale)), new Vector(96 * scale, 96 * scale));
                bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
            }
            finally { _ribbon.Background = background; }
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
                var buttons = scroll.GetVisualDescendants().OfType<ContentControl>().Where(b => b is Button or SplitButton && b.Classes.Contains("ribbon-command")).ToArray();
                await Capture(_ribbon, $"ribbon-tab-{_ribbon.SelectedIndex}");
                Check(buttons.Length > 0 && buttons.All(b => !string.IsNullOrWhiteSpace(AutomationProperties.GetName(b)) && ToolTip.GetTip(b) is string { Length: > 0 }),
                    $"Every ribbon command has an accessible name and tooltip on {tab.Header}");
                foreach (var button in buttons)
                {
                    // Color split buttons also contain a 14px swatch and a separate arrow target.
                    var maxWidth = button.Name is "FontColorDropdown" or "FillColorDropdown" ? 84 : button is SplitButton ? 68 : 64;
                    var labeled = button.Content is RibbonCommandContent command && command.Label.IsVisible;
                    Check((labeled || button.Bounds.Width <= maxWidth) && button.Bounds.Height <= 28,
                        $"Ribbon command fits its row: {AutomationProperties.GetName(button)} ({button.Bounds.Width} x {button.Bounds.Height}, labeled={labeled})");
                    bool Fits(Control control, Control parent)
                    {
                        var topLeft = control.TranslatePoint(default, parent)!.Value;
                        var bottomRight = control.TranslatePoint(new(control.Bounds.Width, control.Bounds.Height), parent)!.Value;
                        return control.Bounds.Width > 0 && control.Bounds.Height > 0 && topLeft.X >= -.1 && topLeft.Y >= -.1 &&
                            bottomRight.X <= parent.Bounds.Width + .1 && bottomRight.Y <= parent.Bounds.Height + .1;
                    }
                    if (button is SplitButton split)
                    {
                        var primary = Named<Button>(split, "PART_PrimaryButton"); var secondary = Named<Button>(split, "PART_SecondaryButton");
                        Check(Fits(primary, split) && Fits(secondary, split) && split.Content is Control splitContent && Fits(splitContent, primary),
                            $"Split button content and both targets fit: {AutomationProperties.GetName(button)}");
                        split.IsEnabled = false;
                        Check(!primary.IsEffectivelyEnabled && !secondary.IsEffectivelyEnabled,
                            $"Both split button targets inherit disabled state: {AutomationProperties.GetName(button)}");
                        split.IsEnabled = true;
                    }
                    if (button.Name == "BorderDropdown")
                    {
                        Check(button.Content is BorderPresetIcon border && border.Bounds.Size == new Size(20, 20) && Fits(border, button) && !border.IsHitTestVisible,
                            "Border sample fits the ribbon and leaves clicks to the primary button");
                        continue;
                    }
                    if (button.Name == "TextRotationDropdown")
                    {
                        Check(button.Content is Border { Child: TextBlock sample } preview && preview.Bounds.Size == new Size(20, 20) &&
                              Fits(preview, button) && !string.IsNullOrWhiteSpace(sample.Text) && sample.RenderTransform is RotateTransform { Angle: -45 },
                            "Text orientation has a visible rotated text sample inside its ribbon button");
                        continue;
                    }
                    if (button.Name == "ConditionalFormatDropdown")
                    {
                        Check(button.Content is ConditionalFormatIcon preview && preview.Bounds.Size == new Size(20, 20) && Fits(preview, button),
                            "Conditional format uses a compact 20px preview without a command label");
                        continue;
                    }
                    Check(button.GetVisualDescendants().OfType<PathIcon>().Count(i => i.Classes.Contains("fluent-ribbon-icon")) == 1,
                        $"Ribbon command has one vector icon: {AutomationProperties.GetName(button)}");
                    var icon = button.GetVisualDescendants().OfType<PathIcon>().Where(i => i.Classes.Contains("fluent-ribbon-icon")).Single();
                    var origin = icon.TranslatePoint(default, button)!.Value;
                    Check(icon.Bounds.Size == new Size(20, 20) && origin.X >= 0 && origin.Y >= 0 &&
                          origin.X + 20 <= button.Bounds.Width && origin.Y + 20 <= button.Bounds.Height,
                        $"20px icon fits button: {AutomationProperties.GetName(button)}");
                    var path = icon.GetVisualDescendants().OfType<global::Avalonia.Controls.Shapes.Path>().Single();
                    Check(path.Stretch == Stretch.None && Equals(path.Fill, icon.Foreground),
                        $"Vector preserves SVG margins and inherits foreground: {AutomationProperties.GetName(button)}");
                    // A color swatch can contain a slash for "no fill"; it is not a command label.
                    var labels = button.Content is StackPanel content ? content.Children.OfType<TextBlock>().Where(t => t.IsVisible).ToArray() : [];
                    Check(!ReferenceEquals(icon.Data, RibbonIcons.Create(RibbonIcon.Default).Data),
                        $"Command has a specific icon: {AutomationProperties.GetName(button)}");
                    Check(labeled ? labels.Length == 1 && labels[0].Text == AutomationProperties.GetName(button) && Fits(labels[0], button)
                        : labels.Length == 0 && button.Content is not string && button.Content is not TextBlock,
                        $"Command label follows the page layout: {AutomationProperties.GetName(button)}");
                }
            }
            await CheckResponsiveRibbonAsync(Check, (control, name) => Capture(control, name));
            _ribbon.SelectedIndex = 0; await Layout();
            foreach (var name in new[] { "FontColorDropdown", "FillColorDropdown", "BorderDropdown" })
                Check(Named<SplitButton>(_ribbon, name).Classes.Contains("ribbon-command"), "Split command is included in ribbon diagnostics: " + name);
            var bold = _ribbon.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == "加粗");
            var boldIcon = bold.GetVisualDescendants().OfType<PathIcon>().Where(i => i.Classes.Contains("fluent-ribbon-icon")).Single();
            bold.Foreground = Brushes.Red; await Layout();
            Check(Equals(boldIcon.Foreground, Brushes.Red) && Equals(boldIcon.GetVisualDescendants().OfType<global::Avalonia.Controls.Shapes.Path>().Single().Fill, Brushes.Red),
                "Icon foreground follows button styling");
            bold.ClearValue(ForegroundProperty); await Layout();
            bold.IsEnabled = false; await Layout();
            Check(!boldIcon.IsEffectivelyEnabled && Equals(boldIcon.Foreground, boldIcon.GetVisualDescendants().OfType<global::Avalonia.Controls.Shapes.Path>().Single().Fill),
                "Disabled icon follows the disabled button and foreground");
            bold.IsEnabled = true; await Layout();
            var firstScroll = (ScrollViewer)((TabItem)_ribbon.Items[0]!).Content!;
            Check(firstScroll.Extent.Width > firstScroll.Viewport.Width, "Narrow ribbon overflows horizontally without wrapping groups");
            firstScroll.Offset = new(firstScroll.Extent.Width - firstScroll.Viewport.Width, 0); await Layout();
            Check(firstScroll.Offset.X > 0, "Last ribbon groups remain reachable by horizontal scrolling");
            firstScroll.Offset = default;
            Width = 1400; await Layout(); Check(Math.Abs(_ribbon.Bounds.Height - ribbonHeight) < .1, "Resizing does not change ribbon height");
            await Capture(_ribbon, "ribbon-home-150pct", 1.5);
            await Capture(_ribbon, "ribbon-home-200pct", 2);
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
            // A selected header row must include the connected data below it, as in a newly created filter.
            _workspace.New(); ShowBook();
            _workspace.Editor.SetCells(0, GridClipboard.Prepare("X\ta\tv\tc\n甲\tX\tX\tvv\n乙\tX\tX\tf\n丙\tX\tX\tg\n丁\tX\tX\tg\n戊\tX\tX\tg", new(0, 0, 0, 0), false)); Synchronize();
            _grid.SelectRange(new(0, 0, 0, 3), new(0, 3)); await EstablishFilterAsync();
            Check(Document.Snapshot.Sheets[0].Data.Filter!.Range == new CellRectangle(0, 0, 5, 3),
                "Selecting only the header row creates a filter covering the connected records");
            _filterDriver = async panel =>
            {
                await Layout(); var list = Named<ListBox>(panel, "FilterCandidates");
                Check(list.ItemCount == 3 && list.GetVisualDescendants().OfType<CheckBox>().Select(c => c.Content?.ToString()).Order().SequenceEqual(new[] { "f", "g", "vv" }),
                    "A new header-row filter displays all distinct text candidates");
                await Capture(panel, "header-row-filter"); Click(panel, "FilterAscending");
            };
            await FilterAsync(3);
            Check(Enumerable.Range(1, 5).Select(r => Document.GetCellInput(0, r, 0)).SequenceEqual(new[] { "乙", "丙", "丁", "戊", "甲" }) && Document.GetCellInput(0, 0, 3) == "c",
                "New filter menu ascending sort moves complete records and keeps the header fixed");
            await History(false);
            Check(Document.GetCellInput(0, 1, 3) == "vv" && Document.GetCellInput(0, 2, 3) == "f", "New filter menu sort undoes as one operation");
            _filterDriver = async panel => { await Layout(); Click(panel, "FilterDescending"); }; await FilterAsync(3);
            Check(Enumerable.Range(1, 5).Select(r => Document.GetCellInput(0, r, 0)).SequenceEqual(new[] { "甲", "丙", "丁", "戊", "乙" }) && Document.GetCellInput(0, 0, 3) == "c",
                "New filter menu descending sort keeps equal keys stable and the header fixed");
            await History(false); await SortAsync(false);
            Check(Document.GetCellInput(0, 1, 0) == "乙" && Document.GetCellInput(0, 0, 3) == "c", "Ribbon sort uses the newly inferred filter data range");
            await History(false);
            _filterDriver = async panel =>
            {
                await Layout(); Named<CheckBox>(panel, "FilterSelectAll").IsChecked = false;
                Named<TextBox>(panel, "FilterSearch").Text = "f"; await Layout();
                Named<CheckBox>(panel, "FilterSelectAll").IsChecked = true; Click(panel, "FilterApply");
            };
            await FilterAsync(3);
            Check(CurrentView.IsFilteredOut(Document.Snapshot, 0, 1) && !CurrentView.IsFilteredOut(Document.Snapshot, 0, 2) && CurrentView.IsFilteredOut(Document.Snapshot, 0, 3),
                "Selecting a candidate from a new header-row filter hides only unmatched records");
            await History(false);
            await ExecuteAsync(new SetAutoFilterCommand(Document.Snapshot.ActiveSheetId, new(new(0, 0, 0, 3))));
            before = Document.Snapshot.Revision;
            _filterDriver = async panel => { await Layout(); Click(panel, "FilterAscending"); };
            var rejectedEmptySort = false;
            try { await FilterAsync(3); }
            catch (ArgumentException ex) when (ex.Message == UiText.Get("请选择包含多行记录的数据区域。")) { rejectedEmptySort = true; }
            Check(rejectedEmptySort && Document.Snapshot.Revision == before, "An existing header-only filter reports insufficient records instead of silently sorting nothing");
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
