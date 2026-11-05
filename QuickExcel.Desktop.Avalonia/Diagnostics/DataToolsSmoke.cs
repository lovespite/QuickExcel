using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunDataToolsSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        async Task Capture(Control control, string name)
        {
            await Layout(); using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        static Button Button(Control control, string label) => control.GetVisualDescendants().OfType<Button>().Single(b => b.Content as string == label);
        static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(global::Avalonia.Controls.Button.ClickEvent));
        static void Set(FieldsDialog dialog, string key, string text) => ((TextBox)dialog.Input(key)).Text = text;
        void AcceptPreview()
        {
            _dialogDriver = async window =>
            {
                await Layout(); Check(window.Title == "数据工具预览", "Data tools require a reviewable preview");
                foreach (var check in window.GetVisualDescendants().OfType<CheckBox>().Where(c => c.IsVisible)) check.IsChecked = true;
                Click(Button(window, "应用"));
            };
        }
        async Task Seed(params string[][] rows)
        {
            _workspace.New(); ShowBook(); var id = Document.Snapshot.ActiveSheetId;
            await ExecuteAsync(new SetCellsCommand(rows.SelectMany((values, r) => values.Select((value, c) =>
                new CellEdit(new(id, r, c), value.StartsWith('=') ? CellContent.FromFormula(value) :
                    double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var number) ? CellContent.FromLiteral(CellValue.FromNumber(number)) : CellContent.FromLiteral(CellValue.FromText(value)))))));
            await CalculatePendingAsync(); await Layout();
        }
        async Task ClickOutline(string name)
        {
            Click(_grid.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == name)); await _operation; await Layout();
        }
        try
        {
            Width = 1200; Height = 880;
            await Seed(["Cat.", "Count"], ["A", "1"], ["A", "2"], ["B", "3"], ["B", "4"], ["B", "1"], ["C", "6"], ["C", "7"], ["C", "8"]);
            var tab = _ribbon.Items.Cast<TabItem>().Single(t => t.Header as string == "数据"); _ribbon.SelectedItem = tab; await Layout();
            foreach (var label in new[] { "分列", "快速填充  Ctrl+E", "删除重复值", "数据验证", "模拟分析", "预测", "分类汇总" })
                Check(_ribbon.GetVisualDescendants().OfType<Button>().Any(b => AutomationProperties.GetName(b) == label), "Ribbon exposes " + label);
            await Capture(_ribbon, "data-ribbon");
            _grid.Select(new(1, 0));
            _dialogDriver = async window =>
            {
                await Layout(); var dialog = (FieldsDialog)window;
                Check(((ComboBox)dialog.Input("group")).Items.Cast<string>().First() == "Cat. (A)", "Subtotal fields use header names");
                Check(dialog.Boolean("replace") && dialog.Boolean("below") && !dialog.Boolean("pages"), "Subtotal defaults match the requested dialog");
                await Capture(window, "subtotal-dialog");
                _dialogDriver = async preview => { await Layout(); Click(Button(preview, "确定")); };
                Click(Button(window, "确定"));
            };
            await SubtotalsAsync(); await Layout();
            Check(Document.Snapshot.Sheets[0].GetCell(12, 1)!.Value.NumberValue == 32, "Subtotal action creates the sample total 32");
            Check(!_book.GetVisualDescendants().OfType<ComboBox>().Any(b => AutomationProperties.GetName(b) == "工作表分组"), "The old outline dropdown is removed");
            Check(_grid.GetVisualDescendants().OfType<Button>().Count(b => (AutomationProperties.GetName(b) ?? "").StartsWith("显示大纲级别")) == 3, "Tree has level 1/2/3 buttons");
            var origin = _grid.GetCellBounds(new(0, 0))!.Value;
            Check(origin.X > 100 && _grid.HitTestCell(origin.Center) == new CellPosition(0, 0), "Outline gutter shifts cells and preserves hit testing");
            await Capture(_grid, "subtotal-tree-expanded");
            await ClickOutline("显示大纲级别 2");
            Check(Document.Snapshot.Sheets[0].Rows[1].Hidden && !Document.Snapshot.Sheets[0].Rows[3].Hidden, "Level 2 shows group summaries and hides detail");
            await Capture(_grid, "subtotal-tree-level2");
            await ClickOutline("展开第 2–3 行（2 级）");
            Check(!Document.Snapshot.Sheets[0].Rows[1].Hidden && Document.Snapshot.Sheets[0].Rows[4].Hidden, "Plus expands only its own branch");
            await ClickOutline("显示大纲级别 1"); await Capture(_grid, "subtotal-tree-level1");
            Check(Document.Snapshot.Sheets[0].Rows[3].Hidden && !Document.Snapshot.Sheets[0].Rows.GetValueOrDefault(12)!.Hidden, "Level 1 leaves grand total visible");
            await ClickOutline("显示大纲级别 3"); await ZoomAsync(150); await FreezeAsync(1, 0); _grid.ScrollTo(0, 20); await Capture(_grid, "subtotal-tree-zoom-freeze-scroll");
            var toggle = _grid.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == "折叠第 2–3 行（2 级）");
            var summary = _grid.GetCellBounds(new(3, 0))!.Value;
            Check(Math.Abs(toggle.TranslatePoint(new(toggle.Bounds.Width / 2, toggle.Bounds.Height / 2), _grid)!.Value.Y - summary.Center.Y) < 1, "Tree toggle aligns with its summary after zoom, freeze and scroll");
            await ZoomAsync(100); await FreezeAsync(0, 0); _grid.ScrollTo(0, 0);
            await CalculatePendingAsync(); var path = Path.Combine(directory, "subtotal.xlsx"); new TypedWorkbookStorage().Export(path, Document.Snapshot);
            Check(WorksheetOutline.Groups(new TypedWorkbookStorage().Open(path).Sheets[0], SheetAxis.Row).Count == 4, "Tree outline survives XLSX save/reopen");
            _grid.Select(new(0, 0));
            _dialogDriver = async window => { await Layout(); _dialogDriver = async p => { await Layout(); Click(Button(p, "确定")); }; Click(Button(window, "全部删除")); };
            await SubtotalsAsync(); Check(!WorksheetOutline.HasGroups(Document.Snapshot.Sheets[0]), "Remove All removes subtotals and tree");
            Check(_grid.GetCellBounds(new(0, 0))!.Value.X == 52, "Empty outline returns header width to normal");
            await History(false); Check(WorksheetOutline.HasGroups(Document.Snapshot.Sheets[0]), "Undo restores the subtotal tree");

            await Seed(["Ada Lovelace", "Ada"], ["Grace Hopper"], ["Alan Turing"]); _grid.Select(new(0, 1));
            _dialogDriver = async window => { await Layout(); var dialog = (FieldsDialog)window; Set(dialog, "range", "A1:B3"); AcceptPreview(); Click(Button(window, "确定")); };
            _grid.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.E, KeyModifiers = KeyModifiers.Control, Source = _grid }); await _operation;
            Check(Document.Snapshot.Sheets[0].GetCell(1, 1)!.Value.TextValue == "Grace", "Ctrl+E runs flash fill through the actual shortcut");
            await History(false); Check(Document.Snapshot.Sheets[0].GetCell(1, 1) is null, "Flash fill is one undo step");

            var splitRows = Enumerable.Range(0, 220).Select(i => new[] { $"xx-{i}|007|02/01/2026|omit" + (i == 219 ? "|25" : "") }).ToArray();
            splitRows[0] = [splitRows[0][0], "existing"];
            await Seed(splitRows); _grid.SelectRange(new(0, 0, 219, 0), new(0, 0));
            var splitBook = Document.Snapshot;
            static async Task Until(Func<bool> predicate)
            {
                var deadline = DateTime.UtcNow.AddSeconds(5);
                while (!predicate() && DateTime.UtcNow < deadline) await Task.Delay(15);
                if (!predicate()) throw new TimeoutException("Timed out waiting for the split wizard.");
            }
            _dialogDriver = async w =>
            {
                await Layout();
                T Named<T>(string name) where T : Control => w.GetVisualDescendants().OfType<T>().Single(c => AutomationProperties.GetName(c) == name);
                RadioButton Radio(string text) => w.GetVisualDescendants().OfType<RadioButton>().Single(c => Equals(c.Content, text));
                var next = Button(w, "下一步"); var finish = Button(w, "完成");
                Check(!Button(w, "上一步").IsEnabled && next.IsEnabled, "Wizard opens on the raw source preview");
                await Capture(w, "split-step-1");
                Click(next); await Layout();
                w.GetVisualDescendants().OfType<CheckBox>().Single(c => Equals(c.Content, "Tab 键")).IsChecked = false;
                w.GetVisualDescendants().OfType<CheckBox>().Single(c => Equals(c.Content, "其他")).IsChecked = true;
                Check(!next.IsEnabled, "Missing delimiter disables navigation with an inline error");
                Named<TextBox>("其他分隔符").Text = "|"; await Layout();
                Check(next.IsEnabled, "Custom delimiter refreshes the live preview");
                await Capture(w, "split-step-2");
                Click(next); await Until(() => w.Title!.Contains("第 3 步")); await Layout();
                Click(Named<Button>("第 2 列，常规")); Radio("文本").IsChecked = true;
                Click(Named<Button>("第 3 列，常规")); Named<ComboBox>("日期顺序").SelectedIndex = 2;
                Click(Named<Button>("第 4 列，常规")); Radio("不导入此列（跳过）").IsChecked = true;
                Click(Button(w, "上一步")); Click(next); await Until(() => w.Title!.Contains("第 3 步")); await Layout();
                Click(Named<Button>("第 2 列，文本"));
                Check(Radio("文本").IsChecked == true && Named<Button>("第 5 列，常规").IsVisible, "Column formats survive navigation and include fields beyond the sample");
                await Capture(w, "split-step-3");
                var scroll = w.GetVisualDescendants().OfType<ScrollViewer>().Single(s => s.Extent.Height > 2000);
                scroll.Offset = new(0, scroll.Extent.Height); await Layout();
                Check(scroll.Offset.Y > 0, "Split table scrolls to the last sampled row");
                Named<TextBox>("目标区域").Text = "A1:B2"; Click(finish); await Layout();
                Check(w.GetVisualDescendants().OfType<TextBlock>().Any(t => t.Text?.Contains("单个起始单元格") == true), "Invalid target is reported inside the wizard");
                Named<TextBox>("目标区域").Text = "$A$1"; Click(finish);
                CheckBox? Confirmation() => w.GetVisualDescendants().OfType<CheckBox>().FirstOrDefault(c => c.IsVisible && c.Content is string label && label.StartsWith("确认覆盖"));
                await Until(() => Confirmation() != null);
                Check(!finish.IsEnabled && Document.Snapshot.Revision == splitBook.Revision, "Existing output requires explicit consent before any edit");
                Confirmation()!.IsChecked = true; Click(finish);
            };
            await SplitTextAsync();
            var splitSheet = Document.Snapshot.Sheets[0];
            Check(splitSheet.GetCell(0, 1)!.Value == CellValue.FromText("007") && splitSheet.GetCell(0, 2)!.Value == CellValue.FromNumber(46024) &&
                splitSheet.GetCell(219, 3)!.Value == CellValue.FromNumber(25), "Column-specific text, DMY, skip and general formats are applied to all rows");
            await History(false);
            Check(Document.Snapshot.Sheets[0].GetCell(0, 1)!.Value == CellValue.FromText("existing") && Document.Snapshot.Sheets[0].GetCell(219, 3) is null,
                "One undo restores the source and the overwritten output");
            _dialogDriver = async w =>
            {
                await Layout(); w.GetVisualDescendants().OfType<RadioButton>().Single(c => Equals(c.Content, "固定宽度")).IsChecked = true;
                Click(Button(w, "下一步")); await Layout();
                w.GetVisualDescendants().OfType<TextBox>().Single(c => AutomationProperties.GetName(c) == "固定宽度分界").Text = "2,3";
                Click(Button(w, "下一步")); await Until(() => w.Title!.Contains("第 3 步")); await Layout();
                Check(w.GetVisualDescendants().OfType<Button>().Count(b => AutomationProperties.GetName(b)?.EndsWith("列，常规") == true) == 3, "Fixed width exposes three selectable fields");
                Click(Button(w, "取消"));
            };
            await SplitTextAsync();
            var bulk = DataCleaning.SplitText(splitBook, splitBook.ActiveSheetId, new(0, 0, 219, 0), new(0, 5), new("|"));
            _dialogDriver = async w =>
            {
                await Layout();
                var list = w.GetVisualDescendants().OfType<ListBox>().Single();
                list.ScrollIntoView(list.ItemCount - 1); await Layout();
                Check(list.ContainerFromIndex(list.ItemCount - 1) is not null, "Shared cleaning preview recycles rows through the last change without crashing");
                list.ScrollIntoView(0); await Layout(); Click(Button(w, "取消"));
            };
            await ApplyCleaningAsync(bulk);
            await Seed(["ID", "Name"], ["1", "first"], ["1", "duplicate"], ["2", "second"]); _grid.SelectRange(new(0, 0, 3, 0), new(0, 0));
            _dialogDriver = async w => { await Layout(); Check(((FieldsDialog)w).Boolean("expand"), "Duplicate removal expands records by default"); AcceptPreview(); Click(Button(w, "确定")); }; await RemoveDuplicatesAsync();
            Check(Document.Snapshot.Sheets[0].GetCell(2, 1)!.Value.TextValue == "second" && Document.Snapshot.Sheets[0].GetCell(3, 1) is null, "Removing duplicates keeps adjacent fields attached to rows");

            await Seed(["1", "=A1*3+2"]); _grid.Select(new(0, 1));
            _dialogDriver = async w => { await Layout(); Set((FieldsDialog)w, "target", "32"); AcceptPreview(); Click(Button(w, "确定")); }; await GoalSeekAsync();
            Check(Math.Abs(Document.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue - 10) < 0.000001, "Goal seek applies only its reviewed solution");
            _dialogDriver = async w => { await Layout(); AcceptPreview(); Click(Button(w, "确定")); }; await DataTableAsync();
            Check(Document.Snapshot.Sheets[0].GetCell(1, 4)!.Value.NumberValue == 5, "Simulation table dialog produces calculated results");
            _settings.Save(_settings.Value with { Calculation = QuickExcel.Presentation.DesktopCalculationMode.Partial });
            Document.SetCell(0, 0, 1, "=A1*4"); await CalculatePendingAsync();
            Check(Math.Abs(Document.Snapshot.Sheets[0].GetCell(0, 1)!.Value.NumberValue - 40) < 1e-6 &&
                Document.GetCellValue(0, 1, 4) == "5", "Partial calculation updates formulas but retains the trial results");
            await RefreshDataTablesAsync();
            Check(Document.GetCellValue(0, 1, 4) == "4", "Separate data table refresh uses the original trial inputs");
            _settings.Save(_settings.Value with { Calculation = QuickExcel.Presentation.DesktopCalculationMode.Automatic });
            Document.SetCell(0, 0, 1, "=A1*6"); await CalculatePendingAsync();
            Check(Document.GetCellValue(0, 1, 4) == "6", "Automatic calculation also refreshes trial results");
            await History(false);
            Check(Document.GetCellInput(0, 0, 1) == "=A1*4" && Document.GetCellValue(0, 1, 4) == "4",
                "One undo restores both the input formula and its automatically refreshed trial outputs");
            await History(true);
            // Repeat the normal creation workflow after history replaced the output rectangle.
            _grid.Select(new(0, 1)); _dialogDriver = async w => { await Layout(); AcceptPreview(); Click(Button(w, "确定")); }; await DataTableAsync();
            Document.SetCell(0, 1, 4, "123"); Document.SetCell(0, 0, 1, "=A1*7"); await CalculatePendingAsync();
            Check(Document.GetCellValue(0, 1, 4) == "123", "Recalculation preserves a user's replacement of a simulation output");
            await Seed(["1", "2", "=A1*B1"]); _grid.Select(new(0, 2));
            _dialogDriver = async w => { await Layout(); AcceptPreview(); Click(Button(w, "确定")); }; await ScenariosAsync();
            Check(Document.Snapshot.Sheets[0].GetCell(1, 6)!.Value.NumberValue == 200 && Document.Snapshot.Sheets[0].GetCell(2, 6)!.Value.NumberValue == 375, "Scenario dialog computes named multi-input summaries");
            await Seed(["Time", "Value"], ["1", "3"], ["2", "5"], ["3", "7"]); _grid.Select(new(0, 0));
            _dialogDriver = async w => { await Layout(); AcceptPreview(); Click(Button(w, "确定")); }; await ForecastAsync();
            Check(Document.Snapshot.Sheets[0].GetCell(4, 5)!.Value.NumberValue == 9, "Forecast dialog extends the historical trend");
            _grid.Select(new(1, 1));
            _dialogDriver = async w => { await Layout(); Click(Button(w, "确定")); }; await EditValidationAsync();
            Check(Document.Snapshot.Sheets[0].Data.Validations.Length == 1, "Data validation remains functional in the data tools group");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
