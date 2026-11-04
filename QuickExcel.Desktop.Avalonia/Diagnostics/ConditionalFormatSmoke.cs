using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls.Presenters;
using Avalonia.Automation;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunConditionalFormatSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static MenuItem Item(IEnumerable<object?> items, string title) => items.OfType<MenuItem>().Single(i => Equals(i.Header, title));
        static void ClickButton(Window window, string title) => window.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, title)).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        async Task Click(MenuItem item)
        {
            for (var i = 0; _working && i < 1000; i++) await Task.Delay(10);
            if (_working) throw new TimeoutException("Conditional menu input remained blocked.");
            item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await _operation; await Layout();
        }
        async Task Capture(Control control, string name)
        {
            await Layout(); using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(control); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        string? Fill(int row) => ConditionalPresentation.GetStyle(Document.Snapshot, Document.Snapshot.Sheets[0], row, 0).FillColor;
        try
        {
            _workspace.New(); ShowBook(); await Layout();
            var sheet = Document.Snapshot.ActiveSheetId;
            _workspace.Editor.SetCells(0, GridClipboard.Prepare("1\n2\n8\n8\n10", new(0, 0, 0, 0), false)); Synchronize();
            _grid.SelectRange(new(0, 0, 4, 0), new(0, 0)); await FlushEditsAsync();
            var button = this.GetVisualDescendants().OfType<DropDownButton>().Single(b => b.Name == "ConditionalFormatDropdown");
            Check(((TabItem)_ribbon.SelectedItem!).Header as string == "开始", "Quick conditional menu is on the Home ribbon");
            var flyout = (MenuFlyout)button.Flyout!;
            flyout.ShowAt(button); await Layout();
            Check(flyout.IsOpen && flyout.Items.OfType<MenuItem>().Count() == 8, "Dropdown opens with the eight reference categories");
            foreach (var title in new[] { "数据条(D)", "色阶(S)", "图标集(I)" })
                Check(Item(flyout.Items, title).IsEnabled && Item(flyout.Items, title).Items.Count > 0, "Visual rule submenu is enabled: " + title);
            var highlight = Item(flyout.Items, "突出显示单元格规则(H)");
            if (TopLevel.GetTopLevel(highlight) is { } popup) await Capture(popup, "conditional-menu");
            highlight.IsSubMenuOpen = true; await Layout();
            var greater = Item(highlight.Items, "大于…");
            Check(greater.IsEffectivelyVisible, "Highlight submenu exposes quick comparisons");
            if (TopLevel.GetTopLevel(greater) is { } sub) await Capture(sub, "conditional-highlight-menu");
            flyout.Hide();
            var revision = Document.Snapshot.Revision;
            _dialogDriver = async window =>
            {
                await Layout(); var form = (FieldsDialog)window;
                ((TextBox)form.Input("first")).Text = "=RAND()"; ClickButton(window, "确定"); await Layout();
                Check(window.IsVisible && Document.Snapshot.Revision == revision, "Unsupported formula keeps dialog open without writes");
                ((TextBox)form.Input("first")).Text = "5"; ((ComboBox)form.Input("style")).SelectedIndex = 2;
                Check(Document.Snapshot.Revision == revision, "Style preview is local to the dialog");
                await Capture(window, "conditional-greater-dialog"); ClickButton(window, "确定");
            };
            await Click(greater);
            Check(Fill(0) is null && Fill(2) == "FFC6EFCE", "Comparison shortcut applies chosen preset only to matching values");
            var ruleId = Document.Snapshot.Sheets[0].Data.ConditionalFormats.Single().ScriptId;
            await History(false); Check(Document.Snapshot.Sheets[0].Data.ConditionalFormats.IsEmpty, "Quick rule is a single undo entry");
            await History(true); Check(Document.Snapshot.Sheets[0].Data.ConditionalFormats.Single().ScriptId == ruleId, "Redo restores stable rule identity");
            revision = Document.Snapshot.Revision;
            _dialogDriver = async window => { await Layout(); ClickButton(window, "取消"); };
            await Click(greater); Check(Document.Snapshot.Revision == revision, "Cancel adds no rule or history");
            var ranking = Item(flyout.Items, "最前/最后规则(T)");
            _dialogDriver = async window =>
            {
                await Layout(); var form = (FieldsDialog)window; ((TextBox)form.Input("rank")).Text = "0"; ClickButton(window, "确定"); await Layout();
                Check(window.IsVisible && Document.Snapshot.Revision == revision, "Invalid rank is rejected before committing");
                ((TextBox)form.Input("rank")).Text = "1"; ClickButton(window, "确定");
            };
            await Click(Item(ranking.Items, "前 10 项…"));
            Check(Fill(4) == "FFFFC7CE" && Fill(2) == "FFC6EFCE", "Rank shortcut takes priority while preserving older rules");
            _grid.SelectRange(new(0, 0, 2, 0), new(0, 0));
            var clear = Item(flyout.Items, "清除规则(C)");
            await Click(Item(clear.Items, "清除所选单元格的规则"));
            Check(Fill(2) is null && Fill(4) == "FFFFC7CE" && Document.Snapshot.Sheets[0].Data.ConditionalFormats.All(r => r.Ranges[0].Top == 3), "Partial clear retains rules outside the selected cells");
            await History(false); Check(Fill(2) == "FFC6EFCE", "Partial clear undoes as one transaction");
            await Click(Item(clear.Items, "清除整个工作表的规则")); Check(Document.Snapshot.Sheets[0].Data.ConditionalFormats.IsEmpty, "Sheet clear removes every rule");
            await History(false); Check(Document.Snapshot.Sheets[0].Data.ConditionalFormats.Length == 2, "Sheet clear can be undone");
            _dialogDriver = async window => { await Layout(); Check(window.Title!.StartsWith("条件格式管理"), "Manage shortcut opens the rule manager"); ClickButton(window, "取消"); };
            await Click(Item(flyout.Items, "管理规则(R)…"));
            _dialogDriver = async window => { await Layout(); Check(window is FieldsDialog && window.Title == "条件格式规则", "New shortcut opens the full editor directly"); ClickButton(window, "取消"); };
            await Click(Item(flyout.Items, "新建规则(N)…"));
            await ExecuteAsync(new SetProtectionCommand(sheet, new())); revision = Document.Snapshot.Revision;
            try { await ClearConditionalFormatsAsync(true); throw new Exception("Protected clear succeeded"); }
            catch (InvalidOperationException) { Check(Document.Snapshot.Revision == revision, "Protection blocks quick rule changes"); }
            await ExecuteAsync(new SetProtectionCommand(sheet, null));
            _grid.Select(new(0, 0)); _grid.BeginEdit();
            var presenter = _grid.CellEditor.GetVisualDescendants().OfType<TextPresenter>().Single(); presenter.PreeditText = "ni";
            revision = Document.Snapshot.Revision;
            await QuickConditionalFormatAsync("大于", ConditionalPreset.GreaterThan); await ClearConditionalFormatsAsync(true);
            Check(Document.Snapshot.Revision == revision, "Composition gates quick creation and clearing"); presenter.PreeditText = null; _grid.Cancel();
            await Capture(this, "conditional-workspace");
            _workspace.New(); ShowBook(); sheet = Document.Snapshot.ActiveSheetId;
            var values = new[] { -20, -10, 0, 10, 20, 40 };
            _workspace.Editor.SetCells(0, GridClipboard.Prepare("渐变数据条\t实心数据条\t三色色阶\t三向箭头\t交通灯\t五象限图\n" +
                string.Join('\n', values.Select(v => string.Join('\t', Enumerable.Repeat(v.ToString(), 6)))), new(0, 0, 0, 0), false)); Synchronize();
            MenuItem Preset(IEnumerable<object?> items, string name) => items.OfType<MenuItem>().Single(m => AutomationProperties.GetName(m) == name);
            var bars = Item(flyout.Items, "数据条(D)"); var scales = Item(flyout.Items, "色阶(S)"); var icons = Item(flyout.Items, "图标集(I)");
            _grid.SelectRange(new(1, 0, 6, 0), new(1, 0)); await Click(Preset(Item(bars.Items, "渐变填充").Items, "蓝色"));
            Check(Document.Snapshot.Sheets[0].Data.ConditionalFormats.Single().DataBar?.Gradient == true, "Gradient gallery immediately creates a live data bar");
            await History(false); Check(Document.Snapshot.Sheets[0].Data.ConditionalFormats.IsEmpty, "Visual preset is one undo entry"); await History(true);
            _grid.SelectRange(new(1, 1, 6, 1), new(1, 1)); await Click(Preset(Item(bars.Items, "实心填充").Items, "绿色"));
            _grid.SelectRange(new(1, 2, 6, 2), new(1, 2)); await Click(Preset(scales.Items, "绿–黄–红"));
            _grid.SelectRange(new(1, 3, 6, 3), new(1, 3)); await Click(Preset(icons.Items, "三向箭头（彩色）"));
            _grid.SelectRange(new(1, 4, 6, 4), new(1, 4)); await Click(Preset(icons.Items, "三色交通灯"));
            _grid.SelectRange(new(1, 5, 6, 5), new(1, 5)); await Click(Preset(icons.Items, "五象限图"));
            Check(Document.Snapshot.Sheets[0].Data.ConditionalFormats.Length == 6, "All three visual categories apply through real menu handlers");
            var model = Document.Snapshot.Sheets[0];
            Check(ConditionalPresentation.GetVisual(Document.Snapshot, model, 1, 0)!.Bar is { End: 0, Color: "FFFF0000" } &&
                ConditionalPresentation.GetVisual(Document.Snapshot, model, 6, 1)!.Bar is { Gradient: false, End: 1 }, "Grid projection includes negative axes and solid positive bars");
            Check(ConditionalPresentation.GetStyle(Document.Snapshot, model, 1, 2).FillColor == "FFF8696B" && ConditionalPresentation.GetVisual(Document.Snapshot, model, 6, 3)!.Icon!.Index == 2, "Color scales and icon thresholds reach the high and low endpoints");
            var oldVisual = model.Data.ConditionalFormats.Single(r => r.DataBar?.Gradient == false);
            var editor = RuleFields(new(1, 1, 6, 1), oldVisual, out var read);
            ((CheckBox)editor.Input("show")).IsChecked = false; ((TextBox)editor.Input("border")).Text = "FF006100";
            var editedVisual = read(); Check(editedVisual.ScriptId == oldVisual.ScriptId && editedVisual.DataBar is { ShowValue: false, BorderColor: "FF006100" }, "Visual rule editor preserves identity and edits display and borders");
            var oldIcons = model.Data.ConditionalFormats.Single(r => r.IconSet?.Name == "3Arrows");
            var iconEditor = RuleFields(new(1, 3, 6, 3), oldIcons, out var readIcons);
            ((ComboBox)iconEditor.Input("icons")).SelectedIndex = ConditionalIconSet.Names.IndexOf("5Quarters");
            Check(readIcons().IconSet?.Thresholds.Length == 5, "Changing a three-icon rule to five icons supplies ordered thresholds");
            await ExecuteAsync(new SetConditionalFormatsCommand(sheet, model.Data.ConditionalFormats.Replace(oldVisual, editedVisual)));
            await Capture(this, "conditional-visuals-workspace");
            flyout.ShowAt(button); await Layout(); scales.IsSubMenuOpen = true; await Layout();
            if (TopLevel.GetTopLevel(Preset(scales.Items, "绿–黄–红")) is { } gallery) await Capture(gallery, "conditional-scales-gallery"); flyout.Hide();
            _grid.SelectRange(new(1, 1, 6, 1), new(1, 1));
            _dialogDriver = async window => { await Layout(); Check(window.Title == "数据条规则", "Custom data bar menu opens its own editor"); await Capture(window, "conditional-databar-editor"); ClickButton(window, "取消"); };
            await Click(Item(bars.Items, "其他数据条规则…"));
            var storage = new TypedWorkbookStorage(); var path = Path.Combine(directory, "conditional-visuals.xlsx"); storage.Export(path, Document.Snapshot);
            var reopened = storage.Open(path); Check(reopened.Sheets[0].Data.ConditionalFormats.Length == 6 && reopened.FileState!.CompatibilityDiagnostics.IsEmpty, "Visual workbook saves and reopens without compatibility loss");
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
