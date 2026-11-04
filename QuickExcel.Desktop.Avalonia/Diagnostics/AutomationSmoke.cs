using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
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
    private async Task RunAutomationSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label + ": " + _status.Text); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        async Task Click(Button button)
        {
            button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await _operation; await Layout();
        }
        static void DoubleClick(Control control)
        {
            var pointer = new Pointer(1, PointerType.Mouse, true);
            var press = new PointerPressedEventArgs(control, pointer, control, new(4, 4), 0,
                new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed), KeyModifiers.None, 2);
            control.RaiseEvent(new TappedEventArgs(DoubleTappedEvent, press));
        }
        Control Dropdown() => _automationActions.GetTemplateDescendants().OfType<Popup>().Single().Child!;
        StackPanel Item(string name) => Dropdown().GetVisualDescendants().OfType<StackPanel>()
            .Single(p => p.Children.OfType<TextBlock>().FirstOrDefault()?.Text == name);
        async Task Search(string keyword) { _automationSearch.Text = keyword; await Layout(); }
        async Task OpenDropdown()
        {
            _automationActions.Focus(); _automationActions.IsDropDownOpen = true; await Layout();
        }
        async Task Idle()
        {
            // Drain deferred recalculation before sending focus-sensitive keyboard gestures.
            for (var attempt = 0; attempt < 100; attempt++)
            {
                await _operation; await Layout();
                if (!_working && !_workspace.IsBusy)
                {
                    if (!_calculationPending) return;
                    await PerformAsync(CalculatePendingAsync);
                }
                await Task.Delay(10);
            }
            throw new InvalidOperationException("The workspace did not become idle.");
        }
        try
        {
            _workspace.New(); ShowBook();
            var tab = _ribbon.Items.Cast<TabItem>().Single(t => Equals(t.Header, "自动执行"));
            _ribbon.SelectedItem = tab; await Layout(); var ribbonHeight = _ribbon.Bounds.Height;
            Check(_automationEmpty.IsVisible && !_automationRun!.IsEnabled, "Empty workbook explains the empty action list and disables Execute");
            const string source = """
                function main(workbook) {
                    const sheet=workbook.getActiveWorksheet(), a=sheet.getRange('A1'), b=sheet.getRange('B1');
                    a.setValue((Number(a.getValue()) || 0) + 1);
                    b.setValue((Number(b.getValue()) || 0) + 10);
                }
                """;
            var dictionary = new Dictionary<string, ScriptModule>
            {
                ["increment"] = new("scripts/increment.js", source),
                ["fail"] = new("scripts/fail.js", "function main(workbook){ workbook.getActiveWorksheet().getRange('A1').setValue(999); throw new Error('自动执行失败测试'); }"),
                ["writeText"] = new("scripts/writeText.js", "function main(workbook){ workbook.getActiveWorksheet().getRange('C1').setValue('已执行'); }")
            };
            for (var i = 0; i < 24; i++) dictionary[$"extra{i:00}"] = new($"scripts/extra{i:00}.js", "function main(){}");
            var names = dictionary.Keys.ToArray();
            var modules = new ScriptModuleSet(dictionary, names, scriptApiVersion: 2,
                bindings: names.Select(name => new ScriptBinding(name, "actions")).ToArray());
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(modules)); Synchronize();
            var path = Path.Combine(directory, "automation.xlsq"); new TypedWorkbookStorage().Save(path, Document.Snapshot);
            await OpenPathAsync(path);
            Check(!Document.ScriptsEnabled && _automationActions.ItemCount == 0 && _automationEmpty.Text!.Contains("加载操作"), "Opening an unsigned action workbook requests explicit loading without running module initialization");
            await Click(_automationLoad!);
            Check(Document.ScriptsEnabled && _automationActions.ItemCount == 27 && _scriptWindow is null, "Load button discovers every action across entry modules without opening the script editor");
            Check(Document.GetCellInput(0, 0, 0) == "" && Document.GetCellInput(0, 0, 2) == "", "Discovery does not execute action bodies");
            Check(!_automationRun!.IsEnabled && !_automationEmpty.IsVisible, "Execute requires an explicit selection after loading");
            Check(_automationActions.ItemTemplate!.Build(null) is null, "Dropdown template safely accepts a null item");
            var increment = _availableScriptActions.Single(a => a.Name == "increment");
            var revision = Document.Snapshot.Revision;
            await Search("  INCREMENT  ");
            Check(_automationActions.ItemCount == 1 && Equals(_automationActions.Items[0], increment) && !_automationRun.IsEnabled, "Name keyword search trims surrounding whitespace and requires explicit selection");
            await OpenDropdown(); _automationActions.SelectedItem = increment; await Layout();
            Check(_automationRun.IsEnabled && Document.Snapshot.Revision == revision && Document.GetCellInput(0, 0, 0) == "", "Selecting an action enables Execute without changing cells");
            Check(Item(increment.Name).Children.OfType<TextBlock>().Select(t => t.Text).SequenceEqual(new[] { increment.Name, "无描述" }), "Action rows display the module name and an empty-description placeholder");
            DoubleClick(Item(increment.Name)); await _operation; await Layout();
            Check(Document.Snapshot.Revision == revision && Document.GetCellInput(0, 0, 0) == "", "Double-clicking a dropdown item never executes an action");
            _automationActions.IsDropDownOpen = false;
            await Search("CREMENT");
            Check(_automationActions.ItemCount == 1 && Equals(_automationActions.SelectedItem, increment), "Name keywords match case-insensitively and preserve a matching selection");
            await Search("不存在的操作");
            Check(_automationActions.ItemCount == 0 && _automationActions.SelectedItem is null && _automationEmpty.IsVisible && _automationEmpty.Text!.Contains("没有匹配") && !_automationRun.IsEnabled,
                "No matches show an empty state and clear the hidden selection");
            await Search("");
            Check(_automationActions.ItemCount == 27 && _automationActions.SelectedItem is null && !_automationRun.IsEnabled, "Clearing search restores all actions without restoring a hidden selection");
            Check(Document.Snapshot.Revision == revision, "Searching and clearing keywords do not change the workbook");
            _automationActions.SelectedItem = increment; await Layout();
            using (var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(Bounds.Width), (int)Math.Ceiling(Bounds.Height)), new Vector(96, 96)))
            { bitmap.Render(this); bitmap.Save(Path.Combine(directory, "automation-ribbon.png"), PngBitmapEncoderOptions.Default); }
            await Click(_automationRun);
            Check(Document.GetCellValue(0, 0, 0) == "1" && Document.GetCellValue(0, 0, 1) == "10" && Document.Snapshot.Revision == revision + 1, "Execute button runs the selected real worker action as one transaction");
            await History(false);
            Check(Document.GetCellInput(0, 0, 0) == "" && Document.GetCellInput(0, 0, 1) == "", "One undo restores every cell changed by the action");
            await History(true);
            Check(Document.GetCellValue(0, 0, 0) == "1" && Document.GetCellValue(0, 0, 1) == "10", "Redo restores recorded action results");
            _automationRun.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            _automationRun.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Check(_automationRun.IsEffectivelyEnabled, "Execute keeps its enabled appearance while duplicate submissions are guarded");
            await _operation; await Layout();
            Check(Document.GetCellValue(0, 0, 0) == "2", "Repeated execution requests during a run do not submit a second action");
            var textAction = _availableScriptActions.Single(a => a.Name == "writeText");
            await Search("writeText"); await OpenDropdown(); var textItem = Item(textAction.Name);
            Check(ToolTip.GetTip(textItem)?.ToString() == textAction.Name + "\n无描述", "Action tooltips remain available for modules without descriptions");
            revision = Document.Snapshot.Revision;
            DoubleClick(textItem); await _operation; await Layout();
            Check(Document.Snapshot.Revision == revision && Document.GetCellInput(0, 0, 2) == "", "Double-clicking a search result does not execute it");
            _automationActions.SelectedItem = textAction; _automationActions.IsDropDownOpen = false; await Click(_automationRun);
            Check(Document.GetCellValue(0, 0, 2) == "已执行" && Document.GetCellValue(0, 0, 0) == "2", "Execute runs the explicitly selected search result");
            revision = Document.Snapshot.Revision; DoubleClick(_automationActions); await Layout();
            Check(Document.Snapshot.Revision == revision, "Double-clicking the closed dropdown never executes the selection");
            await Search(""); await OpenDropdown();
            var last = _availableScriptActions.Last(); _automationActions.ScrollIntoView(last); await Layout();
            Check(Item(last.Name).Children.OfType<TextBlock>().Last().Text == "无描述", "All actions remain reachable by scrolling, including those without descriptions");
            using (var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(Dropdown().Bounds.Width), (int)Math.Ceiling(Dropdown().Bounds.Height)), new Vector(96, 96)))
            { bitmap.Render(Dropdown()); bitmap.Save(Path.Combine(directory, "automation-dropdown.png"), PngBitmapEncoderOptions.Default); }
            _automationActions.IsDropDownOpen = false;
            Width = 760; await Layout();
            Check(Math.Abs(_ribbon.Bounds.Height - ribbonHeight) < .1 && _automationActions.Bounds.Height <= 50.5, "The search and dropdown keep the ribbon height fixed at the minimum window width");
            foreach (var control in new Control[] { _automationSearch, _automationActions })
                for (var ancestor = control.GetVisualParent(); ancestor is not null && ancestor != _ribbon; ancestor = ancestor.GetVisualParent())
                {
                    var origin = control.TranslatePoint(default, ancestor)!.Value;
                    Check(origin.Y >= -.1 && origin.Y + control.Bounds.Height <= ancestor.Bounds.Height + .1, $"{control.Name} fits vertically inside {ancestor.GetType().Name}");
                }
            Width = 1280;
            await Idle(); revision = Document.Snapshot.Revision;
            _automationSearch.Focus(); await Search("xincrement"); _automationSearch.CaretIndex = 0;
            _automationSearch.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Delete }); await _operation; await Layout();
            Check(_automationSearch.Text == "increment" && Document.Snapshot.Revision == revision && Document.GetCellValue(0, 0, 0) == "2", "Delete edits the search keyword without clearing worksheet cells");
            _automationSearch.SelectAll();
            _automationSearch.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.X, KeyModifiers = KeyModifiers.Control }); await _operation; await Layout();
            Check(_automationSearch.Text == "" && Document.Snapshot.Revision == revision && _clipboardCells?.IsCut != true, "Cut edits the search field without cutting worksheet cells");
            await OpenDropdown();
            _automationActions.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Escape }); await Layout();
            Check(!_automationActions.IsDropDownOpen && Document.Snapshot.Revision == revision, "Escape closes the dropdown without executing");
            _automationActions.Focus();
            _automationActions.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Down, KeyModifiers = KeyModifiers.Alt }); await Layout();
            Check(_automationActions.IsDropDownOpen && _filterPopup?.IsOpen != true && !_working, $"Alt+Down opens the action dropdown without invoking the worksheet filter (open={_automationActions.IsDropDownOpen}, focus={FocusManager!.GetFocusedElement()?.GetType().Name}, busy={_workspace.IsBusy})");
            Check(FocusManager!.GetFocusedElement() is ComboBoxItem or ComboBox, $"Dropdown keyboard focus stays on its controls ({FocusManager.GetFocusedElement()?.GetType().Name})");
            ((Control)FocusManager!.GetFocusedElement()!).RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Down }); await Layout();
            ((Control)FocusManager.GetFocusedElement()!).RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Enter }); await _operation; await Layout();
            Check(!_automationActions.IsDropDownOpen && Document.Snapshot.Revision == revision && !_workspace.IsBusy && !_working,
                $"Keyboard selection confirms an action without running it (busy={_workspace.IsBusy}, working={_working}, focus={FocusManager.GetFocusedElement()?.GetType().Name})");
            _automationActions.SelectedItem = _availableScriptActions.Single(a => a.Name == "fail");
            revision = Document.Snapshot.Revision;
            await PerformAsync(() => RunScriptActionAsync(_automationActions.SelectedItem as ScriptActionDescriptor));
            Check(Document.Snapshot.Revision == revision && Document.GetCellValue(0, 0, 0) == "2" && _status.Text!.Contains("自动执行失败测试"), "Worker errors are shown without committing partial edits");
            _grid.Select(new(0, 0)); _formula.Focus(); _formula.Text = "8";
            _automationActions.SelectedItem = increment; await Click(_automationRun);
            Check(Document.GetCellValue(0, 0, 0) == "9", "Execute commits the active cell draft before capturing the action snapshot");
            await Search("increment"); await OpenDropdown();
            DisableScripts();
            await Layout();
            Check(_automationActions.ItemCount == 0 && !_automationRun.IsEnabled && string.IsNullOrEmpty(_automationSearch.Text) && !_automationActions.IsDropDownOpen,
                "Disabling scripts clears search, closes the dropdown, and invalidates cached actions");
            await Click(_automationLoad!);
            _workspace.Editor.Execute(new UpdateScriptSourceCommand(source + "\n// updated")); Synchronize();
            Check(!Document.ScriptsEnabled && _automationActions.ItemCount == 0 && !_automationRun.IsEnabled, "Editing script source invalidates the displayed action list");
            revision = Document.Snapshot.Revision;
            try { await PerformAsync(() => RunScriptActionAsync(textAction)); } catch (InvalidOperationException) { } // PerformAsync reports the expected rejection in the status bar.
            await Layout();
            Check(Document.Snapshot.Revision == revision && _status.Text!.Contains("请先加载"), "A stale action cannot execute after source changes");
            await Click(_automationLoad!);
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(Document.Snapshot.ScriptModules.Modules, ["increment"], scriptApiVersion: 2)));
            Synchronize();
            Check(_automationActions.ItemCount == 0 && !_automationRun.IsEnabled, "Changing action bindings clears previously discovered actions");
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", "// no registrations") }, ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")])));
            Synchronize(); await Click(_automationLoad!);
            Check(_automationActions.ItemCount == 0 && _automationEmpty.Text!.Contains("没有可用") && !_automationRun.IsEnabled, "An action entry with no registrations displays an empty state after discovery");
            await OpenPathAsync(path); await Click(_automationLoad!); _workspace.New();
            Check(_automationActions.ItemCount == 0 && !_automationRun.IsEnabled && !Document.ScriptsEnabled, "Replacing the workbook clears actions from the previous session");
            await CheckEventFilesAsync(directory, Check);
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
