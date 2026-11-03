using System.Text.Json;
using Avalonia;
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
        StackPanel Item(string name) => _automationActions.GetVisualDescendants().OfType<StackPanel>()
            .Single(p => p.Children.OfType<TextBlock>().FirstOrDefault()?.Text == name);
        try
        {
            _workspace.New(); ShowBook();
            var tab = _ribbon.Items.Cast<TabItem>().Single(t => Equals(t.Header, "自动执行"));
            _ribbon.SelectedItem = tab; await Layout(); var ribbonHeight = _ribbon.Bounds.Height;
            Check(_automationEmpty.IsVisible && !_automationRun!.IsEnabled, "Empty workbook explains the empty action list and disables Execute");
            const string source = """
                quickExcel.actions.register({name:'增加计数',description:'将 A1 增加 1，并将 B1 增加 10；全部修改可一次撤销。'}, ({snapshot,patch}) => {
                    const id=snapshot.sheets[0].id;
                    patch.setValue(id,'A1',(snapshot.getCell(id,'A1').value || 0)+1);
                    patch.setValue(id,'B1',(snapshot.getCell(id,'B1').value || 0)+10);
                });
                quickExcel.actions.register('失败操作', ({snapshot,patch}) => {
                    patch.setValue(snapshot.sheets[0].id,'A1',999); throw new Error('自动执行失败测试');
                });
                """;
            var description = string.Concat(Enumerable.Repeat("较长的操作描述，悬停可查看完整内容。", 20));
            var more = "quickExcel.actions.register({name:'写入文本',description:" + JsonSerializer.Serialize(description) + "}, ({snapshot,patch}) => { patch.setValue(snapshot.sheets[0].id,'C1','已执行'); });\n" +
                string.Join("\n", Enumerable.Range(0, 24).Select(i => $"quickExcel.actions.register('附加操作{i:00}', () => {{ }});"));
            var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
            {
                ["main"] = new("scripts/main.js", source), ["extra"] = new("scripts/extra.js", more)
            }, ["main", "extra"], scriptApiVersion: 2, bindings: [new("main", "actions"), new("extra", "actions")]);
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(modules)); Synchronize();
            var path = Path.Combine(directory, "automation.xlsq"); new TypedWorkbookStorage().Save(path, Document.Snapshot);
            await OpenPathAsync(path);
            Check(!Document.ScriptsEnabled && _automationActions.ItemCount == 0 && _automationEmpty.Text!.Contains("加载操作"), "Opening an unsigned action workbook requests explicit loading without running module initialization");
            await Click(_automationLoad!);
            Check(Document.ScriptsEnabled && _automationActions.ItemCount == 27 && _scriptWindow is null, "Load button discovers every action across entry modules without opening the script editor");
            Check(Document.GetCellInput(0, 0, 0) == "" && Document.GetCellInput(0, 0, 2) == "", "Discovery does not execute action bodies");
            Check(!_automationRun!.IsEnabled && !_automationEmpty.IsVisible, "Execute requires an explicit selection after loading");
            Check(_automationActions.ItemTemplate!.Build(null) is null, "Virtualized template safely accepts a null item");
            var increment = _availableScriptActions.Single(a => a.Name == "增加计数");
            _automationActions.SelectedItem = increment; _automationActions.ScrollIntoView(increment); await Layout();
            var revision = Document.Snapshot.Revision;
            Check(_automationRun.IsEnabled && Document.Snapshot.Revision == revision && Document.GetCellInput(0, 0, 0) == "", "Selecting an action enables Execute without changing cells");
            Check(Item(increment.Name).Children.OfType<TextBlock>().Select(t => t.Text).SequenceEqual(new[] { increment.Name, increment.Description }), "Action rows display both the registered name and description");
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
            Check(!_automationRun.IsEffectivelyEnabled, "Execute is disabled while an action is running");
            await _operation; await Layout();
            Check(Document.GetCellValue(0, 0, 0) == "2", "Repeated execution requests during a run do not submit a second action");
            var textAction = _availableScriptActions.Single(a => a.Name == "写入文本");
            _automationActions.ScrollIntoView(textAction); await Layout(); var textItem = Item(textAction.Name);
            Check(ToolTip.GetTip(textItem)?.ToString() == textAction.Name + "\n" + description, "Long descriptions remain available in full through the tooltip");
            DoubleClick(textItem); await _operation; await Layout();
            Check(Document.GetCellValue(0, 0, 2) == "已执行" && Document.GetCellValue(0, 0, 0) == "2" && Equals(_automationActions.SelectedItem, textAction), "Double-click executes the clicked item rather than the previous selection");
            revision = Document.Snapshot.Revision; DoubleClick(_automationActions); await Layout();
            Check(Document.Snapshot.Revision == revision, "Double-clicking list background never executes the previous selection");
            var last = _availableScriptActions.Last(); _automationActions.ScrollIntoView(last); await Layout();
            Check(Item(last.Name).Children.OfType<TextBlock>().Last().Text == "无描述", "All actions remain reachable by scrolling, including those without descriptions");
            Width = 760; await Layout();
            Check(Math.Abs(_ribbon.Bounds.Height - ribbonHeight) < .1 && _automationActions.Bounds.Height <= 90.5, "The action list keeps the ribbon height fixed at the minimum window width");
            Width = 1280;
            _automationActions.SelectedItem = _availableScriptActions.Single(a => a.Name == "失败操作");
            revision = Document.Snapshot.Revision;
            await PerformAsync(() => RunScriptActionAsync(_automationActions.SelectedItem as ScriptActionDescriptor));
            Check(Document.Snapshot.Revision == revision && Document.GetCellValue(0, 0, 0) == "2" && _status.Text!.Contains("自动执行失败测试"), "Worker errors are shown without committing partial edits");
            _grid.Select(new(0, 0)); _formula.Focus(); _formula.Text = "8";
            _automationActions.SelectedItem = increment; await Click(_automationRun);
            Check(Document.GetCellValue(0, 0, 0) == "9", "Execute commits the active cell draft before capturing the action snapshot");
            DisableScripts();
            Check(_automationActions.ItemCount == 0 && !_automationRun.IsEnabled, "Disabling scripts immediately clears cached actions and execution state");
            await Click(_automationLoad!);
            _workspace.Editor.Execute(new UpdateScriptSourceCommand(source + "\n// updated")); Synchronize();
            Check(!Document.ScriptsEnabled && _automationActions.ItemCount == 0 && !_automationRun.IsEnabled, "Editing script source invalidates the displayed action list");
            revision = Document.Snapshot.Revision; DoubleClick(textItem);
            try { await _operation; } catch (InvalidOperationException) { } // PerformAsync reports the expected rejection in the status bar.
            await Layout();
            Check(Document.Snapshot.Revision == revision && _status.Text!.Contains("请先加载"), "A stale realized action item cannot execute after source changes");
            await Click(_automationLoad!);
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(Document.Snapshot.ScriptModules.Modules, ["main", "extra"], scriptApiVersion: 2)));
            Synchronize();
            Check(_automationActions.ItemCount == 0 && !_automationRun.IsEnabled, "Changing action bindings clears previously discovered actions");
            _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", "// no registrations") }, ["main"], scriptApiVersion: 2, bindings: [new("main", "actions")])));
            Synchronize(); await Click(_automationLoad!);
            Check(_automationActions.ItemCount == 0 && _automationEmpty.Text!.Contains("没有可用") && !_automationRun.IsEnabled, "An action entry with no registrations displays an empty state after discovery");
            await OpenPathAsync(path); await Click(_automationLoad!); _workspace.New();
            Check(_automationActions.ItemCount == 0 && !_automationRun.IsEnabled && !Document.ScriptsEnabled, "Replacing the workbook clears actions from the previous session");
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
