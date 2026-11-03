using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckScriptConsoleAsync(Action<bool, string> check, string directory)
    {
        static async Task Layout() => await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        async Task Capture(string name)
        {
            await Layout();
            using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(_scriptOutputTabs!.Bounds.Width), (int)Math.Ceiling(_scriptOutputTabs.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(_scriptOutputTabs); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        _workspace.New();
        _workspace.Editor.Execute(new UpdateScriptSourceCommand("throw Error('Stored module must not execute in console');"));
        await OpenScriptsAsync(); var modules = Document.Snapshot.ScriptModules;
        _consoleInput!.Text = "console.log('console 中文', {value:7}); console.warn('缺少可选字段', {cell:'A1'}); console.error('输入格式错误'); console.info('信息'); console.debug('调试'); 6*7";
        var presenter = _consoleInput.GetVisualDescendants().OfType<TextPresenter>().Single();
        presenter.PreeditText = "zhong"; await RunConsoleAsync();
        check(_consoleInput.IsComposing && !_consoleRun!.IsEnabled && Document.Session.ConsoleEntries.IsEmpty,
            "Synthetic console preedit blocks execution without committing incomplete input");
        presenter.PreeditText = null;
        await RunConsoleAsync();
        check(ConsoleVisibleText.Contains("console 中文 {value: 7}") && ConsoleVisibleText.Contains("=> 42"),
            "Native console displays object logs and immediate expression results");
        _consoleLevel!.SelectedIndex = 2; await Layout();
        var warningRow = _consoleOutput!.GetVisualDescendants().OfType<Border>().Single(b => b.Tag is ScriptConsoleEntry { Level: ScriptConsoleLevel.Warning });
        var warningText = warningRow.GetVisualDescendants().OfType<SelectableTextBlock>().Single(); warningText.SelectAll();
        check(_consoleVisibleEntries.Length == 1 && warningRow.Background is ISolidColorBrush { Color: var warningColor } && warningColor == Color.Parse("#FFFBE6") &&
            warningText.Foreground is ISolidColorBrush { Color: var warningInk } && warningInk == Color.Parse("#7A4F00") && warningText.SelectedText == "缺少可选字段 {cell: A1}",
            "Warning level filters actual rows, uses amber styling and keeps object output selectable");
        _consoleLevel.SelectedIndex = 3; await Layout();
        var errorRow = _consoleOutput!.GetVisualDescendants().OfType<Border>().Single(b => b.Tag is ScriptConsoleEntry { Level: ScriptConsoleLevel.Error });
        check(_consoleVisibleEntries.Length == 1 && errorRow.Background is ISolidColorBrush { Color: var errorColor } && errorColor == Color.Parse("#FCE8E6") &&
            errorRow.GetVisualDescendants().OfType<SelectableTextBlock>().Single().Foreground is ISolidColorBrush { Color: var errorInk } && errorInk == Color.Parse("#B3261E"),
            "Error level filters actual rows and uses red foreground with pale red background");
        check(_consoleCount!.Text!.Contains("1 错误  1 警告"), "Console toolbar counts errors and warnings without counting source text");
        _consoleFilter!.Text = "missing"; await Layout();
        check(_consoleVisibleEntries.Length == 0 && _consoleEmpty!.IsVisible && !_consoleCopy!.IsEnabled, "Empty filters show a clear empty state and disable copying");
        _consoleLevel.SelectedIndex = 0; _consoleFilter.Text = "a1"; await Layout();
        check(_consoleVisibleEntries.Any(e => e.Level == ScriptConsoleLevel.Warning) && !_consoleVisibleEntries.Any(e => e.Level == ScriptConsoleLevel.Error),
            "Console text filtering ignores case and the copy buffer follows visible entries");
        var clipboard = _clipboard; var memoryClipboard = new MemoryClipboard(); _clipboard = memoryClipboard;
        try
        {
            _consoleCopy!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await Layout();
            check(memoryClipboard.Value.Text == ConsoleVisibleText, "Copy toolbar command writes only filtered output through the clipboard adapter");
        }
        finally { _clipboard = clipboard; }
        _consoleFilter.Text = ""; _consoleInput.Text = "snapshot.sheets";
        await Capture("console-styled");
        _scriptWindow!.Width = 760; await Layout(); await Capture("console-styled-narrow");
        check(_consoleFilter.Bounds.Width > 80 && _consoleRun!.Bounds.Width > 40 && _consoleCount.Bounds.Right <= ((Control)_consoleCount.Parent!).Bounds.Width,
            "Console filters and execution controls remain usable at the minimum window width");
        _scriptWindow.Width = 1080;
        check(!Document.ScriptsEnabled && ReferenceEquals(modules, Document.Snapshot.ScriptModules),
            "Console runs only explicit input without authorizing or executing stored modules");
        _consoleInput.Text = "patch.setValue(snapshot.sheets[0].id,'A1',42);";
        await RunConsoleAsync();
        check(Document.GetCellValue(0, 0, 0) == "42" && Document.Undo() && Document.GetCellValue(0, 0, 0) == "",
            "Native console submits one undoable workbook transaction");
        _consoleInput.Text = "console.log('visible before error'); patch.setValue(snapshot.sheets[0].id,'B1',99); throw Error('console failure');";
        await RunConsoleAsync();
        check(ConsoleVisibleText.Contains("visible before error") && ConsoleVisibleText.Contains("console failure") && Document.GetCellValue(0, 0, 1) == "" &&
            Document.Session.ConsoleEntries[^1].Level == ScriptConsoleLevel.Error,
            "Console keeps printed output on error and discards partial writes");
        var last = _consoleInput.Text; _consoleInput.Text = "unfinished input";
        NavigateConsoleHistory(-1); check(_consoleInput.Text == last, "Console history recalls the previous command");
        NavigateConsoleHistory(1); check(_consoleInput.Text == "unfinished input", "Console history restores the current input draft");
        _consoleInput.Text = "patch.setValue(snapshot.sheets[0].id,'C1',9); while(true){}";
        var running = RunConsoleAsync();
        for (var retry = 0; retry < 100 && _scriptCancellation is null && !running.IsCompleted; retry++) await Task.Delay(10);
        _scriptCancellation?.Cancel(); await running;
        check(Document.GetCellValue(0, 0, 2) == "" && _consoleRun!.IsEnabled, "Stopping console execution discards writes and re-enables execution");
        var revision = Document.Snapshot.Revision;
        _scriptWindow!.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == "清空输出").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        check(ConsoleVisibleText == "" && _consoleOutput!.ItemCount == 0 && _consoleEmpty!.IsVisible && Document.Snapshot.Revision == revision, "Clear output changes no workbook content");
        await CloseScriptWindowAsync();
        _workspace.New(); await OpenScriptsAsync();
        check(ConsoleVisibleText == "" && _consoleOutput!.ItemCount == 0 && _consoleHistory.Count == 0, "Console output and command history do not leak into another workbook session");
        await CloseScriptWindowAsync();
    }
}
