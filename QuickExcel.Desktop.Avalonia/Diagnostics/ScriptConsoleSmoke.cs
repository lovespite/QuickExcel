using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Input;
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
        void Press(Key key, KeyModifiers modifiers = KeyModifiers.None) => _consoleInput!.RaiseEvent(new KeyEventArgs
            { RoutedEvent = KeyDownEvent, Key = key, KeyModifiers = modifiers });
        async Task CompleteCommand()
        {
            await _operation;
            // PerformAsync also asynchronously releases the editor hold after its command task finishes.
            for (var retry = 0; retry < 100 && _working; retry++) await Task.Delay(10);
            if (_working) throw new TimeoutException("Console command did not release its editor hold.");
            await Layout();
        }
        async Task Capture(string name)
        {
            await Layout();
            using var bitmap = new RenderTargetBitmap(new PixelSize((int)Math.Ceiling(_scriptOutputTabs!.Bounds.Width), (int)Math.Ceiling(_scriptOutputTabs.Bounds.Height)), new Vector(96, 96));
            bitmap.Render(_scriptOutputTabs); bitmap.Save(Path.Combine(directory, name + ".png"), PngBitmapEncoderOptions.Default);
        }
        _workspace.New();
        _workspace.Editor.Execute(new UpdateScriptSourceCommand("throw Error('Stored module must not execute in console');"));
        await OpenScriptsAsync(); var modules = Document.Snapshot.ScriptModules;
        await Layout();
        check(!_scriptOutputTabs!.IsVisible && !_scriptOutputSplitter!.IsVisible && _scriptEditorPane!.RowDefinitions[2].ActualHeight == 0 &&
            _scriptContent!.Bounds.Height == _scriptEditorPane.Bounds.Height,
            "Script workspace starts with the console fully collapsed and the editor filling the right pane");
        var navigation = (Control)_modulePicker!.Parent!;
        check(navigation.Bounds.Right <= _scriptEditorPane!.Bounds.X && navigation.Bounds.Height == _scriptEditorPane.Bounds.Height,
            "File navigation stays to the left of the editor and spans the workspace height");
        var fileSplitter = ((Panel)navigation.Parent!).Children.OfType<GridSplitter>().Single();
        var navigationWidth = navigation.Bounds.Width;
        fileSplitter.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Right }); await Layout();
        check(navigation.Bounds.Width > navigationWidth, "The file/editor splitter resizes the navigation through its actual keyboard handler");
        fileSplitter.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Left }); await Layout();
        _modulePicker.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.F12 }); await Layout();
        check(_scriptOutputTabs.IsVisible && _scriptOutputSplitter!.IsVisible && _consoleInput!.IsKeyboardFocusWithin &&
            _scriptContent!.Bounds.Bottom <= _scriptOutputTabs.Bounds.Top,
            "F12 from the native file list expands the console below the editor and focuses input");
        var outputHeight = _scriptOutputTabs.Bounds.Height;
        _scriptOutputSplitter!.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Up }); await Layout();
        check(_scriptOutputTabs.Bounds.Height > outputHeight, "The editor/console splitter resizes the output through its actual keyboard handler");
        _consoleInput!.Text = "retained draft";
        _scriptEditorPane.RowDefinitions[0].Height = new(2, GridUnitType.Star);
        _scriptEditorPane.RowDefinitions[2].Height = new(1, GridUnitType.Star);
        Press(Key.F12); await Layout();
        check(!_scriptOutputTabs.IsVisible && _consoleInput.Text == "retained draft", "F12 collapses the console without losing input");
        async Task BrowserF12(bool composing = false, bool repeat = false, bool control = false)
        {
            await _scriptPage!.Browser.RequestAsync("test", "testResult", new { action = "consoleShortcut", composing, repeat, control });
            await Layout();
        }
        await BrowserF12(composing: true); await BrowserF12(repeat: true); await BrowserF12(control: true);
        check(!_scriptOutputTabs.IsVisible, "Editor F12 ignores composition, auto-repeat and modified shortcuts");
        await BrowserF12();
        check(_scriptOutputTabs.IsVisible && _consoleInput.IsKeyboardFocusWithin && _scriptEditorPane.RowDefinitions[0].Height.Value == 2 &&
            _scriptEditorPane.RowDefinitions[2].Height.Value == 1 && _consoleInput.Text == "retained draft",
            "F12 from the actual WebView keyboard bridge opens the console and preserves the resized split and draft");
        var revisionBeforeToggle = Document.Snapshot.Revision;
        _scriptConsoleToggle!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await Layout();
        await OpenScriptConsoleAsync(); await Layout();
        check(_scriptOutputTabs.IsVisible && _consoleInput.IsKeyboardFocusWithin && Document.Snapshot.Revision == revisionBeforeToggle,
            "Toolbar collapse and the console menu reopen preserve document history and focus input");
        _scriptEditorPane.RowDefinitions[0].Height = new(3, GridUnitType.Star);
        _scriptEditorPane.RowDefinitions[2].Height = new(2, GridUnitType.Star);
        _consoleInput!.Focus(); _consoleInput.Text = "";
        Press(Key.Enter); await CompleteCommand();
        check(_consoleInput.Text == "" && Document.Session.ConsoleEntries.IsEmpty && _consoleHistory.Count == 0,
            "Enter on empty input neither inserts a newline nor submits a command");
        _consoleInput.Text = "6 * 7"; SetConsoleCaret(_consoleInput, 1); Press(Key.Enter, KeyModifiers.Control);
        check(_consoleInput.Text == "6" + _consoleInput.NewLine + " * 7" && _consoleInput.CaretIndex == 1 + _consoleInput.NewLine.Length && Document.Session.ConsoleEntries.IsEmpty,
            "Ctrl+Enter inserts one newline at the caret through the actual TextBox key route without execution");
        _consoleInput.Text = "6xx * 7"; _consoleInput.SelectionStart = 1; _consoleInput.SelectionEnd = 3;
        Press(Key.Enter, KeyModifiers.Control);
        check(_consoleInput.Text == "6" + _consoleInput.NewLine + " * 7" && _consoleInput.SelectionStart == _consoleInput.SelectionEnd && Document.Session.ConsoleEntries.IsEmpty,
            "Ctrl+Enter replaces selected text with one newline");
        const string firstCommand = "console.log('console 中文', {value:7}); console.warn('缺少可选字段', {cell:'A1'}); console.error('输入格式错误'); console.info('信息'); console.debug('调试'); 6*7";
        _consoleInput.Text = firstCommand;
        var presenter = _consoleInput.GetVisualDescendants().OfType<TextPresenter>().Single();
        presenter.PreeditText = "zhong"; Press(Key.Enter); await CompleteCommand(); Press(Key.Up); Press(Key.Down); Press(Key.F12);
        check(_consoleInput.IsComposing && !_consoleRun!.IsEnabled && Document.Session.ConsoleEntries.IsEmpty && _scriptOutputTabs.IsVisible,
            "Synthetic console preedit blocks keyboard submission and history without executing incomplete input");
        presenter.PreeditText = null; _consoleInput.Text = firstCommand;
        Press(Key.Enter); await CompleteCommand();
        check(_consoleInput.Text == "" && _consoleInput.IsKeyboardFocusWithin && _consoleHistory[^1] == firstCommand,
            "Enter submits through the TextBox key route, clears accepted input and keeps focus with history intact");
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
        _consoleFilter.Text = ""; _consoleInput.Text = "workbook.getWorksheets";
        await Capture("console-styled");
        // A wide navigation split must still leave room for the editor when the window shrinks.
        var workspace = (global::Avalonia.Controls.Grid)navigation.Parent!;
        workspace.ColumnDefinitions[0].Width = workspace.ColumnDefinitions[2].Width = new(1, GridUnitType.Star);
        _scriptWindow!.Width = 760; _scriptWindow.Height = 500; await Layout(); await Capture("console-styled-narrow");
        check(_consoleFilter.Bounds.Width > 80 && _consoleRun!.Bounds.Width > 40 && _consoleCount.Bounds.Right <= ((Control)_consoleCount.Parent!).Bounds.Width,
            "Console filters and execution controls remain usable at the minimum window width");
        await File.WriteAllTextAsync(Path.Combine(directory, "script-layout-narrow.json"), System.Text.Json.JsonSerializer.Serialize(new {
            editor = _scriptContent!.Bounds.ToString(), output = _scriptOutputTabs.Bounds.ToString(), pane = _scriptEditorPane.Bounds.ToString(),
            navigation = navigation.Bounds.ToString(), window = _scriptWindow.ClientSize.ToString() }));
        check(navigation.Bounds.Width >= 149 && _scriptEditorPane.Bounds.Width >= 419 && _scriptEditorPane.Bounds.Right <= workspace.Bounds.Width + 1 &&
            _scriptContent.Bounds.Height >= 99 && _scriptOutputTabs.Bounds.Height >= 189 &&
            _scriptOutputTabs.Bounds.Bottom <= _scriptEditorPane.Bounds.Height + 1,
            "Widened file navigation, editor and console still fit when shrinking to the minimum window width and height");
        workspace.ColumnDefinitions[0].Width = new(200, GridUnitType.Star); workspace.ColumnDefinitions[2].Width = new(875, GridUnitType.Star);
        _scriptWindow.Width = 1080; _scriptWindow.Height = 760;
        check(!Document.ScriptsEnabled && ReferenceEquals(modules, Document.Snapshot.ScriptModules),
            "Console runs only explicit input without authorizing or executing stored modules");
        _consoleInput.Text = "workbook.getActiveWorksheet().getCell('A1').setValue(42);";
        _consoleRun!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await CompleteCommand();
        check(Document.GetCellValue(0, 0, 0) == "42" && Document.Undo() && Document.GetCellValue(0, 0, 0) == "",
            "Native console submits one undoable workbook transaction");
        check(_consoleInput.Text == "", "Execute button clears accepted input just like Enter");
        _consoleInput.Text = "console.log('visible before error'); workbook.getActiveWorksheet().getCell('B1').setValue(99); throw Error('console failure');";
        Press(Key.Enter); await CompleteCommand();
        check(ConsoleVisibleText.Contains("visible before error") && ConsoleVisibleText.Contains("console failure") && Document.GetCellValue(0, 0, 1) == "" &&
            Document.Session.ConsoleEntries[^1].Level == ScriptConsoleLevel.Error,
            "Console keeps printed output on error and discards partial writes");
        check(_consoleInput.Text == "", "Failed commands also clear accepted input and remain available in history");
        var last = _consoleHistory[^1]; var previous = _consoleHistory[^2]; const string draft = "unfinished\ninput";
        _consoleInput.Text = draft; SetConsoleCaret(_consoleInput, 4);
        Press(Key.Up);
        check(_consoleInput.Text == draft && _consoleInput.CaretIndex == 0 && _consoleHistoryIndex == _consoleHistory.Count,
            "Up inside multiline input first moves to its start without replacing the draft");
        Press(Key.Up); check(_consoleInput.Text == last && _consoleInput.CaretIndex == 0, "Up at the start recalls the previous command");
        Press(Key.Up); check(_consoleInput.Text == previous && _consoleInput.CaretIndex == 0, "Repeated Up traverses older commands without extra cursor movement");
        Press(Key.Down);
        check(_consoleInput.Text == previous && _consoleInput.CaretIndex == previous.Length, "Down first moves to the end before changing history direction");
        Press(Key.Down); check(_consoleInput.Text == last && _consoleInput.CaretIndex == last.Length, "Down at the end recalls a newer command");
        Press(Key.Down); check(_consoleInput.Text == draft && _consoleInput.CaretIndex == draft.Length, "Down past the newest command restores the current input draft");
        Press(Key.Down); check(_consoleInput.Text == draft, "History navigation clamps at the newest draft");
        SetConsoleCaret(_consoleInput, 0); _consoleInput.SelectionEnd = 4;
        Press(Key.Up);
        check(_consoleInput.Text == draft && _consoleInput.CaretIndex == 0 && _consoleInput.SelectionStart == _consoleInput.SelectionEnd,
            "Up collapses a selection to the start before recalling history");
        _consoleInput.Text = "workbook.getActiveWorksheet().getCell('C1').setValue(9); while(true){}";
        var running = RunConsoleAsync();
        for (var retry = 0; retry < 100 && _scriptCancellation is null && !running.IsCompleted; retry++) await Task.Delay(10);
        check(_scriptCancellation is not null && _consoleInput.Text == "", "Input clears when submission is accepted, before worker completion");
        var historyCount = _consoleHistory.Count; _consoleInput.Text = "next draft";
        Press(Key.Enter); await CompleteCommand();
        check(_consoleInput.Text == "next draft" && _consoleHistory.Count == historyCount, "Busy submission keeps the next draft without adding a duplicate command");
        _scriptCancellation?.Cancel(); await running;
        check(Document.GetCellValue(0, 0, 2) == "" && _consoleRun!.IsEnabled, "Stopping console execution discards writes and re-enables execution");
        check(_consoleInput.Text == "next draft", "Completion or cancellation never clears text entered during execution");
        var revision = Document.Snapshot.Revision;
        _scriptWindow!.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == "清空输出").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        check(ConsoleVisibleText == "" && _consoleOutput!.ItemCount == 0 && _consoleEmpty!.IsVisible && Document.Snapshot.Revision == revision, "Clear output changes no workbook content");
        await CloseScriptWindowAsync();
        _workspace.New(); await OpenScriptsAsync();
        check(ConsoleVisibleText == "" && _consoleOutput!.ItemCount == 0 && _consoleHistory.Count == 0 && !_scriptOutputTabs!.IsVisible,
            "Reopened workspace defaults to a collapsed console without leaking output or history into another workbook session");
        await CloseScriptWindowAsync();
    }
}
