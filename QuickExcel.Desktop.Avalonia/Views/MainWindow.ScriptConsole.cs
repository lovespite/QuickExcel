using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Grid.Avalonia;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private CompositionTextBox? _consoleInput;
    private ListBox? _consoleOutput;
    private TextBox? _consoleFilter;
    private ComboBox? _consoleLevel;
    private TextBlock? _consoleCount, _consoleEmpty;
    private Button? _consoleRun, _consoleStop, _consoleCopy;
    private ScriptConsoleEntry[] _consoleVisibleEntries = [];
    private TabControl? _scriptOutputTabs;
    private readonly List<string> _consoleHistory = [];
    private Guid _consoleSession;
    private int _consoleHistoryIndex;
    private string _consolePendingInput = "";
    private long _consoleLastSequence = -1;
    private static readonly IBrush ConsoleInk = Brush.Parse("#202124"), ConsoleMuted = Brush.Parse("#70757A"),
        ConsoleLine = Brush.Parse("#E0E0E0"), ConsoleBlue = Brush.Parse("#1967D2"),
        ConsoleWarningInk = Brush.Parse("#7A4F00"), ConsoleWarningBackground = Brush.Parse("#FFFBE6"),
        ConsoleErrorInk = Brush.Parse("#B3261E"), ConsoleErrorBackground = Brush.Parse("#FCE8E6");

    private string ConsoleVisibleText => string.Join(Environment.NewLine, _consoleVisibleEntries.Select(e =>
        $"{e.Time.ToLocalTime():HH:mm:ss} {(e.Kind == "input" ? "> " : "")}{e.Text}"));

    private async Task OpenScriptConsoleAsync()
    {
        await OpenScriptsAsync();
        if (_scriptOutputTabs is not null) _scriptOutputTabs.SelectedIndex = 0;
        _consoleInput?.Focus();
    }

    private Control CreateScriptOutput()
    {
        if (_consoleSession != Document.Session.SessionId)
        { _consoleSession = Document.Session.SessionId; _consoleHistory.Clear(); _consolePendingInput = ""; }
        _consoleHistoryIndex = _consoleHistory.Count; _consoleLastSequence = -1;
        var panel = new LayoutGrid { RowDefinitions = new("Auto,*,Auto,Auto"), Background = Brushes.White };
        var tools = new LayoutGrid { ColumnDefinitions = new("Auto,*,Auto,Auto,Auto"), Margin = new(6, 3) };
        var clear = ConsoleButton("⊘", "清空输出");
        clear.Click += (_, _) => { Document.Session.ClearConsole(); RefreshScriptConsole(true); }; tools.Children.Add(clear);
        _consoleFilter = new TextBox { Height = 26, MinHeight = 0, Padding = new(7, 3), Margin = new(6, 0), FontSize = 12,
            PlaceholderText = "过滤输出", BorderBrush = ConsoleLine, Background = Brushes.White, Foreground = ConsoleInk };
        AutomationProperties.SetName(_consoleFilter, "过滤控制台输出");
        _consoleFilter.TextChanged += (_, _) => RefreshScriptConsole(true);
        LayoutGrid.SetColumn(_consoleFilter, 1); tools.Children.Add(_consoleFilter);
        _consoleLevel = new ComboBox { ItemsSource = new[] { "所有级别", "日志 / 信息", "警告", "错误", "调试" },
            SelectedIndex = 0, Width = 110, Height = 26, MinHeight = 0, Padding = new(7, 2), FontSize = 12 };
        AutomationProperties.SetName(_consoleLevel, "控制台日志级别");
        _consoleLevel.SelectionChanged += (_, _) => RefreshScriptConsole(true);
        LayoutGrid.SetColumn(_consoleLevel, 2); tools.Children.Add(_consoleLevel);
        _consoleCopy = ConsoleButton("复制", "复制当前筛选的控制台输出");
        _consoleCopy.Click += async (_, _) => await PerformAsync(() => _clipboard.WriteAsync(new(ConsoleVisibleText)));
        LayoutGrid.SetColumn(_consoleCopy, 3); tools.Children.Add(_consoleCopy);
        _consoleCount = new TextBlock { FontSize = 11, Foreground = ConsoleMuted, VerticalAlignment = VerticalAlignment.Center, Margin = new(8, 0) };
        LayoutGrid.SetColumn(_consoleCount, 4); tools.Children.Add(_consoleCount);
        panel.Children.Add(new Border { Background = Brush.Parse("#F3F3F3"), BorderBrush = ConsoleLine, BorderThickness = new(0, 1, 0, 1), Child = tools });

        _consoleOutput = new ListBox { BorderThickness = new(0), Padding = new(0), Background = Brushes.White,
            ItemTemplate = new FuncDataTemplate<ScriptConsoleEntry>((entry, _) => entry is null ? null : CreateConsoleRow(entry)) };
        _consoleOutput.Styles.Add(new Style(x => x.OfType<ListBoxItem>()) { Setters = {
            new Setter(TemplatedControl.PaddingProperty, new Thickness(0)), new Setter(Layoutable.MarginProperty, new Thickness(0)),
            new Setter(Layoutable.MinHeightProperty, 0d), new Setter(ContentControl.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch) } });
        ScrollViewer.SetHorizontalScrollBarVisibility(_consoleOutput, ScrollBarVisibility.Disabled);
        AutomationProperties.SetName(_consoleOutput, "调试控制台输出，可选择复制");
        var outputArea = new LayoutGrid(); outputArea.Children.Add(_consoleOutput);
        _consoleEmpty = new TextBlock { Text = "控制台已就绪。输入表达式，或使用 console.log / warn / error 输出消息。", FontSize = 12,
            Foreground = ConsoleMuted, Margin = new(28, 12), IsHitTestVisible = false, TextWrapping = TextWrapping.Wrap };
        outputArea.Children.Add(_consoleEmpty); LayoutGrid.SetRow(outputArea, 1); panel.Children.Add(outputArea);

        var inputArea = new LayoutGrid { ColumnDefinitions = new("26,*") };
        inputArea.Children.Add(new TextBlock { Text = "›", FontFamily = new("Consolas"), FontSize = 20, Foreground = ConsoleBlue, Margin = new(8, 1) });
        _consoleInput = new CompositionTextBox { AcceptsReturn = true, Height = 54, MinHeight = 0, FontFamily = new("Consolas"), FontSize = 13,
            Padding = new(2, 5), BorderThickness = new(0), Background = Brushes.White, Foreground = ConsoleInk,
            PlaceholderText = "输入 JavaScript 表达式或代码…", MaxLength = DocumentSession.MaxConsoleSourceBytes };
        ToolTip.SetTip(_consoleInput, "可用 snapshot / patch / fs；每次执行的变量独立，工作簿修改可撤销。");
        AutomationProperties.SetName(_consoleInput, "调试控制台 JavaScript 输入");
        _consoleInput.CompositionChanged += () => RefreshScriptConsole();
        _consoleInput.KeyDown += async (_, e) =>
        {
            if (_consoleInput.IsComposing) return;
            if (e.Key == Key.Enter && (e.KeyModifiers & (KeyModifiers.Control | KeyModifiers.Meta)) != 0)
            { e.Handled = true; await PerformAsync(RunConsoleAsync); }
            else if ((e.KeyModifiers & KeyModifiers.Alt) != 0 && e.Key is Key.Up or Key.Down)
            { e.Handled = true; NavigateConsoleHistory(e.Key == Key.Up ? -1 : 1); }
        };
        LayoutGrid.SetColumn(_consoleInput, 1); inputArea.Children.Add(_consoleInput);
        var inputBorder = new Border { BorderBrush = ConsoleLine, BorderThickness = new(0, 1, 0, 0), Child = inputArea };
        LayoutGrid.SetRow(inputBorder, 2); panel.Children.Add(inputBorder);
        var footer = new DockPanel { Background = Brush.Parse("#F8F9FA"), LastChildFill = true };
        var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 2, Margin = new(4, 0) };
        _consoleRun = ConsoleButton("▶ 执行", "执行控制台代码（Ctrl+Enter）");
        _consoleRun.Foreground = ConsoleBlue; _consoleRun.Click += async (_, _) => await PerformAsync(RunConsoleAsync); actions.Children.Add(_consoleRun);
        _consoleStop = ConsoleButton("■ 停止", "停止执行（Esc）");
        _consoleStop.Click += (_, _) => _scriptCancellation?.Cancel(); actions.Children.Add(_consoleStop);
        DockPanel.SetDock(actions, Dock.Right); footer.Children.Add(actions);
        footer.Children.Add(new TextBlock { Text = "JavaScript  ·  Ctrl+Enter 执行  ·  Alt+↑↓ 历史", FontSize = 11, Foreground = ConsoleMuted,
            VerticalAlignment = VerticalAlignment.Center, Margin = new(9, 0) });
        LayoutGrid.SetRow(footer, 3); panel.Children.Add(footer);
        _scriptOutputTabs = new TabControl { Height = 290, Background = Brushes.White, Foreground = ConsoleInk, ItemsSource = new[] {
            new TabItem { Header = "控制台", FontSize = 13, MinHeight = 30, Padding = new(10, 6), Content = panel },
            new TabItem { Header = "运行诊断", FontSize = 13, MinHeight = 30, Padding = new(10, 6), Content = _scriptDiagnostics } } };
        RefreshScriptConsole(); return _scriptOutputTabs;
    }
    private static Button ConsoleButton(string text, string description)
    {
        var button = new Button { Content = text, Height = 26, MinHeight = 0, Padding = new(7, 2), FontSize = 12,
            Background = Brushes.Transparent, BorderThickness = new(0), Foreground = ConsoleInk };
        AutomationProperties.SetName(button, description); ToolTip.SetTip(button, description); return button;
    }
    private static Control CreateConsoleRow(ScriptConsoleEntry entry)
    {
        var (icon, ink, background, label) = entry.Level switch
        {
            ScriptConsoleLevel.Warning => ("⚠", ConsoleWarningInk, ConsoleWarningBackground, "警告"),
            ScriptConsoleLevel.Error => ("⊗", ConsoleErrorInk, ConsoleErrorBackground, "错误"),
            ScriptConsoleLevel.Input => ("›", ConsoleBlue, Brushes.White, "输入"),
            ScriptConsoleLevel.Result => ("‹", ConsoleInk, Brushes.White, "结果"),
            ScriptConsoleLevel.Debug => ("·", ConsoleMuted, Brushes.White, "调试"),
            ScriptConsoleLevel.Info => ("i", ConsoleInk, Brushes.White, "信息"),
            _ => ("", ConsoleInk, Brushes.White, "日志")
        };
        var line = new LayoutGrid { ColumnDefinitions = new("26,*,66"), Margin = new(0, 2) };
        line.Children.Add(new TextBlock { Text = icon, FontSize = 13, Foreground = ink, HorizontalAlignment = HorizontalAlignment.Center });
        var message = new SelectableTextBlock { Text = entry.DisplayText, FontFamily = new("Consolas"), FontSize = 12,
            Foreground = ink, TextWrapping = TextWrapping.Wrap, Margin = new(2, 0, 8, 0) };
        AutomationProperties.SetName(message, $"{label}：{entry.DisplayText}"); LayoutGrid.SetColumn(message, 1); line.Children.Add(message);
        var time = new TextBlock { Text = entry.Time.ToLocalTime().ToString("HH:mm:ss"), FontFamily = new("Consolas"), FontSize = 10,
            Foreground = ConsoleMuted, VerticalAlignment = VerticalAlignment.Top, Margin = new(0, 2, 4, 0) };
        LayoutGrid.SetColumn(time, 2); line.Children.Add(time);
        return new Border { Tag = entry, Background = background, BorderBrush = ConsoleLine, BorderThickness = new(0, 0, 0, 1), Child = line };
    }
    private void NavigateConsoleHistory(int direction)
    {
        if (_consoleInput is null || _consoleInput.IsComposing || _consoleHistory.Count == 0) return;
        if (_consoleHistoryIndex == _consoleHistory.Count) _consolePendingInput = _consoleInput.Text ?? "";
        _consoleHistoryIndex = Math.Clamp(_consoleHistoryIndex + direction, 0, _consoleHistory.Count);
        _consoleInput.Text = _consoleHistoryIndex == _consoleHistory.Count ? _consolePendingInput : _consoleHistory[_consoleHistoryIndex];
        _consoleInput.CaretIndex = _consoleInput.Text.Length;
    }
    private void RefreshScriptConsole(bool force = false)
    {
        if (_consoleOutput is null) return;
        if (_consoleSession != Document.Session.SessionId)
        {
            _consoleSession = Document.Session.SessionId; _consoleHistory.Clear(); _consoleHistoryIndex = 0;
            _consolePendingInput = ""; _consoleLastSequence = -1; if (_consoleInput is not null) _consoleInput.Text = "";
        }
        if (_consoleRun is not null) _consoleRun.IsEnabled = DesktopServices.SupportsScripts && _scriptCancellation is null && _consoleInput?.IsComposing != true;
        if (_consoleStop is not null) _consoleStop.IsEnabled = _scriptCancellation is not null;
        var entries = Document.Session.ConsoleEntries;
        var last = entries.IsEmpty ? 0 : entries[^1].Sequence;
        if (!force && last == _consoleLastSequence) return;
        _consoleLastSequence = last;
        var filter = _consoleFilter?.Text?.Trim() ?? "";
        _consoleVisibleEntries = entries.Where(e => ((_consoleLevel?.SelectedIndex ?? 0) switch {
            1 => e.Level is ScriptConsoleLevel.Log or ScriptConsoleLevel.Info,
            2 => e.Level == ScriptConsoleLevel.Warning, 3 => e.Level == ScriptConsoleLevel.Error,
            4 => e.Level == ScriptConsoleLevel.Debug, _ => true
        }) && (filter.Length == 0 || e.Text.Contains(filter, StringComparison.OrdinalIgnoreCase))).ToArray();
        _consoleOutput.ItemsSource = _consoleVisibleEntries;
        if (_consoleCopy is not null) _consoleCopy.IsEnabled = _consoleVisibleEntries.Length > 0;
        if (_consoleEmpty is not null)
        {
            _consoleEmpty.IsVisible = _consoleVisibleEntries.Length == 0;
            _consoleEmpty.Text = entries.IsEmpty ? "控制台已就绪。输入表达式，或使用 console.log / warn / error 输出消息。" : "没有匹配的输出。";
        }
        if (_consoleCount is not null)
        {
            var errors = entries.Count(e => e.Level == ScriptConsoleLevel.Error); var warnings = entries.Count(e => e.Level == ScriptConsoleLevel.Warning);
            _consoleCount.Text = $"{_consoleVisibleEntries.Length}/{entries.Length}  ·  {errors} 错误  {warnings} 警告";
            _consoleCount.Foreground = errors > 0 ? ConsoleErrorInk : warnings > 0 ? ConsoleWarningInk : ConsoleMuted;
        }
        if (_consoleVisibleEntries.LastOrDefault() is { } latest)
        {
            var output = _consoleOutput;
            Dispatcher.UIThread.Post(() => { if (ReferenceEquals(output, _consoleOutput)) output.ScrollIntoView(latest); }, DispatcherPriority.Loaded);
        }
    }
    private async Task RunConsoleAsync()
    {
        if (_consoleInput is null || _consoleInput.IsComposing || _consoleInput.Text is not { } source || string.IsNullOrWhiteSpace(source)) return;
        if (!DesktopServices.SupportsScripts) throw new PlatformNotSupportedException("此平台的隔离脚本执行尚不可用。");
        if (_scriptCancellation is not null) return;
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync()) return;
        var document = Document;
        if (_consoleHistory.Count == 0 || _consoleHistory[^1] != source) _consoleHistory.Add(source);
        if (_consoleHistory.Count > 50) _consoleHistory.RemoveAt(0);
        _consoleHistoryIndex = _consoleHistory.Count; _consolePendingInput = "";
        if (_scriptOutputTabs is not null) _scriptOutputTabs.SelectedIndex = 0;
        try
        {
            CommandResult? result = null;
            await ScriptWorkAsync(async token =>
            {
                RefreshScriptConsole();
                result = await Task.Run(() => document.RunConsole(source, token), token);
            });
            if (_scriptStatus is not null) _scriptStatus.Text = result?.Changed == true ? "控制台执行完成，修改可撤销。" : "控制台执行完成。";
        }
        catch (OperationCanceledException) { if (_scriptStatus is not null) _scriptStatus.Text = "控制台执行已停止，修改未提交。"; }
        catch (Exception error) { if (_scriptStatus is not null) _scriptStatus.Text = error.Message; }
        finally { RefreshScriptConsole(); }
    }
}
