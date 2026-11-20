using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Interactivity;
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
    private Button? _consoleRun, _consoleCopy;
    private ScriptConsoleEntry[] _consoleVisibleEntries = [];
    private TabControl? _scriptOutputTabs;
    private readonly List<string> _consoleHistory = [];
    private Guid _consoleSession;
    private int _consoleHistoryIndex;
    private string _consolePendingInput = "";
    private long _consoleLastSequence = -1;
    // The script workspace uses the dark Fluent theme, including focused/hovered input templates.
    private static readonly IBrush ConsoleBackground = Brush.Parse("#1E1E1E"), ConsoleToolbar = Brush.Parse("#252526"),
        ConsoleInk = Brush.Parse("#D4D4D4"), ConsoleMuted = Brush.Parse("#ABABAB"),
        ConsoleLine = Brush.Parse("#3F3F46"), ConsoleBlue = Brush.Parse("#75BEFF"),
        ConsoleWarningInk = Brush.Parse("#FFD866"), ConsoleWarningBackground = Brush.Parse("#332B18"),
        ConsoleErrorInk = Brush.Parse("#FF9B9B"), ConsoleErrorBackground = Brush.Parse("#3F2024");

    private string ConsoleVisibleText => string.Join(Environment.NewLine, _consoleVisibleEntries.Select(e =>
        $"{e.Time.ToLocalTime():HH:mm:ss} {(e.Kind == "input" ? "> " : "")}{e.Text}"));

    private async Task OpenScriptConsoleAsync()
    {
        await OpenScriptsAsync();
        SetScriptOutputVisible(true, focusInput: true);
    }

    private Control CreateScriptOutput()
    {
        if (_consoleSession != Document.Session.SessionId)
        { _consoleSession = Document.Session.SessionId; _consoleHistory.Clear(); _consolePendingInput = ""; }
        _consoleHistoryIndex = _consoleHistory.Count; _consoleLastSequence = -1;
        var panel = new LayoutGrid { RowDefinitions = new("Auto,*,Auto,Auto"), Background = ConsoleBackground };
        var tools = new LayoutGrid { ColumnDefinitions = new("Auto,*,Auto,Auto"), Margin = new(6, 3) };
        var clear = ConsoleButton("⊘", UiText.Get("清空输出"));
        clear.Click += (_, _) => { Document.Session.ClearConsole(); RefreshScriptConsole(true); }; tools.Children.Add(clear);
        _consoleFilter = new TextBox { Height = 26, MinHeight = 0, Padding = new(7, 3), Margin = new(6, 0), FontSize = 12,
            PlaceholderText = UiText.Get("过滤输出"), PlaceholderForeground = Brushes.White,
            BorderBrush = ConsoleLine, Background = ConsoleBackground, Foreground = ConsoleInk };
        AutomationProperties.SetName(_consoleFilter, UiText.Get("过滤控制台输出"));
        _consoleFilter.TextChanged += (_, _) => RefreshScriptConsole(true);
        LayoutGrid.SetColumn(_consoleFilter, 1); tools.Children.Add(_consoleFilter);
        _consoleLevel = new ComboBox { ItemsSource = new[] { UiText.Get("所有级别"), UiText.Get("日志 / 信息"), UiText.Get("警告"), UiText.Get("错误"), UiText.Get("调试") },
            SelectedIndex = 0, Width = 110, Height = 26, MinHeight = 0, Padding = new(7, 2), FontSize = 12 };
        AutomationProperties.SetName(_consoleLevel, UiText.Get("控制台日志级别"));
        _consoleLevel.SelectionChanged += (_, _) => RefreshScriptConsole(true);
        LayoutGrid.SetColumn(_consoleLevel, 2); tools.Children.Add(_consoleLevel);
        _consoleCopy = ConsoleButton(UiText.Get("复制"), UiText.Get("复制当前筛选的控制台输出"));
        _consoleCopy.Click += async (_, _) => await PerformAsync(() => _clipboard.WriteAsync(new(ConsoleVisibleText)));
        LayoutGrid.SetColumn(_consoleCopy, 3); tools.Children.Add(_consoleCopy);
        _consoleCount = new TextBlock { FontSize = 11, Foreground = ConsoleMuted, VerticalAlignment = VerticalAlignment.Center, Margin = new(8, 0) };
        var consoleToolbar = new StackPanel(); consoleToolbar.Children.Add(tools); consoleToolbar.Children.Add(_consoleCount);
        panel.Children.Add(new Border { Background = ConsoleToolbar, BorderBrush = ConsoleLine, BorderThickness = new(0, 1, 0, 1), Child = consoleToolbar });

        _consoleOutput = new ListBox { BorderThickness = new(0), Padding = new(0), Background = ConsoleBackground,
            ItemTemplate = new FuncDataTemplate<ScriptConsoleEntry>((entry, _) => entry is null ? null : CreateConsoleRow(entry)) };
        _consoleOutput.Styles.Add(new Style(x => x.OfType<ListBoxItem>()) { Setters = {
            new Setter(TemplatedControl.PaddingProperty, new Thickness(0)), new Setter(Layoutable.MarginProperty, new Thickness(0)),
            new Setter(Layoutable.MinHeightProperty, 0d), new Setter(ContentControl.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch) } });
        ScrollViewer.SetHorizontalScrollBarVisibility(_consoleOutput, ScrollBarVisibility.Disabled);
        AutomationProperties.SetName(_consoleOutput, UiText.Get("调试控制台输出，可选择复制"));
        var outputArea = new LayoutGrid(); outputArea.Children.Add(_consoleOutput);
        _consoleEmpty = new TextBlock { Text = UiText.Get("控制台已就绪。输入表达式，或使用 console.log / warn / error 输出消息。"), FontSize = 12,
            Foreground = ConsoleMuted, Margin = new(28, 12), IsHitTestVisible = false, TextWrapping = TextWrapping.Wrap };
        outputArea.Children.Add(_consoleEmpty); LayoutGrid.SetRow(outputArea, 1); panel.Children.Add(outputArea);

        var inputArea = new LayoutGrid { ColumnDefinitions = new("26,*") };
        inputArea.Children.Add(new TextBlock { Text = "›", FontFamily = new("Consolas"), FontSize = 20, Foreground = ConsoleBlue, Margin = new(8, 1) });
        _consoleInput = new CompositionTextBox { AcceptsReturn = true, Height = 54, MinHeight = 0, FontFamily = new("Consolas"), FontSize = 13,
            Padding = new(2, 5), BorderThickness = new(0), Background = ConsoleBackground, Foreground = ConsoleInk,
            PlaceholderText = UiText.Get("输入 JavaScript 表达式或代码…"), PlaceholderForeground = Brushes.White,
            MaxLength = DocumentSession.MaxConsoleSourceBytes };
        ToolTip.SetTip(_consoleInput, UiText.Get("Enter 执行，Ctrl+Enter 换行，↑↓ 浏览历史。可用 workbook 和 host；工作簿修改可撤销。"));
        AutomationProperties.SetName(_consoleInput, UiText.Get("调试控制台 JavaScript 输入"));
        _consoleInput.CompositionChanged += () => RefreshScriptConsole();
        // TextBox consumes Enter/arrow keys during bubbling. Handle console keys before its editor does.
        _consoleInput.AddHandler(KeyDownEvent, HandleConsoleKey, RoutingStrategies.Tunnel);
        LayoutGrid.SetColumn(_consoleInput, 1); inputArea.Children.Add(_consoleInput);
        var inputBorder = new Border { BorderBrush = ConsoleLine, BorderThickness = new(0, 1, 0, 0), Child = inputArea };
        LayoutGrid.SetRow(inputBorder, 2); panel.Children.Add(inputBorder);
        var footer = new DockPanel { Background = ConsoleToolbar, LastChildFill = true };
        var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 2, Margin = new(4, 0) };
        _consoleRun = ConsoleButton(UiText.Get("▶ 执行"), UiText.Get("执行控制台代码（Enter）"));
        _consoleRun.Foreground = ConsoleBlue; _consoleRun.Click += async (_, _) => await PerformAsync(RunConsoleAsync); actions.Children.Add(_consoleRun);
        DockPanel.SetDock(actions, Dock.Right); footer.Children.Add(actions);
        footer.Children.Add(new TextBlock { Text = UiText.Get("JavaScript  ·  Enter 执行  ·  Ctrl+Enter 换行  ·  ↑↓ 历史"), FontSize = 11, Foreground = ConsoleMuted,
            VerticalAlignment = VerticalAlignment.Center, Margin = new(9, 0), TextWrapping = TextWrapping.Wrap });
        LayoutGrid.SetRow(footer, 3); panel.Children.Add(footer);
        _scriptOutputTabs = new TabControl { Background = ConsoleBackground, Foreground = ConsoleInk, ItemsSource = new[] {
            new TabItem { Header = UiText.Get("控制台"), FontSize = 13, MinHeight = 30, Padding = new(10, 6), Content = panel },
            new TabItem { Header = UiText.Get("运行诊断"), FontSize = 13, MinHeight = 30, Padding = new(10, 6), Content = _scriptDiagnostics },
            new TabItem { Header = UiText.Get("网络"), FontSize = 13, MinHeight = 30, Padding = new(10, 6), Content = CreateScriptNetworkPanel() } } };
        _scriptOutputTabs.SelectionChanged += (_, _) => { UpdateScriptConsoleToggle(_scriptOutputTabs.IsVisible && _scriptOutputTabs.SelectedIndex == 0); UpdateNetworkRecording(); };
        _scriptOutputTabs.PropertyChanged += (_, change) => { if (change.Property == IsVisibleProperty) UpdateNetworkRecording(); };
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
            ScriptConsoleLevel.Warning => ("⚠", ConsoleWarningInk, ConsoleWarningBackground, UiText.Get("警告")),
            ScriptConsoleLevel.Error => ("⊗", ConsoleErrorInk, ConsoleErrorBackground, UiText.Get("错误")),
            ScriptConsoleLevel.Input => ("›", ConsoleBlue, ConsoleBackground, UiText.Get("输入")),
            ScriptConsoleLevel.Result => ("‹", ConsoleInk, ConsoleBackground, UiText.Get("结果")),
            ScriptConsoleLevel.Debug => ("·", ConsoleMuted, ConsoleBackground, UiText.Get("调试")),
            ScriptConsoleLevel.Info => ("i", ConsoleInk, ConsoleBackground, UiText.Get("信息")),
            _ => ("", ConsoleInk, ConsoleBackground, UiText.Get("日志"))
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
    private async void HandleConsoleKey(object? sender, KeyEventArgs e)
    {
        if (e.Handled || sender is not CompositionTextBox input || input.IsComposing) return;
        if (e.Key == Key.Enter && e.KeyModifiers == KeyModifiers.None)
        { e.Handled = true; await PerformAsync(RunConsoleAsync); }
        else if (e.Key == Key.Enter && e.KeyModifiers is KeyModifiers.Control or KeyModifiers.Meta)
        { e.Handled = true; input.SelectedText = input.NewLine; }
        else if (e.KeyModifiers == KeyModifiers.None && e.Key is Key.Up or Key.Down)
        { e.Handled = true; NavigateConsoleHistory(e.Key == Key.Up ? -1 : 1); }
    }
    private static void SetConsoleCaret(CompositionTextBox input, int position)
    { input.CaretIndex = input.SelectionStart = input.SelectionEnd = position; }
    private void NavigateConsoleHistory(int direction)
    {
        if (_consoleInput is not { } input || input.IsComposing) return;
        var boundary = direction < 0 ? 0 : input.Text?.Length ?? 0;
        if (input.CaretIndex != boundary || input.SelectionStart != input.SelectionEnd)
        { SetConsoleCaret(input, boundary); return; }
        if (_consoleHistory.Count == 0) return;
        if (_consoleHistoryIndex == _consoleHistory.Count) _consolePendingInput = input.Text ?? "";
        _consoleHistoryIndex = Math.Clamp(_consoleHistoryIndex + direction, 0, _consoleHistory.Count);
        input.Text = _consoleHistoryIndex == _consoleHistory.Count ? _consolePendingInput : _consoleHistory[_consoleHistoryIndex];
        SetConsoleCaret(input, direction < 0 ? 0 : input.Text.Length);
    }
    private void RefreshScriptConsole(bool force = false)
    {
        if (_consoleOutput is null) return;
        if (_consoleSession != Document.Session.SessionId)
        {
            _consoleSession = Document.Session.SessionId; _consoleHistory.Clear(); _consoleHistoryIndex = 0;
            _consolePendingInput = ""; _consoleLastSequence = -1; if (_consoleInput is not null) _consoleInput.Text = "";
        }
        if (_consoleRun is not null) _consoleRun.IsEnabled = DesktopServices.SupportsScripts && _consoleInput?.IsComposing != true;
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
            _consoleEmpty.Text = entries.IsEmpty ? UiText.Get("控制台已就绪。输入表达式，或使用 console.log / warn / error 输出消息。") : UiText.Get("没有匹配的输出。");
        }
        if (_consoleCount is not null)
        {
            var errors = entries.Count(e => e.Level == ScriptConsoleLevel.Error); var warnings = entries.Count(e => e.Level == ScriptConsoleLevel.Warning);
            _consoleCount.Text = UiText.Format($"{_consoleVisibleEntries.Length}/{entries.Length}  ·  {errors} 错误  {warnings} 警告");
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
        RequireScriptsEnabled();
        if (_consoleInput is not { } input || input.IsComposing || input.Text is not { } source || string.IsNullOrWhiteSpace(source)) return;
        if (!DesktopServices.SupportsScripts) throw new PlatformNotSupportedException(UiText.Get("此平台的隔离脚本执行尚不可用。"));
        if (_scriptCancellation is not null) return;
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync()) return;
        if (!ReferenceEquals(input, _consoleInput) || _scriptCancellation is not null) return;
        var document = Document;
        if (_consoleHistory.Count == 0 || _consoleHistory[^1] != source) _consoleHistory.Add(source);
        if (_consoleHistory.Count > 50) _consoleHistory.RemoveAt(0);
        _consoleHistoryIndex = _consoleHistory.Count; _consolePendingInput = "";
        SetScriptOutputVisible(true);
        // Clear only the accepted command, before execution; never erase a newer draft on completion.
        if (input.Text == source && !input.IsComposing) { input.Text = ""; SetConsoleCaret(input, 0); }
        input.Focus();
        try
        {
            CommandResult? result = null;
            await ScriptWorkAsync(async token =>
            {
                RefreshScriptConsole();
                result = await Task.Run(() => document.RunConsole(source, token), token);
            });
            if (_scriptStatus is not null) _scriptStatus.Text = result?.Changed == true ? UiText.Get("控制台执行完成，修改可撤销。") : UiText.Get("控制台执行完成。");
        }
        catch (OperationCanceledException) { if (_scriptStatus is not null) _scriptStatus.Text = UiText.Get("控制台执行已停止，修改未提交。"); }
        catch (Exception error) { if (_scriptStatus is not null) _scriptStatus.Text = error.Message; }
        finally { RefreshScriptConsole(); }
    }
}
