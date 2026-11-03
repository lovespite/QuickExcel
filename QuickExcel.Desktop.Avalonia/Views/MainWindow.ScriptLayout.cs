using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private LayoutGrid? _scriptEditorPane;
    private GridSplitter? _scriptOutputSplitter;
    private Button? _scriptConsoleToggle;
    private GridLength _scriptEditorHeight, _scriptOutputHeight;

    private Control CreateScriptWorkspace(Control output)
    {
        // Both panes scale with the window after dragging, while retaining usable minimum widths.
        var workspace = new LayoutGrid { ColumnDefinitions = new("200*,5,875*") };
        workspace.ColumnDefinitions[0].MinWidth = 150;
        workspace.ColumnDefinitions[2].MinWidth = 420;
        var navigation = new DockPanel { Background = Brush.Parse("#F8F9FA") };
        var heading = new TextBlock { Text = "脚本文件", FontWeight = FontWeight.SemiBold, Margin = new(12, 10) };
        DockPanel.SetDock(heading, Dock.Top); navigation.Children.Add(heading);
        _modulePicker = new ListBox { Background = Brushes.Transparent, BorderThickness = new(0),
            ItemTemplate = new FuncDataTemplate<string>((id, _) =>
            {
                if (id is null || !Document.Snapshot.ScriptModules.Modules.TryGetValue(id, out var module)) return null;
                var item = new StackPanel { Spacing = 2, Margin = new(2, 4) };
                item.Children.Add(new TextBlock { Text = Path.GetFileName(module.Path), TextTrimming = TextTrimming.CharacterEllipsis });
                item.Children.Add(new TextBlock { Text = id, FontSize = 11, Foreground = Brush.Parse("#70757A"), TextTrimming = TextTrimming.CharacterEllipsis });
                ToolTip.SetTip(item, $"{module.Path}\n模块 ID：{id}"); return item;
            }) };
        AutomationProperties.SetName(_modulePicker, "脚本文件列表");
        navigation.Children.Add(_modulePicker); workspace.Children.Add(navigation);
        var splitter = new GridSplitter { ResizeDirection = GridResizeDirection.Columns, ResizeBehavior = GridResizeBehavior.PreviousAndNext,
            HorizontalAlignment = HorizontalAlignment.Stretch, Background = Brush.Parse("#E0E0E0") };
        AutomationProperties.SetName(splitter, "调整文件列表宽度");
        LayoutGrid.SetColumn(splitter, 1); workspace.Children.Add(splitter);

        _scriptEditorPane = new LayoutGrid { RowDefinitions = new("*,0,0") };
        _scriptEditorPane.RowDefinitions[0].MinHeight = 100;
        _scriptContent = new ContentControl(); _scriptEditorPane.Children.Add(_scriptContent);
        _scriptOutputSplitter = new GridSplitter { ResizeDirection = GridResizeDirection.Rows, ResizeBehavior = GridResizeBehavior.PreviousAndNext,
            HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch, Background = Brush.Parse("#E0E0E0"), IsVisible = false };
        AutomationProperties.SetName(_scriptOutputSplitter, "调整控制台高度");
        LayoutGrid.SetRow(_scriptOutputSplitter, 1); _scriptEditorPane.Children.Add(_scriptOutputSplitter);
        output.IsVisible = false; LayoutGrid.SetRow(output, 2); _scriptEditorPane.Children.Add(output);
        _scriptEditorHeight = new(3, GridUnitType.Star); _scriptOutputHeight = new(2, GridUnitType.Star);
        UpdateScriptConsoleToggle(false);
        LayoutGrid.SetColumn(_scriptEditorPane, 2); workspace.Children.Add(_scriptEditorPane);
        return workspace;
    }

    private void HandleScriptWorkspaceKey(object? sender, KeyEventArgs e)
    {
        if (e.Handled || e.Key != Key.F12 || e.KeyModifiers != KeyModifiers.None || _consoleInput?.IsComposing == true) return;
        e.Handled = true; ToggleScriptConsole();
    }

    private void ToggleScriptConsole()
    {
        if (_scriptExecution?.IsRunning == true || _consoleInput?.IsComposing == true) return;
        SetScriptOutputVisible(_scriptOutputTabs?.IsVisible != true || _scriptOutputTabs.SelectedIndex != 0, focusInput: true);
    }

    private void SetScriptOutputVisible(bool visible, int tab = 0, bool focusInput = false)
    {
        if (_scriptEditorPane is null || _scriptOutputTabs is null || _scriptOutputSplitter is null) return;
        var rows = _scriptEditorPane.RowDefinitions;
        if (visible != _scriptOutputTabs.IsVisible)
        {
            if (!visible) { _scriptEditorHeight = rows[0].Height; _scriptOutputHeight = rows[2].Height; }
            rows[0].Height = visible ? _scriptEditorHeight : new(1, GridUnitType.Star);
            rows[1].Height = new(visible ? 5 : 0);
            rows[2].MinHeight = visible ? 190 : 0;
            rows[2].Height = visible ? _scriptOutputHeight : new(0);
            _scriptOutputTabs.IsVisible = _scriptOutputSplitter.IsVisible = visible;
        }
        if (visible) _scriptOutputTabs.SelectedIndex = tab;
        UpdateScriptConsoleToggle(visible && tab == 0);
        if (focusInput)
        {
            // Newly expanded content must be arranged before it can accept keyboard focus.
            var tabs = _scriptOutputTabs;
            Dispatcher.UIThread.Post(() =>
            {
                if (!ReferenceEquals(tabs, _scriptOutputTabs) || tabs.IsVisible != visible) return;
                if (visible) _consoleInput?.Focus(); else _scriptPage?.Browser.View.Focus();
            }, DispatcherPriority.Loaded);
        }
    }

    private void UpdateScriptConsoleToggle(bool visible)
    {
        if (_scriptConsoleToggle is null) return;
        var label = visible ? "隐藏控制台 · F12" : "显示控制台 · F12";
        _scriptConsoleToggle.Content = label; AutomationProperties.SetName(_scriptConsoleToggle, label);
    }
}
