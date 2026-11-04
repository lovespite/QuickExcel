using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using QuickExcel.Domain;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private LayoutGrid? _scriptEditorPane;
    private GridSplitter? _scriptOutputSplitter;
    private Button? _scriptConsoleToggle;
    private GridLength _scriptEditorHeight, _scriptOutputHeight;
    private readonly List<(string Label, string? Kind, Expander Panel, ListBox Picker)> _scriptModuleGroups = [];
    private string? _scriptNavigationSelection;

    private Control CreateScriptWorkspace(Control output)
    {
        // Both panes scale with the window after dragging, while retaining usable minimum widths.
        var workspace = new LayoutGrid { ColumnDefinitions = new("200*,5,875*") };
        workspace.ColumnDefinitions[0].MinWidth = 150;
        workspace.ColumnDefinitions[2].MinWidth = 420;
        var navigation = new DockPanel { Background = Brush.Parse("#F8F9FA") };
        var heading = new TextBlock { Text = "脚本文件", FontWeight = FontWeight.SemiBold, Margin = new(12, 10) };
        DockPanel.SetDock(heading, Dock.Top); navigation.Children.Add(heading);
        _modulePicker = CreateScriptModulePicker("主入口脚本");
        DockPanel.SetDock(_modulePicker, Dock.Top); navigation.Children.Add(_modulePicker);
        var groups = new StackPanel();
        _scriptModuleGroups.Clear(); _scriptNavigationSelection = null;
        foreach (var (label, kind) in new (string, string?)[] { ("UDF", "formula-functions"), ("Actions", "actions"), ("辅助", null) })
        {
            var picker = CreateScriptModulePicker(label + " 脚本文件列表");
            var panel = new Expander { Header = label, Content = picker, IsExpanded = true,
                HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Stretch };
            AutomationProperties.SetName(panel, label + " 脚本分类");
            _scriptModuleGroups.Add((label, kind, panel, picker)); groups.Children.Add(panel);
        }
        navigation.Children.Add(new ScrollViewer { Content = groups,
            HorizontalScrollBarVisibility = global::Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled });
        workspace.Children.Add(navigation);
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

    private ListBox CreateScriptModulePicker(string name)
    {
        var picker = new ListBox { Background = Brushes.Transparent, BorderThickness = new(0),
            ItemTemplate = new FuncDataTemplate<string>((id, _) =>
            {
                if (id is null || !Document.Snapshot.ScriptModules.Modules.TryGetValue(id, out var module)) return null;
                var item = new StackPanel { Spacing = 2, Margin = new(2, 4) };
                item.Children.Add(new TextBlock { Text = Path.GetFileName(module.Path), TextTrimming = TextTrimming.CharacterEllipsis });
                item.Children.Add(new TextBlock { Text = id, FontSize = 11, Foreground = Brush.Parse("#70757A"), TextTrimming = TextTrimming.CharacterEllipsis });
                ToolTip.SetTip(item, $"{module.Path}\n模块 ID：{id}\n{ScriptModuleProperties(Document.Snapshot.ScriptModules, id)}"); return item;
            }) };
        AutomationProperties.SetName(picker, name);
        picker.SelectionChanged += async (_, _) =>
        {
            if (_scriptPicking || picker.SelectedItem is not string id || id == _selectedModule) return;
            await PerformAsync(async () =>
            {
                try { if (await CommitScriptsAsync()) { await _scriptPage!.SelectAsync(id); _selectedModule = id; } }
                finally { RefreshScriptChrome(); }
            });
            RefreshScriptChrome();
        };
        return picker;
    }

    private static string ScriptModuleProperties(ScriptModuleSet modules, string id)
    {
        var roles = modules.Bindings.Where(b => b.Module == id).Select(b => b.Kind == "formula-functions" ? "UDF" : "Actions").ToArray();
        return $"{(roles.Length == 0 ? "辅助" : string.Join(" / ", roles))} · {(modules.Modules[id].Language == ScriptLanguage.TypeScript ? "TypeScript" : "JavaScript")}";
    }

    private void RefreshScriptNavigation(bool revealSelected)
    {
        if (_modulePicker is null) return;
        var modules = Document.Snapshot.ScriptModules;
        var main = modules.Modules.Where(p => p.Key == "main" ||
            Path.GetFileName(p.Value.Path).Equals("main.js", StringComparison.OrdinalIgnoreCase) ||
            Path.GetFileName(p.Value.Path).Equals("main.ts", StringComparison.OrdinalIgnoreCase)).Select(p => p.Key).Order(StringComparer.Ordinal).ToArray();
        var others = modules.Modules.Keys.Except(main).Order(StringComparer.Ordinal).ToArray();
        void Update(ListBox picker, string[] ids)
        {
            // Routine draft/status refreshes must retain scroll position and collapsed panels.
            if (picker.ItemsSource is not string[] previous || !previous.SequenceEqual(ids)) picker.ItemsSource = ids;
            picker.SelectedItem = ids.Contains(_selectedModule) ? _selectedModule : null;
        }
        _scriptPicking = true;
        try
        {
            Update(_modulePicker, main); _modulePicker.IsVisible = main.Length > 0;
            foreach (var (label, kind, panel, picker) in _scriptModuleGroups)
            {
                var ids = others.Where(id => kind is null ? !modules.Bindings.Any(b => b.Module == id) :
                    modules.Bindings.Any(b => b.Module == id && b.Kind == kind)).ToArray();
                Update(picker, ids); panel.Header = $"{label} ({ids.Length})";
                if ((revealSelected || _scriptNavigationSelection != _selectedModule) && ids.Contains(_selectedModule)) panel.IsExpanded = true;
            }
            _scriptNavigationSelection = _selectedModule;
        }
        finally { _scriptPicking = false; }
    }

    private ListBox? SelectedScriptModulePicker => _modulePicker?.SelectedItem is not null ? _modulePicker :
        _scriptModuleGroups.FirstOrDefault(g => g.Picker.SelectedItem is not null).Picker;

    private void FocusScriptNavigation()
    {
        var picker = SelectedScriptModulePicker;
        foreach (var group in _scriptModuleGroups.Where(g => ReferenceEquals(g.Picker, picker))) group.Panel.IsExpanded = true;
        picker?.Focus();
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
