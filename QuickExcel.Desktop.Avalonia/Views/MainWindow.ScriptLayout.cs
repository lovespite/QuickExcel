using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Presenters;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using QuickExcel.Domain;
using QuickExcel.Desktop.Avalonia.Web;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private LayoutGrid? _scriptEditorPane;
    private GridSplitter? _scriptOutputSplitter;
    private MenuItem? _scriptConsoleToggle;
    private GridLength _scriptEditorHeight, _scriptOutputHeight;
    private readonly List<(string Label, string? Kind, Expander Panel, ListBox Picker)> _scriptModuleGroups = [];
    private string? _scriptNavigationSelection;
    private string? _selectedDefinition;
    private ListBox? _definitionPicker;
    private Expander? _definitionPanel;
    private readonly List<string> _scriptDefinitionFiles = [];

    private Control CreateScriptWorkspace(Control output)
    {
        // Both panes scale with the window after dragging, while retaining usable minimum widths.
        var workspace = new LayoutGrid { ColumnDefinitions = new("240*,1,839*") };
        workspace.ColumnDefinitions[0].MinWidth = 150;
        workspace.ColumnDefinitions[2].MinWidth = 420;
        var navigation = new DockPanel { Background = Brush.Parse("#181818") };
        var heading = new DockPanel { Margin = new(16, 6, 8, 6) };
        var add = new Button { Content = new global::Avalonia.Controls.Shapes.Path { Width = 20, Height = 20,
            Data = Geometry.Parse("M3,2 L11,2 L15,6 L15,17 L3,17 Z M11,2 L11,6 L15,6 M11,12 L19,12 M15,8 L15,16"),
            Stroke = Brush.Parse("#A0A0A0"), StrokeThickness = 1.2 }, Width = 30, Height = 30, Padding = new(0), Background = Brushes.Transparent, BorderThickness = new(0) };
        ToolTip.SetTip(add, UiText.Get("添加模块")); AutomationProperties.SetName(add, UiText.Get("添加模块"));
        add.Click += async (_, _) => await PerformAsync(AddScriptModuleAsync); DockPanel.SetDock(add, Dock.Right); heading.Children.Add(add);
        heading.Children.Add(new TextBlock { Text = "Scripts", VerticalAlignment = VerticalAlignment.Center, FontSize = 13 });
        DockPanel.SetDock(heading, Dock.Top); navigation.Children.Add(heading);
        _modulePicker = CreateScriptModulePicker(UiText.Get("主入口脚本"));
        DockPanel.SetDock(_modulePicker, Dock.Top); navigation.Children.Add(_modulePicker);
        var groups = new StackPanel();
        _scriptModuleGroups.Clear(); _scriptNavigationSelection = null;
        foreach (var (label, kind) in new (string, string?)[] { ("Actions", "actions"), ("Common", null), ("Events", "events"), ("UDF", "formula-functions") })
        {
            var picker = CreateScriptModulePicker(label + UiText.Get(" 脚本文件列表"));
            var panel = new Expander { Header = label, Content = new Border { BorderBrush = Brush.Parse("#333333"), BorderThickness = new(1, 0, 0, 0), Margin = new(27, 0, 0, 0), Child = picker }, IsExpanded = true, Padding = new(0),
                HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Stretch };
            UseScriptTreeTemplate(panel);
            AutomationProperties.SetName(panel, label + UiText.Get(" 脚本分类"));
            _scriptModuleGroups.Add((label, kind, panel, picker)); groups.Children.Add(panel);
        }
        _selectedDefinition = null; _scriptDefinitionFiles.Clear(); _scriptDefinitionFiles.Add("quickexcel.d.ts");
        _definitionPicker = CreateCompactScriptPicker();
        _definitionPicker.ItemTemplate = new FuncDataTemplate<string>((name, _) => new TextBlock
            { Text = name, TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center });
        AutomationProperties.SetName(_definitionPicker, UiText.Get("只读定义文件列表"));
        _definitionPicker.SelectionChanged += async (_, _) =>
        {
            if (_scriptPicking || _definitionPicker.SelectedItem is not string name || name == _selectedDefinition) return;
            await NavigateScriptAsync(new(null, name)); RefreshScriptChrome();
        };
        _definitionPanel = new Expander { Header = UiText.Get("定义文件（只读）"), Content = _definitionPicker, IsExpanded = true, Padding = new(0),
            HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Stretch };
        UseScriptTreeTemplate(_definitionPanel);
        groups.Children.Add(_definitionPanel);
        navigation.Children.Add(new ScrollViewer { Content = groups,
            HorizontalScrollBarVisibility = global::Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled });
        workspace.Children.Add(navigation);
        var splitter = new GridSplitter { ResizeDirection = GridResizeDirection.Columns, ResizeBehavior = GridResizeBehavior.PreviousAndNext,
            HorizontalAlignment = HorizontalAlignment.Stretch, Background = Brush.Parse("#2B2B2B") };
        AutomationProperties.SetName(splitter, UiText.Get("调整文件列表宽度"));
        LayoutGrid.SetColumn(splitter, 1); workspace.Children.Add(splitter);

        _scriptEditorPane = new LayoutGrid { RowDefinitions = new("*,0,0") };
        _scriptEditorPane.RowDefinitions[0].MinHeight = 100;
        _scriptContent = new ContentControl();
        var editor = new DockPanel(); var editorHeader = CreateScriptEditorHeader(); DockPanel.SetDock(editorHeader, Dock.Top); editor.Children.Add(editorHeader); editor.Children.Add(_scriptContent);
        _scriptEditorPane.Children.Add(editor);
        _scriptOutputSplitter = new GridSplitter { ResizeDirection = GridResizeDirection.Rows, ResizeBehavior = GridResizeBehavior.PreviousAndNext,
            HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch, Background = Brush.Parse("#2B2B2B"), IsVisible = false };
        AutomationProperties.SetName(_scriptOutputSplitter, UiText.Get("调整控制台高度"));
        LayoutGrid.SetRow(_scriptOutputSplitter, 1); _scriptEditorPane.Children.Add(_scriptOutputSplitter);
        output.IsVisible = false; LayoutGrid.SetRow(output, 2); _scriptEditorPane.Children.Add(output);
        _scriptEditorHeight = new(3, GridUnitType.Star); _scriptOutputHeight = new(2, GridUnitType.Star);
        UpdateScriptConsoleToggle(false);
        LayoutGrid.SetColumn(_scriptEditorPane, 2); workspace.Children.Add(_scriptEditorPane);
        return workspace;
    }

    private static ListBox CreateCompactScriptPicker()
    {
        var picker = new ListBox { Background = Brushes.Transparent, BorderThickness = new(0), Padding = new(0),
            HorizontalAlignment = HorizontalAlignment.Stretch };
        ScrollViewer.SetHorizontalScrollBarVisibility(picker, global::Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled);
        picker.Styles.Add(new Style(s => s.OfType<ListBoxItem>()) { Setters =
        {
            new Setter(MinHeightProperty, 28d), new Setter(HeightProperty, 28d), new Setter(MarginProperty, new Thickness(0)),
            new Setter(PaddingProperty, new Thickness(12, 0)), new Setter(HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch),
            new Setter(VerticalContentAlignmentProperty, VerticalAlignment.Center), new Setter(CornerRadiusProperty, new CornerRadius(0)),
            new Setter(TemplateProperty, new FuncControlTemplate<ListBoxItem>((item, _) =>
            {
                var content = new ContentPresenter { Margin = new(12, 0), VerticalAlignment = VerticalAlignment.Center };
                content.Bind(ContentPresenter.ContentProperty, item.GetObservable(ContentControl.ContentProperty));
                content.Bind(ContentPresenter.ContentTemplateProperty, item.GetObservable(ContentControl.ContentTemplateProperty));
                var border = new Border { Child = content, Background = item.IsSelected ? Brush.Parse("#37373D") : Brushes.Transparent };
                item.PropertyChanged += (_, change) => { if (change.Property == ListBoxItem.IsSelectedProperty) border.Background = item.IsSelected ? Brush.Parse("#37373D") : Brushes.Transparent; };
                return border;
            }))
        } });
        picker.Styles.Add(new Style(s => s.OfType<ListBoxItem>().Class(":selected")) { Setters =
        { new Setter(BackgroundProperty, Brush.Parse("#37373D")), new Setter(ForegroundProperty, Brush.Parse("#FFFFFF")) } });
        return picker;
    }

    private static void UseScriptTreeTemplate(Expander panel)
    {
        panel.Template = new FuncControlTemplate<Expander>((owner, _) =>
        {
            var rows = new StackPanel();
            // Both chevrons share a centered 12px viewport, independent of font baselines.
            var expandedArrow = Geometry.Parse("M2,4 L6,8 L10,4");
            var collapsedArrow = Geometry.Parse("M4,2 L8,6 L4,10");
            var arrow = new global::Avalonia.Controls.Shapes.Path { Data = owner.IsExpanded ? expandedArrow : collapsedArrow,
                Width = 12, Height = 12, Stretch = Stretch.None, Stroke = Brush.Parse("#A0A0A0"), StrokeThickness = 1.25,
                HorizontalAlignment = HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Center, IsHitTestVisible = false };
            var label = new TextBlock { Text = owner.Header?.ToString(), FontSize = 14, VerticalAlignment = VerticalAlignment.Center, Foreground = Brush.Parse("#CCCCCC") };
            var caption = new LayoutGrid { ColumnDefinitions = new("17,*"), VerticalAlignment = VerticalAlignment.Center };
            caption.Children.Add(arrow); LayoutGrid.SetColumn(label, 1); caption.Children.Add(label);
            var toggle = new Button { Content = caption, Height = 28, Padding = new(16, 0, 0, 0), Background = Brushes.Transparent,
                BorderThickness = new(0), HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Left,
                VerticalContentAlignment = VerticalAlignment.Center };
            var content = new ContentPresenter { Content = owner.Content, IsVisible = owner.IsExpanded };
            toggle.Click += (_, _) => owner.IsExpanded = !owner.IsExpanded;
            owner.PropertyChanged += (_, change) =>
            {
                if (change.Property == Expander.IsExpandedProperty) { content.IsVisible = owner.IsExpanded; arrow.Data = owner.IsExpanded ? expandedArrow : collapsedArrow; }
                else if (change.Property == HeaderedContentControl.HeaderProperty) label.Text = owner.Header?.ToString();
            };
            rows.Children.Add(toggle); rows.Children.Add(content); return rows;
        });
    }

    private ListBox CreateScriptModulePicker(string name)
    {
        var picker = CreateCompactScriptPicker();
        picker.ItemTemplate = new FuncDataTemplate<string>((id, _) =>
            {
                if (id is null || !Document.Snapshot.ScriptModules.Modules.TryGetValue(id, out var module)) return null;
                var item = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, VerticalAlignment = VerticalAlignment.Center };
                item.Children.Add(new TextBlock { Text = module.Language == ScriptLanguage.TypeScript ? "TS" : "JS", FontSize = 11, FontWeight = FontWeight.Bold,
                    Foreground = Brush.Parse(module.Language == ScriptLanguage.TypeScript ? "#519ABA" : "#CBCB41"), VerticalAlignment = VerticalAlignment.Center });
                item.Children.Add(new TextBlock { Text = Path.GetFileName(module.Path), TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center });
                ToolTip.SetTip(item, UiText.Format($"{module.Path}\n模块 ID：{id}\n{ScriptModuleProperties(Document.Snapshot.ScriptModules, id)}") +
                    (string.IsNullOrWhiteSpace(module.Description) ? string.Empty : UiText.Format($"\n描述：{module.Description}"))); return item;
            });
        AutomationProperties.SetName(picker, name);
        picker.SelectionChanged += async (_, _) =>
        {
            if (_scriptPicking || picker.SelectedItem is not string id || id == _selectedModule && _selectedDefinition is null) return;
            await PerformAsync(async () =>
            {
                try { if (await CommitScriptsAsync()) { await _scriptPage!.SelectAsync(id); _selectedModule = id; _selectedDefinition = null; } }
                finally { RefreshScriptChrome(); }
            });
            RefreshScriptChrome();
        };
        return picker;
    }

    private static string ScriptModuleProperties(ScriptModuleSet modules, string id)
    {
        var roles = modules.Bindings.Where(b => b.Module == id).Select(b => b.Kind switch { "formula-functions" => "UDF", "events" => "Events", "initialization" => UiText.Get("主入口"), _ => "Actions" }).ToArray();
        var role = IsMainScriptModule(modules, id) ? UiText.Get("主入口") : roles.Length == 0 ? "Common" : string.Join(" / ", roles);
        return $"{role} · {(modules.Modules[id].Language == ScriptLanguage.TypeScript ? "TypeScript" : "JavaScript")}";
    }

    private static bool IsMainScriptModule(ScriptModuleSet modules, string id) => modules.ScriptApiVersion == 4
        ? modules.Bindings.Any(b => b.Module == id && b.Kind == "initialization")
        : id == "main" || Path.GetFileName(modules.Modules[id].Path).Equals("main.js", StringComparison.OrdinalIgnoreCase) ||
            Path.GetFileName(modules.Modules[id].Path).Equals("main.ts", StringComparison.OrdinalIgnoreCase);

    private void RefreshScriptNavigation(bool revealSelected)
    {
        if (_modulePicker is null) return;
        var modules = Document.Snapshot.ScriptModules;
        var main = modules.Modules.Where(p => IsMainScriptModule(modules, p.Key)).Select(p => p.Key).Order(StringComparer.Ordinal).ToArray();
        var others = modules.Modules.Keys.Except(main).Order(StringComparer.Ordinal).ToArray();
        void Update(ListBox picker, string[] ids)
        {
            // Routine draft/status refreshes must retain scroll position and collapsed panels.
            if (picker.ItemsSource is not string[] previous || !previous.SequenceEqual(ids)) picker.ItemsSource = ids;
            picker.SelectedItem = _selectedDefinition is null && ids.Contains(_selectedModule) ? _selectedModule : null;
        }
        _scriptPicking = true;
        try
        {
            Update(_modulePicker, main); _modulePicker.IsVisible = main.Length > 0;
            foreach (var (label, kind, panel, picker) in _scriptModuleGroups)
            {
                var ids = others.Where(id => kind is null ? !modules.Bindings.Any(b => b.Module == id) :
                    modules.Bindings.Any(b => b.Module == id && b.Kind == kind)).ToArray();
                Update(picker, ids); panel.Header = label;
                if (_selectedDefinition is null && (revealSelected || _scriptNavigationSelection != _selectedModule) && ids.Contains(_selectedModule)) panel.IsExpanded = true;
            }
            if (_definitionPicker is not null)
            {
                if (_definitionPicker.ItemsSource is not string[] prior || !prior.SequenceEqual(_scriptDefinitionFiles))
                    _definitionPicker.ItemsSource = _scriptDefinitionFiles.ToArray();
                _definitionPicker.SelectedItem = _selectedDefinition;
                if (revealSelected && _selectedDefinition is not null) _definitionPanel!.IsExpanded = true;
            }
            _scriptNavigationSelection = _selectedDefinition is null ? _selectedModule : null;
        }
        finally { _scriptPicking = false; }
    }

    private ListBox? SelectedScriptModulePicker => _selectedDefinition is not null ? _definitionPicker : _modulePicker?.SelectedItem is not null ? _modulePicker :
        _scriptModuleGroups.FirstOrDefault(g => g.Picker.SelectedItem is not null).Picker;

    private void FocusScriptNavigation()
    {
        var picker = SelectedScriptModulePicker;
        if (_selectedDefinition is not null) _definitionPanel!.IsExpanded = true;
        foreach (var group in _scriptModuleGroups.Where(g => ReferenceEquals(g.Picker, picker))) group.Panel.IsExpanded = true;
        picker?.Focus();
    }

    private async Task<bool> NavigateScriptAsync(ScriptEditorPage.Navigation target)
    {
        var accepted = false;
        await PerformAsync(async () =>
        {
            try
            {
                if (!await CommitScriptsAsync()) return;
                if (target.Definition is { } definition)
                {
                    await _scriptPage!.OpenDefinitionAsync(definition, target.Line, target.Column);
                    if (!_scriptDefinitionFiles.Contains(definition)) _scriptDefinitionFiles.Add(definition);
                    _selectedDefinition = definition;
                }
                else if (target.Module is { } module)
                {
                    await _scriptPage!.LocateAsync(module, target.Line, target.Column);
                    _selectedModule = module; _selectedDefinition = null;
                }
                accepted = true;
            }
            finally { RefreshScriptChrome(revealSelected: true); }
        });
        return accepted;
    }

    private async void HandleScriptWorkspaceKey(object? sender, KeyEventArgs e)
    {
        if (e.Handled || e.KeyModifiers != KeyModifiers.None || _consoleInput?.IsComposing == true) return;
        if (e.Key == Key.F5) { e.Handled = true; await RequestScriptApplyAsync(); }
        else if (e.Key == Key.F12) { e.Handled = true; ToggleScriptConsole(); }
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
        var label = visible ? UiText.Get("隐藏控制台 · F12") : UiText.Get("显示控制台 · F12");
        _scriptConsoleToggle.Header = label; AutomationProperties.SetName(_scriptConsoleToggle, label);
    }
}
