using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly TextBox _automationSearch = new() { Name = "AutomationSearch", Height = 32, PlaceholderText = UiText.Get("搜索操作名称或描述") };
    private readonly ComboBox _automationActions = new()
    {
        Name = "AutomationActions", Height = 50, MaxDropDownHeight = 320, PlaceholderText = UiText.Get("请选择操作"),
        HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Stretch
    };
    private readonly TextBlock _automationEmpty = new() { Margin = new(12, 0), VerticalAlignment = VerticalAlignment.Center, TextWrapping = TextWrapping.Wrap };
    private Button? _automationRun, _automationLoad;
    private IReadOnlyList<ScriptActionDescriptor> _availableScriptActions = [];
    private Guid _actionsSession;
    private ScriptModuleSet? _actionsModules;
    private bool AutomationPickerOwnsKeyboard => _automationSearch.IsKeyboardFocusWithin || _automationActions.IsKeyboardFocusWithin || _automationActions.IsDropDownOpen;
    private bool ScriptActionsCurrent => Document.ScriptsEnabled && _actionsSession == Document.Session.SessionId && ReferenceEquals(_actionsModules, Document.Snapshot.ScriptModules);

    private void InstallAutomationRibbon()
    {
        AutomationProperties.SetName(_automationSearch, UiText.Get("搜索脚本操作名称或描述"));
        AutomationProperties.SetName(_automationActions, UiText.Get("可用脚本操作，选择后点击执行"));
        _automationSearch.TextChanged += (_, _) => FilterScriptActions();
        ScrollViewer.SetHorizontalScrollBarVisibility(_automationActions, ScrollBarVisibility.Disabled);
        ScrollViewer.SetVerticalScrollBarVisibility(_automationActions, ScrollBarVisibility.Auto);
        _automationActions.Styles.Add(new Style(s => s.OfType<ComboBoxItem>())
        {
            Setters =
            {
                new Setter(HeightProperty, 44d), new Setter(PaddingProperty, new Thickness(8, 2)),
                new Setter(HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch)
            }
        });
        _automationActions.ItemTemplate = new FuncDataTemplate<ScriptActionDescriptor>((action, _) =>
        {
            if (action is null) return null;
            var item = new StackPanel { Spacing = 1, HorizontalAlignment = HorizontalAlignment.Stretch, Background = Brushes.Transparent };
            item.Children.Add(new TextBlock { Text = action.Name, FontWeight = FontWeight.SemiBold, TextTrimming = TextTrimming.CharacterEllipsis });
            item.Children.Add(new TextBlock { Text = string.IsNullOrWhiteSpace(action.Description) ? UiText.Get("无描述") : action.Description, FontSize = 12, TextTrimming = TextTrimming.CharacterEllipsis });
            ToolTip.SetTip(item, action.Name + "\n" + (string.IsNullOrWhiteSpace(action.Description) ? UiText.Get("无描述") : action.Description));
            AutomationProperties.SetName(item, action.Name + "，" + action.Description);
            return item;
        });
        _automationActions.SelectionChanged += (_, _) => RefreshScriptActionButtons();
        var picker = new global::Avalonia.Controls.Grid { Height = 50 };
        picker.Children.Add(_automationActions); picker.Children.Add(_automationEmpty);
        var body = new StackPanel { Width = 560, Spacing = 4, VerticalAlignment = VerticalAlignment.Center };
        body.Children.Add(_automationSearch); body.Children.Add(picker);
        RibbonContentGroup(UiText.Get("自动执行"), "Actions", body);
        _automationRun = RibbonButton(RibbonIcon.Run, UiText.Get("执行"), () => RunScriptActionAsync(_automationActions.SelectedItem as ScriptActionDescriptor), false);
        _automationRun.Name = "AutomationRun";
        _automationLoad = RibbonButton(RibbonIcon.LoadActions, UiText.Get("加载操作"), ApplyScriptsAsync, false); _automationLoad.Name = "AutomationLoad";
        ToolTip.SetTip(_automationLoad, UiText.Get("应用并授权当前脚本、重算工作簿，然后加载全部操作。操作内容仅在执行时运行。"));
        RibbonGroup(UiText.Get("自动执行"), UiText.Get("操作"), _automationRun, _automationLoad, RibbonButton(RibbonIcon.Code, UiText.Get("脚本编辑器"), OpenScriptsAsync, false));
    }

    private void ClearScriptActions()
    {
        _actionsModules = null; _actionsSession = Guid.Empty; _availableScriptActions = [];
        _automationActions.IsDropDownOpen = false;
        _automationActions.SelectedItem = null; _automationSearch.Text = "";
        _automationActions.ItemsSource = null;
        if (_actionPicker is not null) _actionPicker.ItemsSource = null;
    }

    private void SetScriptActions(WorkbookDocument document, ScriptModuleSet modules, IReadOnlyList<ScriptActionDescriptor> actions)
    {
        if (!ReferenceEquals(document, Document) || !Document.ScriptsEnabled || !ReferenceEquals(modules, Document.Snapshot.ScriptModules))
            throw new InvalidOperationException(UiText.Get("脚本操作列表已过期，请重新加载操作。"));
        _actionsSession = document.Session.SessionId; _actionsModules = modules; _availableScriptActions = actions;
        _automationActions.SelectedItem = null; _automationSearch.Text = "";
        FilterScriptActions();
        if (_actionPicker is not null) { _actionPicker.ItemsSource = actions; _actionPicker.SelectedIndex = actions.Count > 0 ? 0 : -1; }
        RefreshScriptActions();
    }

    private void RefreshScriptActions()
    {
        if (!ScriptActionsCurrent && _actionsModules is not null) ClearScriptActions();
        _automationSearch.IsEnabled = _availableScriptActions.Count > 0;
        _automationEmpty.IsVisible = _automationActions.ItemCount == 0;
        _automationActions.IsVisible = !_automationEmpty.IsVisible;
        _automationEmpty.Text = _availableScriptActions.Count > 0 ? UiText.Get("没有匹配的操作，请更换或清空关键词。") :
            !DesktopServices.SupportsScripts ? UiText.Get("此平台尚未开放脚本执行。") :
            Document.Snapshot.FileState?.FormulaReadOnly == true ? UiText.Get("当前工作簿为只读，无法执行脚本操作。") :
            !Document.Snapshot.ScriptModules.HasActions || ScriptActionsCurrent ? UiText.Get("当前工作簿没有可用的脚本操作。") :
            UiText.Get("点击“加载操作”应用脚本，加载全部可用操作。");
        RefreshScriptActionButtons();
    }

    private void FilterScriptActions()
    {
        var keyword = _automationSearch.Text?.Trim() ?? "";
        var selected = _automationActions.SelectedItem as ScriptActionDescriptor;
        var matches = _availableScriptActions.Where(action => keyword.Length == 0 ||
            action.Name.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
            action.Description?.Contains(keyword, StringComparison.OrdinalIgnoreCase) == true).ToArray();
        _automationActions.ItemsSource = matches;
        // A filtered-out selection must not remain executable, and searching never selects an action implicitly.
        _automationActions.SelectedItem = selected is not null && matches.Contains(selected) ? selected : null;
        if (matches.Length == 0) _automationActions.IsDropDownOpen = false;
        RefreshScriptActions();
    }

    private void RefreshScriptActionButtons()
    {
        var enabled = DesktopServices.SupportsScripts && !NonExecutionBusy && Document.Snapshot.FileState?.FormulaReadOnly != true;
        if (_automationLoad is not null) _automationLoad.IsEnabled = enabled;
        if (_automationRun is not null) _automationRun.IsEnabled = enabled && ScriptActionsCurrent && _automationActions.SelectedItem is ScriptActionDescriptor;
    }
}
