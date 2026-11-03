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
    private readonly ListBox _automationActions = new() { Name = "AutomationActions", Width = 560, Height = 90, SelectionMode = SelectionMode.Single };
    private readonly TextBlock _automationEmpty = new() { Margin = new(12), VerticalAlignment = VerticalAlignment.Center, TextWrapping = TextWrapping.Wrap };
    private Button? _automationRun, _automationLoad;
    private IReadOnlyList<ScriptActionDescriptor> _availableScriptActions = [];
    private Guid _actionsSession;
    private ScriptModuleSet? _actionsModules;
    private bool ScriptActionsCurrent => Document.ScriptsEnabled && _actionsSession == Document.Session.SessionId && ReferenceEquals(_actionsModules, Document.Snapshot.ScriptModules);

    private void InstallAutomationRibbon()
    {
        AutomationProperties.SetName(_automationActions, "可用脚本操作，单击选中，双击执行");
        ScrollViewer.SetHorizontalScrollBarVisibility(_automationActions, ScrollBarVisibility.Disabled);
        ScrollViewer.SetVerticalScrollBarVisibility(_automationActions, ScrollBarVisibility.Auto);
        _automationActions.Styles.Add(new Style(s => s.OfType<ListBoxItem>())
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
            item.Children.Add(new TextBlock { Text = string.IsNullOrWhiteSpace(action.Description) ? "无描述" : action.Description, FontSize = 12, TextTrimming = TextTrimming.CharacterEllipsis });
            ToolTip.SetTip(item, action.Name + "\n" + (string.IsNullOrWhiteSpace(action.Description) ? "无描述" : action.Description));
            AutomationProperties.SetName(item, action.Name + "，" + action.Description);
            // Handle the item only: double-clicking empty list space must not rerun a previous selection.
            item.DoubleTapped += async (_, e) =>
            {
                e.Handled = true; _automationActions.SelectedItem = action;
                await PerformAsync(() => RunScriptActionAsync(action));
            };
            return item;
        });
        _automationActions.SelectionChanged += (_, _) => RefreshScriptActionButtons();
        var body = new global::Avalonia.Controls.Grid { Width = 560 };
        body.Children.Add(_automationActions); body.Children.Add(_automationEmpty);
        RibbonContentGroup("自动执行", "脚本操作 · 单击选中，双击执行", body);
        _automationRun = RibbonButton("执行", () => RunScriptActionAsync(_automationActions.SelectedItem as ScriptActionDescriptor), false);
        _automationRun.Name = "AutomationRun";
        _automationLoad = RibbonButton("加载操作", ApplyScriptsAsync, false); _automationLoad.Name = "AutomationLoad";
        ToolTip.SetTip(_automationLoad, "应用并授权当前脚本、重算工作簿，然后加载全部操作。操作内容仅在执行时运行。");
        RibbonGroup("自动执行", "操作", _automationRun, _automationLoad, RibbonButton("脚本编辑器", OpenScriptsAsync, false));
    }

    private void ClearScriptActions()
    {
        _actionsModules = null; _actionsSession = Guid.Empty; _availableScriptActions = [];
        _automationActions.ItemsSource = null;
        if (_actionPicker is not null) _actionPicker.ItemsSource = null;
    }

    private void SetScriptActions(WorkbookDocument document, ScriptModuleSet modules, IReadOnlyList<ScriptActionDescriptor> actions)
    {
        if (!ReferenceEquals(document, Document) || !Document.ScriptsEnabled || !ReferenceEquals(modules, Document.Snapshot.ScriptModules))
            throw new InvalidOperationException("脚本操作列表已过期，请重新加载操作。");
        _actionsSession = document.Session.SessionId; _actionsModules = modules; _availableScriptActions = actions;
        _automationActions.ItemsSource = actions; _automationActions.SelectedIndex = -1;
        if (_actionPicker is not null) { _actionPicker.ItemsSource = actions; _actionPicker.SelectedIndex = actions.Count > 0 ? 0 : -1; }
        RefreshScriptActions();
    }

    private void RefreshScriptActions()
    {
        if (!ScriptActionsCurrent && _actionsModules is not null) ClearScriptActions();
        _automationEmpty.IsVisible = _availableScriptActions.Count == 0;
        _automationEmpty.Text = !DesktopServices.SupportsScripts ? "此平台尚未开放脚本执行。" :
            Document.Snapshot.FileState?.FormulaReadOnly == true ? "当前工作簿为只读，无法执行脚本操作。" :
            !Document.Snapshot.ScriptModules.HasActions || ScriptActionsCurrent ? "当前工作簿没有可用的脚本操作。" :
            "点击“加载操作”应用脚本，加载全部可用操作。";
        RefreshScriptActionButtons();
    }

    private void RefreshScriptActionButtons()
    {
        var enabled = DesktopServices.SupportsScripts && !_workspace.IsBusy && _scriptCancellation is null && Document.Snapshot.FileState?.FormulaReadOnly != true;
        if (_automationLoad is not null) _automationLoad.IsEnabled = enabled;
        if (_automationRun is not null) _automationRun.IsEnabled = enabled && ScriptActionsCurrent && _automationActions.SelectedItem is ScriptActionDescriptor;
    }
}
