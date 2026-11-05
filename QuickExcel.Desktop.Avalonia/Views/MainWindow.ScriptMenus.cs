using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Presentation.Web;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private MenuItem? _scriptRunMenu;
    private Button? _scriptEventButton;
    private TextBlock? _scriptFileTab;

    private Menu CreateScriptMenu()
    {
        var menu = new Menu { Background = Brush.Parse("#181818") };
        MenuItem Group(string label) { var group = new MenuItem { Header = UiText.Get(label) }; menu.Items.Add(group); return group; }
        MenuItem Item(MenuItem parent, string label, Func<Task> action)
        { var item = CommandMenu(UiText.Get(label), action, requiresWorkbook: false, mutation: false); parent.Items.Add(item); return item; }
        var file = Group("文件");
        Item(file, "添加模块", AddScriptModuleAsync); Item(file, "保存", async () => await SaveAsync(false));
        file.Items.Add(new Separator());
        _scriptApiMigration = Item(file, "确认使用 API v4", MigrateScriptApiAsync);
        Item(file, "移除全部脚本", RemoveScriptsAsync);
        _scriptRecovery = Item(file, "恢复已确认草稿", RecoverScriptPageAsync); _scriptRecovery.IsVisible = false;
        var edit = Group("编辑");
        foreach (var (label, command) in new[] { ("撤销", "undo"), ("重做", "redo"), ("补全", "editor.action.triggerSuggest"),
            ("参数提示", "editor.action.triggerParameterHints"), ("类型提示", "editor.action.showHover"), ("转到定义", "editor.action.revealDefinition"),
            ("下一问题 · F8", "editor.action.marker.next"), ("格式化代码", "editor.action.formatDocument") })
            Item(edit, label, () => _scriptPage?.ExecuteEditorCommandAsync(command) ?? Task.CompletedTask);
        var view = Group("视图");
        _scriptConsoleToggle = Item(view, "显示控制台 · F12", () => { ToggleScriptConsole(); return Task.CompletedTask; });
        Item(view, "运行日志", () => { SetScriptOutputVisible(true, tab: 1); return Task.CompletedTask; });
        var run = Group("运行");
        Item(run, "应用并重算", ApplyScriptsAsync);
        Item(run, "停用脚本", () => { DisableScripts(); return Task.CompletedTask; });
        _scriptRunMenu = new MenuItem { Header = UiText.Get("运行操作") }; run.Items.Add(_scriptRunMenu); RefreshScriptRunMenu();
        var security = Group("签名");
        Item(security, "签名详情", ScriptSignatureDetailsAsync); Item(security, "签署脚本", SignScriptsAsync); Item(security, "受信任证书", TrustedCertificatesAsync);
        RefreshScriptApiMigrationButton(); return menu;
    }

    private void RefreshScriptRunMenu()
    {
        if (_scriptRunMenu is null) return;
        _scriptRunMenu.Items.Clear();
        foreach (var action in _availableScriptActions)
        {
            var item = CommandMenu(action.Name, () => RunScriptActionAsync(action));
            if (!string.IsNullOrWhiteSpace(action.Description)) ToolTip.SetTip(item, action.Description);
            _scriptRunMenu.Items.Add(item);
        }
        _scriptRunMenu.IsEnabled = ScriptActionsCurrent && _availableScriptActions.Count > 0 && _scriptCancellation is null;
    }

    private Control CreateScriptEditorHeader()
    {
        var header = new DockPanel { Background = Brush.Parse("#181818"), MinHeight = 36 };
        _scriptEventButton = new Button { Content = UiText.Get("添加事件处理器"), Margin = new(8, 3), FontSize = 12, IsVisible = false };
        AutomationProperties.SetName(_scriptEventButton, UiText.Get("添加事件处理器"));
        _scriptEventButton.Click += async (_, _) => await PerformAsync(AddEventHandlerAsync);
        DockPanel.SetDock(_scriptEventButton, Dock.Right); header.Children.Add(_scriptEventButton);
        _scriptFileTab = new TextBlock { Margin = new(14, 8), VerticalAlignment = VerticalAlignment.Center };
        header.Children.Add(new Border { Background = Brush.Parse("#1E1E1E"), BorderBrush = Brush.Parse("#007ACC"), BorderThickness = new(0, 1, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Left, Child = _scriptFileTab });
        return header;
    }

    private void RefreshScriptEditorHeader()
    {
        var modules = Document.Snapshot.ScriptModules;
        if (_scriptFileTab is not null) _scriptFileTab.Text = _selectedDefinition ?? Path.GetFileName(modules.Modules[_selectedModule].Path);
        if (_scriptEventButton is not null)
        {
            _scriptEventButton.IsVisible = _selectedDefinition is null && modules.Bindings.Any(b => b.Module == _selectedModule && b.Kind == "events");
            _scriptEventButton.IsEnabled = _scriptPage is { Drafts.IsReadOnly: false, Drafts.Failed: false } && _scriptCancellation is null;
        }
        RefreshScriptRunMenu();
    }

    private async Task AddEventHandlerAsync()
    {
        var page = _scriptPage; var document = Document; var id = _selectedModule;
        if (page is null || _selectedDefinition is not null || !document.Snapshot.ScriptModules.Bindings.Any(b => b.Module == id && b.Kind == "events")) return;
        var anchor = await page.CaptureInsertionAsync(id);
        var dialog = new ScriptEventDialog { RequestedThemeVariant = _scriptWindow!.ActualThemeVariant };
        var selected = await ShowOwnedAsync<ScriptEventTemplate?>(dialog, _scriptWindow);
        if (selected is null) return;
        if (!ReferenceEquals(page, _scriptPage) || !ReferenceEquals(document, Document) || id != _selectedModule)
            throw new InvalidOperationException(UiText.Get("脚本或光标所属模块已变化，请重新添加事件处理器。"));
        await page.InsertTemplateAsync(anchor, selected.Code(document.Snapshot.ScriptModules.Modules[id].Language));
    }
}
