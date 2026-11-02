using System.Text.RegularExpressions;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Desktop.Avalonia.Web;
using QuickExcel.Infrastructure.Web;
using QuickExcel.Presentation.Web;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly string _userDataDirectory;
    private WebAssetCatalog? _webAssets;
    private ScriptEditorPage? _scriptPage;
    private Window? _scriptWindow;
    private ContentControl? _scriptContent;
    private ComboBox? _modulePicker, _actionPicker;
    private ListBox? _scriptDiagnostics;
    private TextBlock? _scriptStatus;
    private Button? _scriptRecovery;
    private Guid _scriptDocument;
    private long _webGeneration;
    private string _selectedModule = "main";
    private bool _scriptClosing, _scriptPicking;
    private CancellationTokenSource? _scriptCancellation;
    private sealed record ScriptDiagnostic(string Label, long Revision, string? Module = null, int Line = 1, SheetId? Sheet = null, CellPosition? Cell = null);
    private WebAssetCatalog WebAssets => _webAssets ??= new(Path.Combine(AppContext.BaseDirectory, "Web"), allowTestResources: Program.WebTestMode);
    private void InstallScriptCommands()
    {
        var menu = new MenuItem { Header = "脚本" };
        menu.Items.Add(CommandMenu("脚本编辑器…", OpenScriptsAsync));
        menu.Items.Add(CommandMenu("应用脚本并重算", ApplyScriptsAsync));
        menu.Items.Add(CommandMenu("停用脚本", () => { DisableScripts(); return Task.CompletedTask; }));
        menu.Items.Add(CommandMenu("签名详情…", ScriptSignatureDetailsAsync));
        menu.Items.Add(CommandMenu("签署脚本…", SignScriptsAsync));
        menu.Items.Add(CommandMenu("受信任证书…", TrustedCertificatesAsync)); _menu.Items.Add(menu);
        RibbonGroup("公式", "脚本", RibbonButton("脚本编辑器…", OpenScriptsAsync, false), RibbonButton("应用并重算", ApplyScriptsAsync, false));
    }
    private async Task OpenScriptsAsync()
    {
        if (_scriptWindow is not null) { _scriptWindow.Activate(); return; }
        if (!await CommitEditorsAsync()) return;
        _scriptDocument = Document.Session.SessionId; _selectedModule = Document.Snapshot.ScriptModules.EntryModules[0];
        var window = _scriptWindow = new Window { Title = "脚本工作区 · QuickExcel", Width = 1080, Height = 760, MinWidth = 760, MinHeight = 500, WindowStartupLocation = WindowStartupLocation.CenterOwner };
        var root = new DockPanel(); var toolbar = new WrapPanel { Margin = new(8), Orientation = Orientation.Horizontal };
        _modulePicker = new ComboBox { MinWidth = 170, Margin = new(4) }; AutomationProperties.SetName(_modulePicker, "脚本模块"); toolbar.Children.Add(_modulePicker);
        void Button(string text, Func<Task> action)
        { var button = new Button { Content = text, Margin = new(4) }; button.Click += async (_, _) => await PerformAsync(action); toolbar.Children.Add(button); }
        Button("添加模块…", AddScriptModuleAsync); Button("模块角色…", ScriptModuleRoleAsync); Button("应用并重算", ApplyScriptsAsync);
        Button("停用", () => { DisableScripts(); return Task.CompletedTask; }); Button("保存", async () => await SaveAsync(false));
        _actionPicker = new ComboBox { MinWidth = 140, Margin = new(4), ItemTemplate = new FuncDataTemplate<ScriptActionDescriptor>((item, _) => new TextBlock { Text = item?.Name }) };
        AutomationProperties.SetName(_actionPicker, "脚本操作"); toolbar.Children.Add(_actionPicker); Button("运行操作", RunScriptActionAsync);
        var stop = new Button { Content = "停止", Margin = new(4) }; stop.Click += (_, _) => _scriptCancellation?.Cancel(); toolbar.Children.Add(stop);
        Button("签名…", ScriptSignatureDetailsAsync);
        _scriptRecovery = new Button { Content = "恢复已确认草稿…", IsVisible = false, Margin = new(4) };
        _scriptRecovery.Click += async (_, _) => await PerformAsync(RecoverScriptPageAsync); toolbar.Children.Add(_scriptRecovery);
        DockPanel.SetDock(toolbar, Dock.Top); root.Children.Add(toolbar);
        _scriptStatus = new TextBlock { Margin = new(12, 5), TextWrapping = TextWrapping.Wrap };
        DockPanel.SetDock(_scriptStatus, Dock.Bottom); root.Children.Add(_scriptStatus);
        _scriptDiagnostics = new ListBox { MaxHeight = 150, MinHeight = 32, ItemTemplate = new FuncDataTemplate<ScriptDiagnostic>((item, _) => new TextBlock { Text = item?.Label, TextWrapping = TextWrapping.Wrap }) };
        AutomationProperties.SetName(_scriptDiagnostics, "运行日志与错误；静态错误在编辑器内显示");
        _scriptDiagnostics.DoubleTapped += async (_, _) => await PerformAsync(ActivateScriptDiagnosticAsync);
        _scriptDiagnostics.KeyDown += async (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; await PerformAsync(ActivateScriptDiagnosticAsync); } };
        DockPanel.SetDock(_scriptDiagnostics, Dock.Bottom); root.Children.Add(_scriptDiagnostics);
        _scriptContent = new ContentControl(); root.Children.Add(_scriptContent); window.Content = root;
        _modulePicker.SelectionChanged += async (_, _) =>
        {
            if (_scriptPicking || _modulePicker.SelectedItem is not string id || id == _selectedModule) return;
            await PerformAsync(async () =>
            {
                try { if (await CommitScriptsAsync()) { await _scriptPage!.SelectAsync(id); _selectedModule = id; } }
                finally { RefreshScriptChrome(); }
            });
            RefreshScriptChrome();
        };
        window.Closing += async (_, e) =>
        {
            if (_scriptClosing) return; e.Cancel = true;
            await PerformAsync(async () =>
            {
                if (_scriptPage is null || _scriptPage.Drafts.Failed)
                {
                    if (await Dialogs.Confirm(this, "关闭失败的编辑器", "放弃尚未提交到工作簿的脚本草稿并关闭编辑器？已提交的文档更改仍然保留。可选择取消，然后恢复已确认草稿。")) await CloseScriptWindowAsync();
                }
                else if (await CommitScriptsAsync()) await CloseScriptWindowAsync();
            });
        };
        window.KeyDown += (_, e) => { if (e.Key == Key.Escape && _scriptCancellation is not null) { _scriptCancellation.Cancel(); e.Handled = true; } };
        window.Show(this);
        try { await CreateScriptPageAsync(new(Document.Snapshot.ScriptModules, Document.Snapshot.FileState?.FormulaReadOnly == true)); }
        catch (Exception ex) { ScriptPageFailed(ex); throw; }
    }
    private async Task CreateScriptPageAsync(ScriptDraftSession drafts)
    {
        var page = _scriptPage = new ScriptEditorPage(WebAssets, drafts, _scriptDocument.ToString("N"), ++_webGeneration, _userDataDirectory);
        page.Failed += ScriptPageFailed;
        page.DraftChanged += () => { DisableScripts(); UpdateTitle(); if (_scriptWindow is not null) _scriptWindow.Title = "* 脚本工作区 · QuickExcel"; if (_scriptStatus is not null) _scriptStatus.Text = "源码草稿已修改，脚本授权已撤销。"; };
        page.SaveRequested += () => _ = PerformAsync(async () => await SaveAsync(false));
        page.FocusRequested += () => _modulePicker?.Focus();
        _scriptContent!.Content = page.Browser.View; RefreshScriptChrome();
        await page.LoadAsync(_selectedModule); RefreshScriptChrome();
    }
    private void ScriptPageFailed(Exception error)
    {
        DisableScripts(); if (_scriptContent is not null) _scriptContent.IsEnabled = false;
        if (_scriptRecovery is not null) _scriptRecovery.IsVisible = true;
        if (_scriptStatus is not null) _scriptStatus.Text = "页面已停止，保存和运行已阻止。已确认草稿仍在主程序内。" + error.Message;
        Error(error);
    }
    private async Task RecoverScriptPageAsync()
    {
        if (!await Dialogs.Confirm(this, "恢复脚本页面", "重新加载主程序已确认的草稿。浏览器尚未确认的最后输入无法恢复；请核对源码后再保存或运行。")) return;
        var drafts = _scriptPage?.Drafts ?? new(Document.Snapshot.ScriptModules, Document.Snapshot.FileState?.FormulaReadOnly == true);
        if (_scriptPage is not null) await _scriptPage.DisposeAsync();
        drafts.RecoverConfirmed(); _scriptPage = null; _scriptContent!.IsEnabled = true; _scriptRecovery!.IsVisible = false;
        try { await CreateScriptPageAsync(drafts); } catch (Exception ex) { ScriptPageFailed(ex); }
    }
    private async Task<bool> CommitScriptsAsync()
    {
        if (_scriptWindow is null) return true;
        try
        {
            if (_scriptPage is null || _scriptDocument != Document.Session.SessionId) throw new InvalidOperationException("脚本页面未加载或所属文档已变化。");
            await _scriptPage.FlushAsync(() => Document.Snapshot.ScriptModules, modules =>
            {
                var current = Document.Snapshot.ScriptModules;
                if (modules.Modules.Any(p => current.Modules[p.Key].Source != p.Value.Source))
                    _workspace.Editor.Execute(new UpdateScriptModulesCommand(modules));
                Synchronize(); return Task.CompletedTask;
            });
            RefreshScriptChrome(); return true;
        }
        catch (Exception ex) { Error(ex); if (_scriptStatus is not null) _scriptStatus.Text = ex.Message; return false; }
    }
    private async Task CloseScriptWindowAsync()
    {
        if (_scriptWindow is null) return;
        var window = _scriptWindow; _scriptClosing = true;
        try { if (_scriptPage is not null) await _scriptPage.DisposeAsync(); window.Close(); }
        finally { _scriptWindow = null; _scriptPage = null; _scriptContent = null; _modulePicker = null; _actionPicker = null; _scriptStatus = null; _scriptDiagnostics = null; _scriptRecovery = null; _scriptClosing = false; }
    }
    private void RefreshScriptChrome()
    {
        if (_scriptWindow is null) return;
        _scriptWindow.Title = (_scriptPage?.Drafts.IsDirty == true ? "* " : "") + "脚本工作区 · QuickExcel";
        _scriptPicking = true;
        try { _modulePicker!.ItemsSource = Document.Snapshot.ScriptModules.EntryModules.Concat(Document.Snapshot.ScriptModules.Modules.Keys.Order()).Distinct().ToArray(); _modulePicker.SelectedItem = _selectedModule; }
        finally { _scriptPicking = false; }
        var modules = Document.Snapshot.ScriptModules;
        if (_scriptStatus is not null && _scriptPage?.Drafts.Failed != true)
            _scriptStatus.Text = $"{modules.Modules[_selectedModule].Path} · API v{modules.ScriptApiVersion} · {(Document.ScriptsEnabled ? "已授权" : "未授权")} · {(modules.Signature is null ? "未签名" : "已签名")} · " + (OperatingSystem.IsWindows() ? "独立隔离进程执行" : "此平台尚未开放脚本执行");
    }
    private void DisableScripts()
    { _scriptCancellation?.Cancel(); Document.DisableScripts(); if (_actionPicker is not null) _actionPicker.ItemsSource = null; RefreshScriptChrome(); }
    private async Task AddScriptModuleAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var old = Document.Snapshot.ScriptModules;
        if (old.ScriptApiVersion != 2) throw new InvalidOperationException("API v1 工作簿保持原有执行语义；添加模块需要 API v2 工作簿。");
        var id = await Dialogs.Text(this, "新建模块", "模块逻辑 ID", "helpers"); if (id is null) return;
        if (old.Modules.ContainsKey(id)) throw new ArgumentException("模块 ID 已存在。");
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(old.Modules.Add(id, new($"scripts/{id}.js", "export function example(value) {\n    return value;\n}\n")), old.EntryModules, old.Revision, old.BindingRevision, 2, old.Bindings)));
        // The module set changed; start a new page generation after the existing page has acknowledged every draft.
        if (_scriptPage is not null) await _scriptPage.DisposeAsync(); _selectedModule = id;
        await CreateScriptPageAsync(new(Document.Snapshot.ScriptModules)); Synchronize();
    }
    private async Task ScriptModuleRoleAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var old = Document.Snapshot.ScriptModules;
        if (old.ScriptApiVersion != 2) throw new InvalidOperationException("API v1 不支持更改模块角色。");
        var role = (old.Bindings.Any(b => b.Module == _selectedModule && b.Kind == "formula-functions") ? 1 : 0) + (old.Bindings.Any(b => b.Module == _selectedModule && b.Kind == "actions") ? 2 : 0);
        var dialog = new FieldsDialog("模块角色").Choice("role", _selectedModule, ["辅助模块", "公式函数入口", "操作入口", "公式和操作入口"], role);
        if (!await dialog.ShowAsync(this) || !await CommitScriptsAsync()) return;
        old = Document.Snapshot.ScriptModules; role = dialog.Index("role");
        var bindings = old.Bindings.Where(b => b.Module != _selectedModule).ToList();
        if ((role & 1) != 0) bindings.Add(new(_selectedModule, "formula-functions")); if ((role & 2) != 0) bindings.Add(new(_selectedModule, "actions"));
        var entries = old.EntryModules.Where(id => bindings.Any(b => b.Module == id)).ToList();
        if (role != 0 && !entries.Contains(_selectedModule)) entries.Add(_selectedModule);
        if (entries.Count == 0) throw new InvalidOperationException("至少保留一个公式或操作入口。");
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(old.Modules, entries, old.Revision, old.BindingRevision, old.ScriptApiVersion, bindings)));
        DisableScripts(); Synchronize();
    }
    private async Task ApplyScriptsAsync()
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync()) return;
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("此平台的脚本隔离尚未开放，不能执行脚本。");
        if (!await CheckScriptTrustAsync()) return;
        await ScriptWorkAsync(async token =>
        {
            var document = Document; await Task.Run(() => document.EnableScriptsAndCalculate(token), token);
            var actions = document.Snapshot.ScriptModules.HasActions ? await Task.Run(() => document.ListActions(token), token) : [];
            if (_actionPicker is not null) { _actionPicker.ItemsSource = actions; _actionPicker.SelectedIndex = actions.Count > 0 ? 0 : -1; }
        });
    }
    private async Task RunScriptActionAsync()
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync()) return;
        if (!Document.ScriptsEnabled || _actionPicker?.SelectedItem is not ScriptActionDescriptor action) throw new InvalidOperationException("请先应用脚本并选择操作。");
        await ScriptWorkAsync(token => Task.Run(() => Document.RunAction(action.Name, token), token));
    }
    private async Task ScriptWorkAsync(Func<CancellationToken, Task> work)
    {
        _calculationTimer.Stop(); using var cancellation = new CancellationTokenSource(); _scriptCancellation = cancellation;
        _book.IsEnabled = false;
        try { await work(cancellation.Token); _calculationPending = false; }
        finally { _scriptCancellation = null; _book.IsEnabled = true; Synchronize(); RefreshScriptChrome(); RefreshScriptDiagnostics(); }
    }
    private void RefreshScriptDiagnostics()
    {
        if (_scriptDiagnostics is null) return;
        var modules = Document.Snapshot.ScriptModules; var entries = new List<ScriptDiagnostic>();
        foreach (var error in Document.ScriptErrors.Take(100))
        {
            var module = modules.Modules.FirstOrDefault(m => error.Contains(m.Value.Path, StringComparison.Ordinal) || error.Contains(m.Key + ":", StringComparison.Ordinal));
            var names = module.Key is not null ? "(?:" + Regex.Escape(module.Value.Path) + "|" + Regex.Escape(module.Key) + ")" : "(?:<quickexcel>|quickexcel|<eval>|script)";
            var match = Regex.Match(error, names + @":(\d+)(?::\d+)?(?:\)|\s|$)");
            entries.Add(new(error, modules.Revision, module.Key ?? modules.EntryModules[0], match.Success && int.TryParse(match.Groups[1].Value, out var line) ? line : 1));
        }
        foreach (var sheet in Document.Snapshot.Sheets)
        foreach (var (position, cell) in sheet.Cells.Where(p => p.Value.Value.Kind == CellValueKind.Error).Take(200 - entries.Count))
            entries.Add(new($"{sheet.Name}!{WorkbookGrid.ColumnName(position.Column)}{position.Row + 1}: {cell.Value}", modules.Revision, Sheet: sheet.Id, Cell: position));
        _scriptDiagnostics.ItemsSource = entries;
    }
    private async Task ActivateScriptDiagnosticAsync()
    {
        if (_scriptDiagnostics?.SelectedItem is not ScriptDiagnostic entry || !await CommitEditorsAsync()) return;
        if (entry.Revision != Document.Snapshot.ScriptModules.Revision) throw new InvalidOperationException("此运行诊断属于旧源码，请重新运行后定位。");
        if (entry.Sheet is { } sheet && entry.Cell is { } cell) { Document.Session.UpdateNavigation(sheet, cell); Synchronize(); ShowBook(); Activate(); _grid.Select(cell); }
        else if (entry.Module is { } module && _scriptPage is not null) { await _scriptPage.LocateAsync(module, entry.Line, 1); _selectedModule = module; RefreshScriptChrome(); }
    }
}
