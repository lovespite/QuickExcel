using System.Text.RegularExpressions;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Desktop.Avalonia.Web;
using QuickExcel.Infrastructure.Web;
using QuickExcel.Presentation.Web;
using QuickExcel.Runtime;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly string _userDataDirectory;
    private WebAssetCatalog? _webAssets;
    private ScriptEditorPage? _scriptPage;
    private Window? _scriptWindow;
    private ContentControl? _scriptContent;
    private ListBox? _modulePicker;
    private ComboBox? _actionPicker;
    private ListBox? _scriptDiagnostics;
    private TextBlock? _scriptStatus;
    private Button? _scriptRecovery;
    private Guid _scriptDocument;
    private long _webGeneration;
    private string _selectedModule = "main";
    private bool _scriptClosing, _scriptPicking;
    private CancellationTokenSource? _scriptCancellation;
    private sealed record ScriptDiagnostic(string Label, long Revision, string? Module = null, int Line = 1, SheetId? Sheet = null, CellPosition? Cell = null, int Column = 1, long? DraftVersion = null);
    private WebAssetCatalog WebAssets => _webAssets ??= new(Path.Combine(AppContext.BaseDirectory, "Web"), allowTestResources: Program.WebTestMode);
    private void InstallScriptCommands()
    {
        var menu = new MenuItem { Header = "脚本" };
        menu.Items.Add(CommandMenu("脚本编辑器", OpenScriptsAsync));
        menu.Items.Add(CommandMenu("调试控制台", OpenScriptConsoleAsync));
        menu.Items.Add(CommandMenu("应用脚本并重算", ApplyScriptsAsync));
        menu.Items.Add(CommandMenu("停用脚本", () => { DisableScripts(); return Task.CompletedTask; }));
        menu.Items.Add(CommandMenu("签名详情", ScriptSignatureDetailsAsync));
        menu.Items.Add(CommandMenu("签署脚本", SignScriptsAsync));
        menu.Items.Add(CommandMenu("受信任证书", TrustedCertificatesAsync)); _menu.Items.Add(menu);
        RibbonGroup("公式", "脚本", RibbonButton("脚本编辑器", OpenScriptsAsync, false), RibbonButton("应用并重算", ApplyScriptsAsync, false));
        InstallAutomationRibbon();
    }
    private async Task OpenScriptsAsync()
    {
        if (_scriptWindow is not null) { _scriptWindow.Activate(); return; }
        if (!await CommitEditorsAsync()) return;
        _scriptDocument = Document.Session.SessionId; _selectedModule = Document.Snapshot.ScriptModules.EntryModules[0];
        var window = _scriptWindow = new Window { Title = "脚本工作区 · QuickExcel", Width = 1080, Height = 760, MinWidth = 760, MinHeight = 500, WindowStartupLocation = WindowStartupLocation.CenterOwner };
        var root = new DockPanel(); var toolbar = new WrapPanel { Margin = new(8), Orientation = Orientation.Horizontal };
        void Button(string text, Func<Task> action)
        { var button = new Button { Content = text, Margin = new(4) }; button.Click += async (_, _) => await PerformAsync(action); toolbar.Children.Add(button); }
        Button("添加模块", AddScriptModuleAsync); Button("模块角色", ScriptModuleRoleAsync); Button("应用并重算", ApplyScriptsAsync);
        Button("JS / TS", ChangeScriptLanguageAsync);
        Button("停用", () => { DisableScripts(); return Task.CompletedTask; }); Button("保存", async () => await SaveAsync(false));
        _actionPicker = new ComboBox { MinWidth = 140, Margin = new(4), ItemTemplate = new FuncDataTemplate<ScriptActionDescriptor>((item, _) => new TextBlock { Text = item?.Name }) };
        AutomationProperties.SetName(_actionPicker, "脚本操作"); toolbar.Children.Add(_actionPicker); Button("运行操作", RunScriptActionAsync);
        RefreshScriptActions(); _actionPicker.ItemsSource = _availableScriptActions; _actionPicker.SelectedIndex = _availableScriptActions.Count > 0 ? 0 : -1;
        Button("签名", ScriptSignatureDetailsAsync);
        _scriptConsoleToggle = new Button { Margin = new(4) };
        _scriptConsoleToggle.Click += (_, _) => ToggleScriptConsole(); toolbar.Children.Add(_scriptConsoleToggle);
        _scriptRecovery = new Button { Content = "恢复已确认草稿", IsVisible = false, Margin = new(4) };
        _scriptRecovery.Click += async (_, _) => await PerformAsync(RecoverScriptPageAsync); toolbar.Children.Add(_scriptRecovery);
        DockPanel.SetDock(toolbar, Dock.Top); root.Children.Add(toolbar);
        _scriptStatus = new TextBlock { Margin = new(12, 5), TextWrapping = TextWrapping.Wrap };
        DockPanel.SetDock(_scriptStatus, Dock.Bottom); root.Children.Add(_scriptStatus);
        _scriptDiagnostics = new ListBox { ItemTemplate = new FuncDataTemplate<ScriptDiagnostic>((item, _) => new TextBlock { Text = item?.Label, TextWrapping = TextWrapping.Wrap }) };
        AutomationProperties.SetName(_scriptDiagnostics, "运行日志与错误；静态错误在编辑器内显示");
        _scriptDiagnostics.DoubleTapped += async (_, _) => await PerformAsync(ActivateScriptDiagnosticAsync);
        _scriptDiagnostics.KeyDown += async (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; await PerformAsync(ActivateScriptDiagnosticAsync); } };
        root.Children.Add(CreateScriptWorkspace(CreateScriptOutput()));
        var scriptContent = _scriptContent!;
        window.Content = _scriptExecution = new ExecutionOverlay(window, root, () => _consoleInput?.IsComposing == true,
            // Native WebViews occupy their own surface, above Avalonia overlays. Hide only after the delay.
            visible => scriptContent.IsVisible = !visible);
        window.AddHandler(KeyDownEvent, HandleScriptWorkspaceKey, RoutingStrategies.Tunnel);
        _modulePicker!.SelectionChanged += async (_, _) =>
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
        page.ConsoleRequested += ToggleScriptConsole;
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
    private async Task<bool> CommitScriptsAsync(bool validateTypeScript = false)
    {
        if (_scriptWindow is null) return true;
        try
        {
            if (_scriptPage is null || _scriptDocument != Document.Session.SessionId) throw new InvalidOperationException("脚本页面未加载或所属文档已变化。");
            await _scriptPage.FlushAsync(() => Document.Snapshot.ScriptModules, async modules =>
            {
                var document = Document;
                var current = Document.Snapshot.ScriptModules;
                if (validateTypeScript && modules.HasTypeScript)
                {
                    if (_scriptCancellation is not null) throw new InvalidOperationException("已有脚本工作正在进行。");
                    using var cancellation = new CancellationTokenSource();
                    using var feedback = BeginExecution("正在检查 TypeScript…", cancellation);
                    _scriptCancellation = cancellation;
                    var revision = document.Snapshot.Revision;
                    if (_scriptStatus is not null) _scriptStatus.Text = "正在检查 TypeScript；按 Esc 取消…";
                    try
                    {
                        await ScriptCompilation.PrepareAsync(modules, new(document.Session.SessionId, revision, 0, modules.Revision), cancellation.Token);
                        cancellation.Token.ThrowIfCancellationRequested();
                        if (!ReferenceEquals(document, Document) || document.Snapshot.Revision != revision || !ReferenceEquals(current, document.Snapshot.ScriptModules))
                            throw new InvalidOperationException("编译期间文档已变化，未应用过期产物。");
                    }
                    finally { _scriptCancellation = null; }
                }
                if (modules.Modules.Any(p => current.Modules[p.Key].Source != p.Value.Source))
                    _workspace.Editor.Execute(new UpdateScriptModulesCommand(modules));
                Synchronize(); await Task.CompletedTask;
            });
            RefreshScriptChrome(); return true;
        }
        catch (ScriptCompilationException ex)
        {
            SetScriptOutputVisible(true, tab: 1);
            if (_scriptStatus is not null) _scriptStatus.Text = ex.Message;
            if (_scriptDiagnostics is not null) _scriptDiagnostics.ItemsSource = ex.Diagnostics.Select(d => new ScriptDiagnostic(
                $"{d.Module}:{d.Line}:{d.Column} · {d.Message}", Document.Snapshot.ScriptModules.Revision, d.Module, d.Line ?? 1,
                Column: d.Column ?? 1, DraftVersion: d.Module is { } id && _scriptPage!.Drafts.Models.TryGetValue(id, out var model) ? model.Version : null)).ToArray();
            return false;
        }
        catch (OperationCanceledException) { if (_scriptStatus is not null) _scriptStatus.Text = "编译已取消，草稿未应用。"; return false; }
        catch (Exception ex) { Error(ex); if (_scriptStatus is not null) _scriptStatus.Text = ex.Message; return false; }
    }
    private async Task CloseScriptWindowAsync()
    {
        if (_scriptWindow is null) return;
        _scriptCancellation?.Cancel();
        var window = _scriptWindow; _scriptClosing = true;
        try { if (_scriptPage is not null) await _scriptPage.DisposeAsync(); window.Close(); }
        finally { _scriptWindow = null; _scriptPage = null; _scriptContent = null; _modulePicker = null; _actionPicker = null; _scriptStatus = null; _scriptDiagnostics = null; _scriptRecovery = null; _scriptClosing = false;
            _scriptExecution = null;
            _scriptEditorPane = null; _scriptOutputSplitter = null; _scriptConsoleToggle = null;
            _consoleInput = null; _consoleOutput = null; _consoleRun = null; _consoleCopy = null;
            _consoleFilter = null; _consoleLevel = null; _consoleCount = null; _consoleEmpty = null;
            _consoleVisibleEntries = []; _scriptOutputTabs = null; }
    }
    private async Task RefreshScriptPageAfterMutationAsync()
    {
        if (_scriptPage is not { } page || _scriptWindow is null) return;
        var modules = Document.Snapshot.ScriptModules;
        if (page.Drafts.Models.Count == modules.Modules.Count && modules.Modules.All(p => page.Drafts.Models.TryGetValue(p.Key, out var draft) &&
            draft.Path == p.Value.Path && draft.Language == p.Value.Language && draft.Committed == p.Value.Source)) return;
        if (page.Drafts.IsDirty || page.Drafts.Failed)
            throw new InvalidOperationException("包内源码与编辑器草稿冲突；已保留草稿，不能覆盖新的包源码。");
        await page.DisposeAsync(); _scriptPage = null;
        if (!modules.Modules.ContainsKey(_selectedModule)) _selectedModule = modules.EntryModules[0];
        try { await CreateScriptPageAsync(new(modules, Document.Snapshot.FileState?.FormulaReadOnly == true)); }
        catch (Exception error) { ScriptPageFailed(error); throw; }
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
            _scriptStatus.Text = $"{modules.Modules[_selectedModule].Path} · API v{modules.ScriptApiVersion} · {(Document.ScriptsEnabled ? "已授权" : "未授权")} · {(modules.Signature is null ? "未签名" : "已签名")} · " + (DesktopServices.SupportsScripts ? "独立隔离进程执行" : "此平台尚未开放脚本执行");
    }
    private void DisableScripts()
    { _scriptCancellation?.Cancel(); Document.DisableScripts(); ClearScriptActions(); RefreshScriptActions(); RefreshScriptChrome(); }
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
        var dialog = new FieldsDialog("模块角色").Choice("role", _selectedModule, ["辅助模块", "公式函数入口", "操作与事件入口", "公式、操作与事件入口"], role);
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
        if (!await CommitEditorsAsync(validateTypeScript: true)) return;
        if (!DesktopServices.SupportsScripts) throw new PlatformNotSupportedException("此平台的脚本隔离尚未开放，不能执行脚本。");
        ClearScriptActions(); RefreshScriptActions();
        if (!await CheckScriptTrustAsync()) return;
        await ScriptWorkAsync(async token =>
        {
            var document = Document; var modules = document.Snapshot.ScriptModules;
            await Task.Run(() => document.EnableScriptsAndCalculate(token), token);
            var actions = document.Snapshot.ScriptModules.HasActions ? await Task.Run(() => document.ListActions(token), token) : [];
            SetScriptActions(document, modules, actions);
            _status.Text = $"已加载 {actions.Count} 个脚本操作。";
        });
        await DispatchEventsAsync();
    }
    private Task RunScriptActionAsync() => RunScriptActionAsync(_actionPicker?.SelectedItem as ScriptActionDescriptor);
    private async Task RunScriptActionAsync(ScriptActionDescriptor? action)
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync(validateTypeScript: true)) return;
        if (!ScriptActionsCurrent || action is null || !_availableScriptActions.Contains(action)) throw new InvalidOperationException("请先加载脚本并选择操作。");
        var document = Document;
        _status.Text = "正在执行：" + action.Name;
        await ScriptWorkAsync(async token =>
        {
            var result = await Task.Run(() => document.RunAction(action.Name, token), token);
            _status.Text = $"已执行：{action.Name}" + (result.Changed ? "（可撤销）" : "（无变更）");
        });
    }
    private async Task ScriptWorkAsync(Func<CancellationToken, Task> work)
    {
        _calculationTimer.Stop(); using var cancellation = new CancellationTokenSource();
        using var feedback = BeginExecution("正在执行代码…", cancellation);
        _scriptCancellation = cancellation;
        RefreshScriptActionButtons(); RefreshScriptConsole();
        try { await work(cancellation.Token); _calculationPending = false; }
        finally
        {
            _scriptCancellation = null;
            await RefreshScriptPageAfterMutationAsync();
            Synchronize(); RefreshScriptChrome(); RefreshScriptDiagnostics();
        }
    }
    private void RefreshScriptDiagnostics()
    {
        RefreshScriptConsole();
        if (_scriptDiagnostics is null) return;
        var modules = Document.Snapshot.ScriptModules; var entries = new List<ScriptDiagnostic>();
        foreach (var error in Document.ScriptErrors.Take(100))
        {
            var module = modules.Modules.FirstOrDefault(m => error.Contains(m.Value.Path, StringComparison.Ordinal) || error.Contains(m.Key + ":", StringComparison.Ordinal));
            var names = module.Key is not null ? "(?:" + Regex.Escape(module.Value.Path) + "|" + Regex.Escape(module.Key) + ")" : "(?:<quickexcel>|quickexcel|<eval>|script)";
            var match = Regex.Match(error, names + @":(\d+)(?::(\d+))?(?:\)|\s|$)");
            entries.Add(new(error, modules.Revision, module.Key ?? modules.EntryModules[0], match.Success && int.TryParse(match.Groups[1].Value, out var line) ? line : 1,
                Column: match.Success && int.TryParse(match.Groups[2].Value, out var column) ? column : 1));
        }
        foreach (var sheet in Document.Snapshot.Sheets)
        foreach (var (position, cell) in sheet.Cells.Where(p => p.Value.Value.Kind == CellValueKind.Error).Take(200 - entries.Count))
            entries.Add(new($"{sheet.Name}!{WorkbookGrid.ColumnName(position.Column)}{position.Row + 1}: {cell.Value}", modules.Revision, Sheet: sheet.Id, Cell: position));
        _scriptDiagnostics.ItemsSource = entries;
    }
    private async Task ActivateScriptDiagnosticAsync()
    {
        if (_scriptDiagnostics?.SelectedItem is not ScriptDiagnostic entry) return;
        if (entry.DraftVersion is { } version)
        {
            if (entry.Module is not { } id || _scriptPage is null || !_scriptPage.Drafts.Models.TryGetValue(id, out var model) || model.Version != version || model.ObservedVersion != version)
                throw new InvalidOperationException("此编译诊断属于旧草稿，请重新检查。");
            await _scriptPage.LocateAsync(id, entry.Line, entry.Column); return;
        }
        if (!await CommitEditorsAsync()) return;
        if (entry.Revision != Document.Snapshot.ScriptModules.Revision) throw new InvalidOperationException("此运行诊断属于旧源码，请重新运行后定位。");
        if (entry.Sheet is { } sheet && entry.Cell is { } cell) { Document.Session.UpdateNavigation(sheet, cell); Synchronize(); ShowBook(); Activate(); _grid.Select(cell); }
        else if (entry.Module is { } module && _scriptPage is not null) { await _scriptPage.LocateAsync(module, entry.Line, entry.Column); _selectedModule = module; RefreshScriptChrome(); }
    }
    private async Task ChangeScriptLanguageAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var current = Document.Snapshot.ScriptModules;
        if (current.ScriptApiVersion != 2) throw new InvalidOperationException("TypeScript 仅支持 API v2 模块；API v1 保持原有 JavaScript 语义。");
        var module = current.Modules[_selectedModule];
        var dialog = new FieldsDialog("模块语言").Choice("language", _selectedModule, ["JavaScript", "TypeScript"], (int)module.Language);
        if (!await dialog.ShowAsync(this) || !await CommitScriptsAsync()) return;
        current = Document.Snapshot.ScriptModules; module = current.Modules[_selectedModule];
        var language = (ScriptLanguage)dialog.Index("language");
        if (language == module.Language) return;
        var path = System.IO.Path.ChangeExtension(module.Path, language == ScriptLanguage.TypeScript ? ".ts" : ".js").Replace('\\', '/');
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(current.Modules.SetItem(_selectedModule, module with { Language = language, Path = path }),
            current.EntryModules, current.Revision, current.BindingRevision, 2, current.Bindings)));
        DisableScripts(); await RefreshScriptPageAfterMutationAsync(); Synchronize();
    }
}
