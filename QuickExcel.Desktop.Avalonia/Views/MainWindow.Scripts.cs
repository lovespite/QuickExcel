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
        var menu = new MenuItem { Header = UiText.Get("脚本") };
        menu.Items.Add(CommandMenu(UiText.Get("脚本编辑器"), OpenScriptsAsync));
        menu.Items.Add(CommandMenu(UiText.Get("调试控制台"), OpenScriptConsoleAsync));
        menu.Items.Add(CommandMenu(UiText.Get("应用脚本并重算"), ApplyScriptsAsync));
        menu.Items.Add(CommandMenu(UiText.Get("停用脚本"), () => { DisableScripts(); return Task.CompletedTask; }));
        menu.Items.Add(CommandMenu(UiText.Get("签名详情"), ScriptSignatureDetailsAsync));
        menu.Items.Add(CommandMenu(UiText.Get("签署脚本"), SignScriptsAsync));
        menu.Items.Add(CommandMenu(UiText.Get("受信任证书"), TrustedCertificatesAsync)); _menu.Items.Add(menu);
        RibbonGroup(UiText.Get("公式"), UiText.Get("脚本"), RibbonButton(RibbonIcon.Code, UiText.Get("脚本编辑器"), OpenScriptsAsync, false), RibbonButton(UiText.Get("应用并重算"), ApplyScriptsAsync, false));
        InstallAutomationRibbon();
    }
    private async Task OpenScriptsAsync()
    {
        if (_scriptWindow is not null) { _scriptWindow.Activate(); return; }
        if (!await CommitEditorsAsync()) return;
        _scriptDocument = Document.Session.SessionId; _selectedModule = Document.Snapshot.ScriptModules.EntryModules[0];
        var window = _scriptWindow = new Window { Title = UiText.Get("脚本工作区 · QuickExcel"), Width = 1080, Height = 760, MinWidth = 760, MinHeight = 500, WindowStartupLocation = WindowStartupLocation.CenterOwner };
        var root = new DockPanel(); var toolbar = new WrapPanel { Margin = new(8), Orientation = Orientation.Horizontal };
        void Button(string text, Func<Task> action)
        { var button = new Button { Content = text, Margin = new(4) }; button.Click += async (_, _) => await PerformAsync(action); toolbar.Children.Add(button); }
        Button(UiText.Get("确认使用 API v3"), MigrateScriptApiAsync); Button(UiText.Get("移除全部脚本"), RemoveScriptsAsync); Button(UiText.Get("添加模块"), AddScriptModuleAsync); Button(UiText.Get("应用并重算"), ApplyScriptsAsync);
        Button(UiText.Get("停用"), () => { DisableScripts(); return Task.CompletedTask; }); Button(UiText.Get("保存"), async () => await SaveAsync(false));
        _actionPicker = new ComboBox
        {
            MinWidth = 180, Margin = new(4), ItemTemplate = new FuncDataTemplate<ScriptActionDescriptor>((item, _) =>
            {
                if (item is null) return null;
                var text = new StackPanel { Spacing = 1 };
                text.Children.Add(new TextBlock { Text = item.Name, FontWeight = FontWeight.SemiBold, TextTrimming = TextTrimming.CharacterEllipsis });
                if (!string.IsNullOrWhiteSpace(item.Description))
                    text.Children.Add(new TextBlock { Text = item.Description, FontSize = 11, TextTrimming = TextTrimming.CharacterEllipsis });
                ToolTip.SetTip(text, string.IsNullOrWhiteSpace(item.Description) ? item.Name : item.Name + "\n" + item.Description);
                return text;
            })
        };
        AutomationProperties.SetName(_actionPicker, UiText.Get("脚本操作")); toolbar.Children.Add(_actionPicker); Button(UiText.Get("运行操作"), RunScriptActionAsync);
        RefreshScriptActions(); _actionPicker.ItemsSource = _availableScriptActions; _actionPicker.SelectedIndex = _availableScriptActions.Count > 0 ? 0 : -1;
        Button(UiText.Get("签名"), ScriptSignatureDetailsAsync);
        _scriptConsoleToggle = new Button { Margin = new(4) };
        _scriptConsoleToggle.Click += (_, _) => ToggleScriptConsole(); toolbar.Children.Add(_scriptConsoleToggle);
        _scriptRecovery = new Button { Content = UiText.Get("恢复已确认草稿"), IsVisible = false, Margin = new(4) };
        _scriptRecovery.Click += async (_, _) => await PerformAsync(RecoverScriptPageAsync); toolbar.Children.Add(_scriptRecovery);
        DockPanel.SetDock(toolbar, Dock.Top); root.Children.Add(toolbar);
        _scriptStatus = new TextBlock { Margin = new(12, 5), TextWrapping = TextWrapping.Wrap };
        DockPanel.SetDock(_scriptStatus, Dock.Bottom); root.Children.Add(_scriptStatus);
        _scriptDiagnostics = new ListBox { ItemTemplate = new FuncDataTemplate<ScriptDiagnostic>((item, _) => new TextBlock { Text = item?.Label, TextWrapping = TextWrapping.Wrap }) };
        AutomationProperties.SetName(_scriptDiagnostics, UiText.Get("运行日志与错误；静态错误在编辑器内显示"));
        _scriptDiagnostics.DoubleTapped += async (_, _) => await PerformAsync(ActivateScriptDiagnosticAsync);
        _scriptDiagnostics.KeyDown += async (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; await PerformAsync(ActivateScriptDiagnosticAsync); } };
        root.Children.Add(CreateScriptWorkspace(CreateScriptOutput()));
        var scriptContent = _scriptContent!;
        window.Content = _scriptExecution = new ExecutionOverlay(window, root, () => _consoleInput?.IsComposing == true,
            // Native WebViews occupy their own surface, above Avalonia overlays. Hide only after the delay.
            visible => scriptContent.IsVisible = !visible);
        window.AddHandler(KeyDownEvent, HandleScriptWorkspaceKey, RoutingStrategies.Tunnel);
        window.Closing += async (_, e) =>
        {
            if (_scriptClosing) return; e.Cancel = true;
            await PerformAsync(async () =>
            {
                if (_scriptPage is null || _scriptPage.Drafts.Failed)
                {
                    if (await Dialogs.Confirm(this, UiText.Get("关闭失败的编辑器"), UiText.Get("放弃尚未提交到工作簿的脚本草稿并关闭编辑器？已提交的文档更改仍然保留。可选择取消，然后恢复已确认草稿。"))) await CloseScriptWindowAsync();
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
        page.DraftChanged += () => { DisableScripts(); UpdateTitle(); if (_scriptWindow is not null) _scriptWindow.Title = UiText.Get("* 脚本工作区 · QuickExcel"); if (_scriptStatus is not null) _scriptStatus.Text = UiText.Get("源码草稿已修改，脚本授权已撤销。"); };
        page.SaveRequested += () => _ = PerformAsync(async () => await SaveAsync(false));
        page.FocusRequested += FocusScriptNavigation;
        page.ConsoleRequested += ToggleScriptConsole;
        page.NavigateRequested += NavigateScriptAsync;
        _scriptContent!.Content = page.Browser.View; RefreshScriptChrome();
        await page.LoadAsync(_selectedModule);
        if (_selectedDefinition is { } definition) await page.OpenDefinitionAsync(definition);
        RefreshScriptChrome();
    }
    private void ScriptPageFailed(Exception error)
    {
        DisableScripts(); if (_scriptContent is not null) _scriptContent.IsEnabled = false;
        if (_scriptRecovery is not null) _scriptRecovery.IsVisible = true;
        if (_scriptStatus is not null) _scriptStatus.Text = UiText.Get("页面已停止，保存和运行已阻止。已确认草稿仍在主程序内。") + error.Message;
        Error(error);
    }
    private async Task RecoverScriptPageAsync()
    {
        if (!await Dialogs.Confirm(this, UiText.Get("恢复脚本页面"), UiText.Get("重新加载主程序已确认的草稿。浏览器尚未确认的最后输入无法恢复；请核对源码后再保存或运行。"))) return;
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
            if (_scriptPage is null || _scriptDocument != Document.Session.SessionId) throw new InvalidOperationException(UiText.Get("脚本页面未加载或所属文档已变化。"));
            await _scriptPage.FlushAsync(() => Document.Snapshot.ScriptModules, async modules =>
            {
                var document = Document;
                var current = Document.Snapshot.ScriptModules;
                if (validateTypeScript && modules.HasTypeScript)
                {
                    if (_scriptCancellation is not null) throw new InvalidOperationException(UiText.Get("已有脚本工作正在进行。"));
                    using var cancellation = new CancellationTokenSource();
                    using var feedback = BeginExecution(UiText.Get("正在检查 TypeScript…"), cancellation);
                    _scriptCancellation = cancellation;
                    var revision = document.Snapshot.Revision;
                    if (_scriptStatus is not null) _scriptStatus.Text = UiText.Get("正在检查 TypeScript；按 Esc 取消…");
                    try
                    {
                        await ScriptCompilation.PrepareAsync(modules, new(document.Session.SessionId, revision, 0, modules.Revision), cancellation.Token);
                        cancellation.Token.ThrowIfCancellationRequested();
                        if (!ReferenceEquals(document, Document) || document.Snapshot.Revision != revision || !ReferenceEquals(current, document.Snapshot.ScriptModules))
                            throw new InvalidOperationException(UiText.Get("编译期间文档已变化，未应用过期产物。"));
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
        catch (OperationCanceledException) { if (_scriptStatus is not null) _scriptStatus.Text = UiText.Get("编译已取消，草稿未应用。"); return false; }
        catch (Exception ex) { Error(ex); if (_scriptStatus is not null) _scriptStatus.Text = ex.Message; return false; }
    }
    private async Task CloseScriptWindowAsync()
    {
        if (_scriptWindow is null) return;
        _scriptCancellation?.Cancel();
        var window = _scriptWindow; _scriptClosing = true;
        try { if (_scriptPage is not null) await _scriptPage.DisposeAsync(); window.Close(); }
        finally
        {
            _scriptWindow = null; _scriptPage = null; _scriptContent = null; _modulePicker = null; _actionPicker = null; _scriptStatus = null; _scriptDiagnostics = null; _scriptRecovery = null; _scriptClosing = false;
            _scriptExecution = null;
            _scriptEditorPane = null; _scriptOutputSplitter = null; _scriptConsoleToggle = null;
            _scriptModuleGroups.Clear(); _scriptNavigationSelection = null;
            _definitionPicker = null; _definitionPanel = null; _selectedDefinition = null; _scriptDefinitionFiles.Clear();
            _consoleInput = null; _consoleOutput = null; _consoleRun = null; _consoleCopy = null;
            _consoleFilter = null; _consoleLevel = null; _consoleCount = null; _consoleEmpty = null;
            _consoleVisibleEntries = []; _scriptOutputTabs = null;
        }
    }
    private async Task RefreshScriptPageAfterMutationAsync()
    {
        if (_scriptPage is not { } page || _scriptWindow is null) return;
        var modules = Document.Snapshot.ScriptModules;
        if (page.Drafts.Models.Count == modules.Modules.Count && modules.Modules.All(p => page.Drafts.Models.TryGetValue(p.Key, out var draft) &&
            draft.Path == p.Value.Path && draft.Language == p.Value.Language && draft.Committed == p.Value.Source)) return;
        if (page.Drafts.IsDirty || page.Drafts.Failed)
            throw new InvalidOperationException(UiText.Get("包内源码与编辑器草稿冲突；已保留草稿，不能覆盖新的包源码。"));
        await page.DisposeAsync(); _scriptPage = null;
        if (!modules.Modules.ContainsKey(_selectedModule)) _selectedModule = modules.EntryModules[0];
        try { await CreateScriptPageAsync(new(modules, Document.Snapshot.FileState?.FormulaReadOnly == true)); }
        catch (Exception error) { ScriptPageFailed(error); throw; }
    }
    private void RefreshScriptChrome(bool revealSelected = false)
    {
        if (_scriptWindow is null) return;
        _scriptWindow.Title = (_scriptPage?.Drafts.IsDirty == true ? "* " : "") + UiText.Get("脚本工作区 · QuickExcel");
        RefreshScriptNavigation(revealSelected);
        var modules = Document.Snapshot.ScriptModules;
        if (_scriptStatus is not null && _scriptPage?.Drafts.Failed != true && _selectedDefinition is { } definition)
            _scriptStatus.Text = UiText.Format($"{definition} · 定义文件（只读）");
        else if (_scriptStatus is not null && _scriptPage?.Drafts.Failed != true)
            _scriptStatus.Text = $"{modules.Modules[_selectedModule].Path} · {ScriptModuleProperties(modules, _selectedModule)} · API v{modules.ScriptApiVersion} · {(Document.ScriptsEnabled ? UiText.Get("已授权") : UiText.Get("未授权"))}" +
             (!modules.CanExecute ? " · " + UiText.Get(ScriptModuleSet.MigrationMessage) : DesktopServices.SupportsScripts ? string.Empty : UiText.Get("(此平台尚未开放脚本执行)"));
    }
    private void DisableScripts()
    { _scriptCancellation?.Cancel(); Document.DisableScripts(); ClearScriptActions(); RefreshScriptActions(); RefreshScriptChrome(); }
    private async Task MigrateScriptApiAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var old = Document.Snapshot.ScriptModules;
        if (old.CanExecute) return;
        var dialog = new FieldsDialog(UiText.Get("确认脚本已重写"))
            .Note(UiText.Get("请先将全部模块重写为 QuickExcel API v3。此操作仅更新版本声明，不转换源码；原签名和执行授权将失效。"));
        if (!await dialog.ShowAsync(this)) return;
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(new(old.Modules, old.EntryModules,
            old.Revision, old.BindingRevision, 3, old.Bindings)));
        Synchronize(); RefreshScriptChrome();
    }
    private async Task RemoveScriptsAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var dialog = new FieldsDialog(UiText.Get("移除全部脚本")).Note(UiText.Get("移除全部模块和签名，工作表数据保留。此操作可撤销。"));
        if (!await dialog.ShowAsync(this)) return;
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(ScriptModuleSet.Single("")));
        _selectedModule = "main";
        await RefreshScriptPageAfterMutationAsync(); Synchronize();
    }
    private async Task AddScriptModuleAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var document = Document;
        if (document.Snapshot.ScriptModules.ScriptApiVersion != 3) throw new InvalidOperationException(UiText.Get("请先重写旧脚本，并选择“确认使用 API v3”；添加模块需要 API v3。"));
        if (document.Snapshot.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(UiText.Get(WorkbookFileState.FormulaReadOnlyMessage));
        var dialog = new FieldsDialog(UiText.Get("新建模块")).Text("id", UiText.Get("模块逻辑 ID"))
            .Text("description", UiText.Get("描述（可选）"), "")
            .Choice("role", UiText.Get("角色"), ["Common", "Actions", "UDF"], selected: 1)
            .Choice("language", UiText.Get("语言"), ["JavaScript (.js)", "TypeScript (.ts)"], selected: (int)ScriptLanguage.TypeScript)
            .Note(UiText.Get("UDF 只用于注册函数；Actions 只用于定义操作，一个文件对应一个操作。事件请在主入口或由主入口导入的 Common 模块中注册。"))
            .Note(UiText.Get("描述显示在操作列表中，留空时只显示模块 ID。角色和语言在创建后不可更改。"));
        ScriptModuleSet BuildModules()
        {
            if (!ReferenceEquals(document, Document)) throw new InvalidOperationException(UiText.Get("所属文档已变化，请重新创建模块。"));
            var old = document.Snapshot.ScriptModules; var id = dialog.String("id").Trim();
            if (old.Modules.ContainsKey(id)) throw new ArgumentException(UiText.Get("模块 ID 已存在。"));
            if (id.Length is < 1 or > 128 || id is "." or ".." || id.Any(c => !char.IsAsciiLetterOrDigit(c) && c is not ('_' or '-' or '.')))
                throw new ArgumentException(UiText.Get("模块 ID 须为 1–128 个英文字母、数字、下划线、连字符或点，且不能仅为 . 或 ..。"));
            var description = dialog.String("description").Trim();
            if (description.Length > ScriptModuleSet.MaximumDescriptionLength)
                throw new ArgumentException(UiText.Format($"模块描述不能超过 {ScriptModuleSet.MaximumDescriptionLength} 个字符。"));
            var binding = dialog.Index("role") switch
            {
                0 => null,
                1 => "actions",
                2 => "formula-functions",
                _ => throw new ArgumentException(UiText.Get("请选择模块角色。"))
            };
            var language = (ScriptLanguage)dialog.Index("language");
            var bindings = old.Bindings.ToList(); var entries = old.EntryModules.ToList();
            if (binding is not null) { bindings.Add(new(id, binding)); entries.Add(id); }
            var module = new ScriptModule($"scripts/{id}.{(language == ScriptLanguage.TypeScript ? "ts" : "js")}",
                ScriptModuleTemplates.Create(id, language, binding == "formula-functions", binding == "actions"), language, description);
            return new(old.Modules.Add(id, module), entries, old.Revision, old.BindingRevision, old.ScriptApiVersion, bindings);
        }
        dialog.Validate = () => BuildModules();
        if (!await dialog.ShowAsync(this) || !await CommitScriptsAsync()) return;
        _workspace.Editor.Execute(new UpdateScriptModulesCommand(BuildModules()));
        // The module set changed; start a new page generation after the existing page has acknowledged every draft.
        if (_scriptPage is not null) await _scriptPage.DisposeAsync(); _selectedModule = dialog.String("id").Trim(); _selectedDefinition = null;
        await CreateScriptPageAsync(new(Document.Snapshot.ScriptModules)); Synchronize();
    }
    private async Task ApplyScriptsAsync()
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync(validateTypeScript: true)) return;
        if (!DesktopServices.SupportsScripts) throw new PlatformNotSupportedException(UiText.Get("此平台的脚本隔离尚未开放，不能执行脚本。"));
        ClearScriptActions(); RefreshScriptActions();
        if (!await CheckScriptTrustAsync()) return;
        var modules = Document.Snapshot.ScriptModules;
        await EnableVerifiedScriptsAsync();
        if (!ReferenceEquals(modules, Document.Snapshot.ScriptModules) || !Document.ScriptsEnabled || modules.Signature is null) return;
        try
        {
            var trust = await Task.Run(() => ScriptTrust.Verify(modules));
            if (trust.CanAutoEnable && !trust.ExecutionPolicyAccepted && ReferenceEquals(modules, Document.Snapshot.ScriptModules))
                await Task.Run(() => ScriptTrust.AcceptExecutionPolicy(modules));
        }
        catch (QuickExcel.Security.UserProtectionUnavailableException ex)
        { _status.Text = UiText.Get("脚本已应用；无法记住当前引擎策略，下次打开仍需手动应用：") + ex.Message; }
    }
    private async Task EnableVerifiedScriptsAsync()
    {
        await ScriptWorkAsync(async token =>
        {
            var document = Document; var modules = document.Snapshot.ScriptModules;
            await Task.Run(() => document.EnableScriptsAndCalculate(token), token);
            var actions = document.Snapshot.ScriptModules.HasActions ? await Task.Run(() => document.ListActions(token), token) : [];
            SetScriptActions(document, modules, actions);
            _calculationPending = false;
            _status.Text = UiText.Format($"已加载 {actions.Count} 个脚本操作。");
        });
        await DispatchEventsAsync();
    }
    private Task RunScriptActionAsync() => RunScriptActionAsync(_actionPicker?.SelectedItem as ScriptActionDescriptor);
    private async Task RunScriptActionAsync(ScriptActionDescriptor? action)
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (!await CommitEditorsAsync(validateTypeScript: true)) return;
        if (!ScriptActionsCurrent || action is null || !_availableScriptActions.Contains(action)) throw new InvalidOperationException(UiText.Get("请先加载脚本并选择操作。"));
        var document = Document;
        _status.Text = UiText.Get("正在执行：") + action.Name;
        await ScriptWorkAsync(async token =>
        {
            var result = await Task.Run(() => document.RunAction(action.Name, token), token);
            _status.Text = UiText.Format($"已执行：{action.Name}") + (result.Changed ? UiText.Get("（可撤销）") : UiText.Get("（无变更）"));
        });
    }
    private async Task ScriptWorkAsync(Func<CancellationToken, Task> work)
    {
        _calculationTimer.Stop(); using var cancellation = new CancellationTokenSource();
        using var feedback = BeginExecution(UiText.Get("正在执行代码…"), cancellation);
        _scriptCancellation = cancellation;
        RefreshScriptActionButtons(); RefreshScriptConsole();
        try { await work(cancellation.Token); }
        finally
        {
            _scriptCancellation = null;
            await RefreshScriptPageAfterMutationAsync();
            Synchronize(); RefreshScriptChrome(); RefreshScriptDiagnostics();
            ScheduleCalculation();
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
                throw new InvalidOperationException(UiText.Get("此编译诊断属于旧草稿，请重新检查。"));
            await _scriptPage.LocateAsync(id, entry.Line, entry.Column); _selectedModule = id; _selectedDefinition = null; RefreshScriptChrome(revealSelected: true); return;
        }
        if (!await CommitEditorsAsync()) return;
        if (entry.Revision != Document.Snapshot.ScriptModules.Revision) throw new InvalidOperationException(UiText.Get("此运行诊断属于旧源码，请重新运行后定位。"));
        if (entry.Sheet is { } sheet && entry.Cell is { } cell) { Document.Session.UpdateNavigation(sheet, cell); Synchronize(); ShowBook(); Activate(); _grid.Select(cell); }
        else if (entry.Module is { } module && _scriptPage is not null) { await _scriptPage.LocateAsync(module, entry.Line, entry.Column); _selectedModule = module; _selectedDefinition = null; RefreshScriptChrome(revealSelected: true); }
    }
}
