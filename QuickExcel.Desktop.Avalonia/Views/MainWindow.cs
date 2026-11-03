using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow : Window
{
    private readonly WorkbookWorkspace _workspace;
    private readonly WorkbookServices _services;
    private readonly WorkbookGrid _grid = new();
    private readonly CompositionTextBox _formula = new() { MinHeight = 30 };
    private readonly TextBlock _address = new() { Width = 86, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _status = new() { Text = "就绪", Margin = new(12, 5) };
    private readonly ListBox _sheets = new() { MinHeight = 36, SelectionMode = SelectionMode.Single };
    private readonly DockPanel _book = new();
    private readonly StackPanel _start = new() { Margin = new(60), Spacing = 18 };
    private readonly StackPanel _recent = new() { Spacing = 6 };
    private readonly Menu _menu = new();
    private readonly RecentFileStore _recentFiles;
    private readonly DesktopSettings _settings;
    private ILocalFileDialogs _files;
    private Guid _boundSession;
    private SheetId _boundSheet;
    private bool _synchronizing, _formulaDirty, _allowClose, _closePending, _working;
    private CellAddress? _formulaOwner;
    private Task _operation = Task.CompletedTask;
#if QUICKEXCEL_DIAGNOSTICS
    private Func<Task<SaveDecision>>? _discardChoice;
#endif
    private WorkbookDocument Document => _workspace.Document;
    private int SheetIndex => Document.Snapshot.Sheets.IndexOf(Document.Snapshot.GetSheet(Document.Snapshot.ActiveSheetId));

    public MainWindow(string userDataDirectory)
    {
        Title = "QuickExcel · Avalonia"; Width = 1280; Height = 820; MinWidth = 760; MinHeight = 480;
        FontFamily = new FontFamily("Segoe UI, Noto Sans CJK SC, sans-serif"); FontSize = 13;
        var services = _services = DesktopServices.Create(); _workspace = new(services, new WorkbookFileAccess(services.Storage));
        _files = new AvaloniaFileDialogs(this);
        _recentFiles = new(Path.Combine(userDataDirectory, "recent-files.json")); _recentFiles.Load();
        _settings = new(Path.Combine(userDataDirectory, "settings.json")); _settings.Load();
        _userDataDirectory = userDataDirectory;
        BuildView(); InstallEditingCommands(); InstallFormulaEditing(); InstallRibbon(); InstallCalculation(); InstallSheetCommands(); InstallSearch(); InstallDataCommands(); InstallValidation(); InstallCleaning(); InstallObjects(); InstallFileCommands(); InstallScriptCommands(); InstallPrintCommands();
        ArrangeRibbonTabs(); InstallScriptEvents();
        _grid.ResolveStyle = ConditionalPresentation.GetStyle;
        _grid.RenderingFailed += Error;
        _grid.CommitCellAsync = CommitCellAsync;
        _grid.BeforeSelection = CommitFormula;
        _grid.ZoomRequested += async zoom => await PerformAsync(() => ZoomAsync(zoom, preserveDraft: true));
        _grid.DraftEnded += () => { CloseFormulaHelp(); _grid.HighlightReference(null); };
        _grid.SelectionChanged += () =>
        {
            if (_synchronizing) return;
            Document.Session.UpdateNavigation(Document.Snapshot.ActiveSheetId, _grid.Current); QueueCellEvent("selection.changed"); UpdateFormula(); UpdateTitle(); UpdateRibbon(); UpdateValidation(); QueueStatistics();
        };
        _formula.GotFocus += (_, _) => _formulaOwner ??= new(Document.Snapshot.ActiveSheetId, _grid.Current.Row, _grid.Current.Column);
        _formula.PropertyChanged += (_, e) => { if (e.Property == TextBox.TextProperty && !_synchronizing) { if (!_formulaDirty) QueueCellEvent("cell.edit"); _formulaDirty = true; } };
        _formula.AddHandler(KeyDownEvent, async (_, e) =>
        {
            if (_formula.IsComposing) return;
            if (e.Key == Key.Enter) { e.Handled = true; if (await CommitFormulaAsync()) _grid.Select(_grid.Current); }
            if (e.Key == Key.Escape) { e.Handled = true; _formulaDirty = false; _formulaOwner = null; CloseFormulaHelp(); _grid.HighlightReference(null); UpdateFormula(); _grid.Select(_grid.Current); }
        }, RoutingStrategies.Tunnel);
        _sheets.SelectionChanged += async (_, _) =>
        {
            if (_synchronizing || _sheets.SelectedItem is not SheetItem selected || selected.Id == Document.Snapshot.ActiveSheetId) return;
            if (!await CommitEditorsAsync()) { Synchronize(); return; }
            Document.Session.UpdateNavigation(selected.Id, Document.Snapshot.GetSheet(selected.Id).View.ActiveCell); Synchronize();
        };
        _workspace.Changed += Synchronize;
        AddHandler(KeyDownEvent, Shortcuts, RoutingStrategies.Tunnel);
        Closing += OnClosing;
        Deactivated += (_, _) => _grid.CancelInputGesture();
        Closed += (_, _) => { _grid.Dispose(); _workspace.Dispose(); };
        Opened += async (_, _) =>
        {
            // Opened can run before the platform main loop starts. GTK's native WebView
            // needs the GLib dispatcher running before an owned page is constructed.
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            if (Program.SelfCheckDirectory is { } evidence) { await RunSelfCheckAsync(evidence); return; }
#if QUICKEXCEL_DIAGNOSTICS
            if (Program.SmokeDirectory is { } executionDirectory && Program.SmokeMode == "execution") { await RunExecutionSmokeAsync(executionDirectory); return; }
            if (Program.SmokeDirectory is { } automationDirectory && Program.SmokeMode == "automation") { await RunAutomationSmokeAsync(automationDirectory); return; }
            if (Program.SmokeDirectory is { } featureDirectory && Program.SmokeMode == "ribbon-filter") { await RunRibbonFilterSmokeAsync(featureDirectory); return; }
            if (Program.SmokeDirectory is { } drawingDirectory && Program.SmokeMode == "drawing") { await RunDrawingSmokeAsync(drawingDirectory); return; }
            if (Program.SmokeDirectory is { } pageBreakDirectory && Program.SmokeMode == "page-break") { await RunPageBreakSmokeAsync(pageBreakDirectory); return; }
            if (Program.SmokeDirectory is { } compatibilityDirectory && Program.SmokeMode == "compatibility") { await RunCompatibilitySmokeAsync(compatibilityDirectory); return; }
            if (Program.SmokeDirectory is { } smoke && Program.SmokeMode != "manual") { if (Program.SmokeMode == "grid") await RunGridSmokeAsync(smoke); else if (Program.SmokeMode == "benchmark") await RunGridBenchmarkAsync(smoke); else if (Program.SmokeMode == "web") await RunWebSmokeAsync(smoke); else if (Program.SmokeMode == "web-benchmark") await RunWebBenchmarkAsync(smoke); else await RunWorkspaceSmokeAsync(smoke); return; }
#endif
            if (Program.InitialFile is { } path) await PerformAsync(() => OpenPathAsync(path));
            if (Program.SettingsMigrationError is { } migrationError) _status.Text = migrationError;
        };
        Synchronize(); ShowHome();
    }
    private sealed record SheetItem(SheetId Id, string Name) { public override string ToString() => Name; }
    private void BuildView()
    {
        MenuItem Item(string text, Func<Task> action) { var item = new MenuItem { Header = text }; item.Click += async (_, _) => await PerformAsync(action); return item; }
        _menu.Items.Add(new MenuItem { Header = "文件", Items =
        {
            Item("新建  Ctrl+N", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } }),
            Item("打开  Ctrl+O", OpenAsync), Item("保存  Ctrl+S", async () => await SaveAsync(false)),
            Item("另存为", async () => await SaveAsync(true)), Item("最近文件", async () => { if (await ConfirmDiscardAsync()) ShowHome(); }),
            Item("关闭工作簿", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowHome(); } }), Item("退出", () => { Close(); return Task.CompletedTask; })
        } });
        _menu.Items.Add(new MenuItem { Header = "编辑", Items =
        { Item("撤销  Ctrl+Z", () => History(false)), Item("重做  Ctrl+Y", () => History(true)), Item("清除内容", ClearAsync) } });
        _menu.Items.Add(new MenuItem { Header = "工作表", Items =
        { Item("添加工作表", AddSheetAsync), Item("重命名", RenameSheetAsync), Item("删除工作表", DeleteSheetAsync) } });
        _menu.Items.Add(new MenuItem { Header = "视图", Items =
        { Item("100%", () => ZoomAsync(100)), Item("125%", () => ZoomAsync(125)), Item("150%", () => ZoomAsync(150)),
          Item("标签在底部", () => PlaceSheets(SheetNavigationPlacement.Bottom)), Item("标签在左侧", () => PlaceSheets(SheetNavigationPlacement.Left)), Item("标签在右侧", () => PlaceSheets(SheetNavigationPlacement.Right)) } });
        _menu.Items.Add(Item("重算  F9", CalculateAsync));
        var formulaRow = new DockPanel { Margin = new(8, 4) }; DockPanel.SetDock(_address, Dock.Left); formulaRow.Children.Add(_address); formulaRow.Children.Add(_formula);
        DockPanel.SetDock(formulaRow, Dock.Top); _book.Children.Add(formulaRow); _book.Children.Add(_sheets); _book.Children.Add(_grid); ApplySheetPlacement();
        _start.Children.Add(new TextBlock { Text = "QuickExcel", FontSize = 36, FontWeight = FontWeight.SemiBold });
        _start.Children.Add(new TextBlock { Text = "工作簿与脚本", FontSize = 18, Foreground = Brushes.Gray });
        void StartButton(string text, Func<Task> action) { var button = new Button { Content = text, MinWidth = 220 }; button.Click += async (_, _) => await PerformAsync(action); _start.Children.Add(button); }
        StartButton("新建工作簿", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } });
        StartButton("打开工作簿", OpenAsync);
        StartButton("打开示例", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(true); ShowBook(); } });
        StartButton("返回当前工作簿", () => { ShowBook(); return Task.CompletedTask; });
        _start.Children.Add(new TextBlock { Text = "最近文件", FontSize = 18 }); _start.Children.Add(_recent);
        var content = new global::Avalonia.Controls.Grid(); content.Children.Add(_book); content.Children.Add(new ScrollViewer { Content = _start });
        var root = new DockPanel(); DockPanel.SetDock(_menu, Dock.Top);
        var footer = new DockPanel(); DockPanel.SetDock(_selectionStatistics, Dock.Right); footer.Children.Add(_selectionStatistics); footer.Children.Add(_status);
        DockPanel.SetDock(footer, Dock.Bottom); root.Children.Add(_menu); root.Children.Add(footer); root.Children.Add(content);
        Content = _execution = new ExecutionOverlay(this, root, () => _grid.IsComposing || _formula.IsComposing);
    }
    private void ShowHome()
    {
        if (_book.IsVisible) Document.Session.RaiseScriptEvent("workbook.deactivate");
        _book.IsVisible = false; ((Control)_start.Parent!).IsVisible = true; _recent.Children.Clear();
        foreach (var path in _recentFiles.Paths)
        {
            var button = new Button { Content = Path.GetFileName(path), HorizontalAlignment = HorizontalAlignment.Left }; ToolTip.SetTip(button, path);
            button.Click += async (_, _) => await PerformAsync(async () => { if (await ConfirmDiscardAsync()) await OpenPathAsync(path); }); _recent.Children.Add(button);
        }
    }
    private void ShowBook() { if (!_book.IsVisible) Document.Session.RaiseScriptEvent("workbook.activate"); ((Control)_start.Parent!).IsVisible = false; _book.IsVisible = true; }
    private void Synchronize()
    {
        if (_synchronizing) return;
        if (_filterContextCurrent?.Invoke() == false) CloseFilterPanel();
        ObserveDocument();
        _synchronizing = true;
        try
        {
            _book.IsEnabled = !NonExecutionBusy; _menu.IsEnabled = !NonExecutionBusy;
            if (_boundSession != Document.Session.SessionId || _boundSheet != Document.Snapshot.ActiveSheetId)
            {
                _formulaDirty = false; _formulaOwner = null;
                _boundSession = Document.Session.SessionId; _boundSheet = Document.Snapshot.ActiveSheetId;
                _grid.Bind(Document.Snapshot, SheetIndex, WorkbookLimits.MaxRows, WorkbookLimits.MaxColumns);
            }
            else _grid.Refresh(Document.Snapshot);
            var items = Document.Snapshot.Sheets.Where(s => s.Visibility == SheetVisibility.Visible).Select(s => new SheetItem(s.Id, s.Name)).ToArray();
            _sheets.ItemsSource = items; _sheets.SelectedItem = items.FirstOrDefault(s => s.Id == Document.Snapshot.ActiveSheetId);
        }
        finally { _synchronizing = false; }
        UpdateFormula(); UpdateTitle(); UpdateRibbon(); UpdateValidation(); UpdateOutline(); QueueStatistics(); RefreshPrintStatus(); RefreshScriptConsole();
    }
    private void UpdateTitle()
    {
        Title = (Document.IsDirty || _scriptPage?.Drafts.IsDirty == true ? "* " : "") + (_workspace.FilePath is { } path ? Path.GetFileName(path) : "未命名.xlsq") + " — QuickExcel";
        _address.Text = WorkbookGrid.ColumnName(_grid.Current.Column) + (_grid.Current.Row + 1);
        var sheet = Document.Snapshot.Sheets[SheetIndex];
        var value = Document.GetCellValue(SheetIndex, _grid.Current.Row, _grid.Current.Column);
        var readOnly = _grid.IsCellReadOnly(_grid.Current);
        AutomationProperties.SetName(_grid, $"工作表 {sheet.Name}，{WorkbookGrid.ColumnName(_grid.Current.Column)} 列，第 {_grid.Current.Row + 1} 行，值 {value}，选中 {_grid.SelectionRanges.Sum(r => r.Count)} 格，{(readOnly ? "只读" : "可编辑")}");
        AutomationProperties.SetName(_grid.CellEditor, $"单元格 {_address.Text}，{(readOnly ? "只读" : "编辑")}");
        AutomationProperties.SetName(_formula, "公式栏 " + _address.Text); AutomationProperties.SetName(_sheets, "工作表列表");
    }
    private void UpdateFormula()
    {
        if (_formulaDirty || _formula.IsComposing) return;
        var sheet = Document.Snapshot.Sheets[SheetIndex];
        _synchronizing = true; _formula.Text = ProtectionSafety.HidesFormula(Document.Snapshot, sheet, _grid.Current.Row, _grid.Current.Column) ? "" : Document.GetCellInput(SheetIndex, _grid.Current.Row, _grid.Current.Column); _synchronizing = false;
        _formulaOwner = null; _formula.IsReadOnly = _grid.SelectedDrawingObject is not null || _grid.IsCellReadOnly(_grid.Current);
    }
    private Task<bool>? _formulaCommit;
    private bool CommitFormula()
    {
        var task = CommitFormulaAsync(); return task.IsCompletedSuccessfully && task.Result;
    }
    private Task<bool> CommitFormulaAsync()
    {
        if (_formulaCommit is { } pending) return pending;
        if (_formula.IsComposing) return Task.FromResult(false);
        if (!_formulaDirty) return Task.FromResult(true);
        var task = CommitFormulaDraftAsync();
        if (!task.IsCompleted) _formulaCommit = task;
        return task;
    }
    private async Task<bool> CommitFormulaDraftAsync()
    {
        _formula.IsReadOnly = true;
        try
        {
            var owner = _formulaOwner ?? new(Document.Snapshot.ActiveSheetId, _grid.Current.Row, _grid.Current.Column);
            var index = Document.Snapshot.Sheets.IndexOf(Document.Snapshot.GetSheet(owner.SheetId));
            if (!await CommitCellAsync(index, new(owner.Row, owner.Column), _formula.Text ?? "")) return false;
            _formulaDirty = false; _formulaOwner = null; CloseFormulaHelp(); _grid.HighlightReference(null); UpdateFormula(); return true;
        }
        finally { _formulaCommit = null; _formula.IsReadOnly = _grid.IsCellReadOnly(_grid.Current); }
    }
    private async Task<bool> CommitCellAsync(int index, CellPosition owner, string text)
    {
        if (_workspace.IsBusy || _execution.IsRunning) return false;
        try
        {
            var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[index];
            var command = new SetCellsCommand([new(new(sheet.Id, owner.Row, owner.Column), CellInput.ParseForCell(text, snapshot, sheet, owner.Row, owner.Column))]);
            await _workspace.Editor.ExecuteAsync(command, ConfirmValidationAsync, snapshot.Revision);
            QueueCellEvent("cell.commit", sheet.Id, owner);
            Synchronize(); _status.Text = "已修改"; return true;
        }
        catch (Exception ex) { Error(ex); return false; }
    }
    private bool CommitEditors()
    {
        var task = CommitEditorsAsync(); return task.IsCompletedSuccessfully && task.Result;
    }
    private async Task<bool> CommitEditorsAsync(bool validateTypeScript = false) => !_workspace.IsBusy && !_execution.IsRunning && await _grid.CommitAsync() && await CommitFormulaAsync() && await CommitScriptsAsync(validateTypeScript);
    private bool TryEdit(Func<CommandResult> edit)
    {
        if (_workspace.IsBusy || _execution.IsRunning) return false;
        try { edit(); Synchronize(); _status.Text = "已修改"; return true; }
        catch (Exception ex) { Error(ex); return false; }
    }
    private void Error(Exception ex) => _status.Text = ex.Message;
    private async Task PerformAsync(Func<Task> action)
    {
        if (_working || _execution.IsRunning) return;
        _working = true;
        try { await using var hold = _scriptPage?.HoldEdits(); _operation = action(); await _operation; }
        catch (Exception ex) { Error(ex); }
        finally { _working = false; }
    }
    private async Task<bool> ConfirmDiscardAsync()
    {
        if (!await CommitEditorsAsync()) return false;
        Document.Session.RaiseScriptEvent("workbook.beforeClose");
        try { await DispatchEventsAsync(); } catch (Exception error) { Error(error); }
        var accepted = !Document.IsDirty;
        if (!accepted)
        {
#if QUICKEXCEL_DIAGNOSTICS
            var choice = _discardChoice is null ? await Dialogs.Unsaved(this) : await _discardChoice();
#else
            var choice = await Dialogs.Unsaved(this);
#endif
            accepted = choice == SaveDecision.Discard || choice == SaveDecision.Save && await SaveAsync(false);
        }
        if (accepted) await CloseScriptWindowAsync();
        return accepted;
    }
    private async Task OpenAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var selected = await _files.OpenAsync(new("打开工作簿", [new("工作簿", ["*.xlsx", "*.xlsq"])]));
        if (selected is not null && await ConfirmDiscardAsync()) await OpenPathAsync(selected.Path);
    }
    private async Task OpenPathAsync(string path)
    {
        await _workspace.OpenAsync(new(path)); _recentFiles.Remember(path); ShowBook(); _status.Text = "已打开 " + Path.GetFileName(path);
        if (DesktopServices.SupportsScripts && Document.Snapshot.ScriptModules.Signature is not null)
        {
            var modules = Document.Snapshot.ScriptModules;
            try
            {
                var trust = await Task.Run(() => ScriptTrust.Verify(modules));
                // Source signatures remain verifiable, but existing certificate
                // pins do not silently consent to a new engine policy.
                if (ReferenceEquals(modules, Document.Snapshot.ScriptModules) && trust.CanAutoEnable)
                    _status.Text = "脚本签名有效；引擎策略已更新，请在脚本页显式应用一次。";
            }
            catch (QuickExcel.Security.UserProtectionUnavailableException ex) { _status.Text = "文件已打开；无法读取持久信任，脚本保持停用：" + ex.Message; }
        }
    }
    private async Task<bool> SaveAsync(bool saveAs)
    {
        await using var hold = _scriptPage?.HoldEdits();
        if (!await FlushEditsAsync()) return false;
        Document.Session.RaiseScriptEvent("workbook.beforeSave"); await DispatchEventsAsync();
        await CalculatePendingAsync();
        var snapshot = Document.Snapshot;
        var migration = snapshot.FileState?.RequiresMigration == true;
        var needsScripts = TypedWorkbookStorage.RequiresScriptContainer(snapshot);
        var current = _workspace.FilePath;
        var requiresNewContainer = needsScripts && current is not null && Path.GetExtension(current).Equals(".xlsx", StringComparison.OrdinalIgnoreCase);
        var issues = TypedWorkbookStorage.SaveIssues(snapshot); var discard = false; var overwrite = false;
        if (issues.Count > 0)
        {
            var report = await _workspace.InspectAsync(snapshot => TypedWorkbookStorage.InspectCompatibility(snapshot));
            var choice = await ShowCompatibilityReportAsync(report, saveMode: true,
                canOverwrite: current is not null && _workspace.Version?.Sha256 is not null && !migration && !requiresNewContainer);
            if (choice == CompatibilitySaveDecision.Cancel) return false;
            overwrite = choice == CompatibilitySaveDecision.Overwrite;
            discard = true; saveAs = !overwrite;
        }
        if (migration || requiresNewContainer) saveAs = true;
        LocalFileSelection? selected = !saveAs && current is not null ? new(current) : null;
        if (selected is null)
        {
            var extension = needsScripts || migration ? ".xlsq" : ".xlsx";
            FileDialogType[] types = needsScripts ? [new("QuickExcel 工作簿（含脚本）", ["*.xlsq"])] : [new("Excel 工作簿", ["*.xlsx"]), new("QuickExcel 工作簿", ["*.xlsq"])];
            selected = await _files.SaveAsync(new("保存工作簿", types, Path.GetFileNameWithoutExtension(current ?? "未命名") + (migration ? "-v2" : discard ? "-受支持副本" : "") + extension, extension));
        }
        if (selected is null) return false;
        try { await _workspace.SaveAsync(selected, discard, overwrite); _recentFiles.Remember(selected.Path); _status.Text = "已保存 " + Path.GetFileName(selected.Path) + (_calculationFailure is null ? "" : " · 部分结果待重算：" + _calculationFailure); return true; }
        catch (Exception ex) { Error(ex); return false; }
    }
    private async Task History(bool redo) { if (await CommitEditorsAsync()) TryEdit(() => redo ? _workspace.Editor.Redo() : _workspace.Editor.Undo()); }
    private async Task ClearAsync()
    {
        if (_grid.SelectedDrawingObject is { } target)
        {
            var snapshot = Document.Snapshot;
            if (await ExecuteAsync(target.Kind == DrawingObjectKind.Image ? new SetImageCommand(snapshot.ActiveSheetId, target.Index, null) : new SetChartCommand(snapshot.ActiveSheetId, target.Index, null), snapshot.Revision)) _grid.ClearDrawingObjectSelection();
            return;
        }
        if (await CommitEditorsAsync())
        {
            var view = CurrentView; view.RequireCurrentFilter(Document.Snapshot);
            TryEdit(() => _workspace.Editor.ClearCells(SheetIndex, SelectedAddresses().Where(p => !view.IsFilteredOut(Document.Snapshot, SheetIndex, p.Row) &&
                Document.Snapshot.Sheets[SheetIndex].Rows.GetValueOrDefault(p.Row)?.Hidden != true).Select(p => (p.Row, p.Column))));
        }
    }
    private async Task AddSheetAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var sheet = Document.AddSheet(); Document.Session.UpdateNavigation(sheet.Id, new(0, 0)); Synchronize();
    }
    private async Task RenameSheetAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var sheet = Document.Snapshot.Sheets[SheetIndex]; var name = await Dialogs.Text(this, "重命名工作表", "工作表名称", sheet.Name);
        if (name is not null) TryEdit(() => _workspace.Editor.Execute(new RenameSheetCommand(sheet.Id, name)));
    }
    private async Task DeleteSheetAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var sheet = Document.Snapshot.Sheets[SheetIndex];
        if (await Dialogs.Confirm(this, "删除工作表", $"删除“{sheet.Name}”？引用此表的公式将变为 #REF!，可撤销。")) TryEdit(() => _workspace.Editor.Execute(new DeleteSheetCommand(sheet.Id)));
    }
    private async Task ZoomAsync(int value, bool preserveDraft = false)
    {
        if (_grid.IsComposing || _formula.IsComposing || _grid.IsCommitPending || _formulaCommit is not null) return;
        if (preserveDraft || await CommitEditorsAsync()) TryEdit(() => _workspace.Editor.Execute(new SetSheetViewCommand(Document.Snapshot.ActiveSheetId, Document.Snapshot.Sheets[SheetIndex].View with { Zoom = value })));
    }
    private Task PlaceSheets(SheetNavigationPlacement placement) { _settings.Save(_settings.Value with { Sheets = placement }); ApplySheetPlacement(); return Task.CompletedTask; }
    private void ApplySheetPlacement()
    {
        var placement = _settings.Value.Sheets; var vertical = placement != SheetNavigationPlacement.Bottom;
        DockPanel.SetDock(_sheets, placement switch { SheetNavigationPlacement.Left => Dock.Left, SheetNavigationPlacement.Right => Dock.Right, _ => Dock.Bottom });
        _sheets.Width = vertical ? 160 : double.NaN;
        _sheets.ItemsPanel = new FuncTemplate<Panel?>(() => new StackPanel { Orientation = vertical ? Orientation.Vertical : Orientation.Horizontal });
    }
    private async Task CalculateAsync()
    {
        if (!await CommitEditorsAsync()) return;
        _calculationPending = true; await CalculatePendingCoreAsync(true);
    }
    private async void Shortcuts(object? sender, KeyEventArgs e)
    {
        if (_filterPopup?.IsOpen == true) return;
        if (e.Handled || _grid.IsComposing || _formula.IsComposing) return;
        if (e.Key == Key.Escape && _calculationCancellation is not null) { _calculationCancellation.Cancel(); e.Handled = true; return; }
        if (e.Key == Key.F6)
        {
            e.Handled = true;
            if (await CommitEditorsAsync())
            {
                Control[] stops = [_grid.CellEditor, _formula, _sheets, _menu];
                var at = Array.FindIndex(stops, control => control.IsKeyboardFocusWithin);
                stops[(at + (e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? stops.Length - 1 : 1)) % stops.Length].Focus();
            }
            return;
        }
        // Search editing and dropdown navigation own their keys; they must not clear or paste into worksheet cells.
        if (AutomationPickerOwnsKeyboard) return;
        if (e.Key == Key.Escape && _formatBrush is not null) { _formatBrush = null; _status.Text = "已取消格式刷。"; e.Handled = true; return; }
        if (e.Key == Key.Escape && _clipboardCells?.IsCut == true && !_grid.IsEditing)
        { _clipboardCells = null; _status.Text = "已取消剪切。"; e.Handled = true; return; }
        if (e.Key == Key.Delete && !_grid.IsEditing && !_formula.IsKeyboardFocusWithin) { e.Handled = true; await PerformAsync(ClearAsync); return; }
        var command = e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta);
        if (command && e.Key is Key.Z or Key.Y && (_grid.IsEditing || _formula.IsKeyboardFocusWithin || e.Source is TextBox box && box != _grid.CellEditor)) return;
        if (command && !_grid.IsEditing && !_formula.IsKeyboardFocusWithin && e.Key is Key.C or Key.X or Key.V or Key.D or Key.R)
        {
            e.Handled = true;
            await PerformAsync(e.Key switch { Key.C => () => CopyAsync(false), Key.X => () => CopyAsync(true), Key.V => () => PasteAsync(e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? "formulas" : "all"), Key.D => () => FillDirectionAsync(false), _ => () => FillDirectionAsync(true) });
            return;
        }
        Func<Task>? action = command ? e.Key switch { Key.S => async () => await SaveAsync(e.KeyModifiers.HasFlag(KeyModifiers.Shift)), Key.O => OpenAsync, Key.F => SearchAsync, Key.H => SearchAsync, Key.G => NavigateAsync,
            Key.N => async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } }, Key.W => async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowHome(); } }, Key.Z => () => History(false), Key.Y => () => History(true), _ => null } : e.Key == Key.F9 ? CalculateAsync : null;
        if (action is null) return; e.Handled = true; await PerformAsync(action);
    }
    private async void OnClosing(object? sender, WindowClosingEventArgs e)
    {
        if (_allowClose) return;
        CloseFilterPanel();
        e.Cancel = true; if (_closePending) return; _closePending = true;
        try
        {
            await _operation;
            if (!await ConfirmDiscardAsync()) return;
            await ClosePrintPreviewAsync();
            _allowClose = true; Close();
        }
        catch (Exception ex) { Error(ex); }
        finally { _closePending = false; }
    }
}
