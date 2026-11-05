using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
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
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow : Window
{
    private readonly WorkbookWorkspace _workspace;
    private readonly WorkbookServices _services;
    private readonly WorkbookGrid _grid = new();
    private readonly CompositionTextBox _formula = new() { AcceptsReturn = true, NewLine = "\n",
        TextWrapping = TextWrapping.Wrap, VerticalContentAlignment = VerticalAlignment.Top };
    private readonly TextBlock _address = new() { Width = 86, VerticalAlignment = VerticalAlignment.Center };
    private readonly TextBlock _status = new() { Text = UiText.Get("就绪"), Margin = new(12, 5) };
    private readonly TextBlock _noScriptsStatus = new() { Name = "NoScriptsStatus", Margin = new(12, 5), Foreground = Brushes.DarkGoldenrod };
    private readonly ListBox _sheets = new() { MinHeight = 0, Padding = new(0), SelectionMode = SelectionMode.Single };
    private readonly DockPanel _book = new();
    private readonly global::Avalonia.Controls.Grid _start = new() { Margin = new(40, 36), ColumnDefinitions = new("300,40,*") };
    private readonly StackPanel _recent = new() { Spacing = 6 };
    private readonly Menu _menu = new();
    private readonly List<MenuItem> _workbookMenuItems = [];
    private readonly TextBlock _readOnlyNotice = new() { Text = UiText.Get(DocumentSession.ReadOnlyMessage),
        Margin = new(12, 8), TextWrapping = TextWrapping.Wrap, Foreground = Brushes.DarkGoldenrod, IsVisible = false };
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

    private static DesktopSettings LoadSettings(string userDataDirectory)
    {
        var settings = new DesktopSettings(Path.Combine(userDataDirectory, "settings.json"));
        settings.Load();
        UiText.Initialize(settings.Value.Language);
        return settings;
    }

    public MainWindow(string userDataDirectory) : this(userDataDirectory, LoadSettings(userDataDirectory)) { }

    internal MainWindow(string userDataDirectory, DesktopSettings settings)
    {
        Title = "QuickExcel · Avalonia"; Width = 1280; Height = 820; MinWidth = 760; MinHeight = 480;
        FontFamily = new FontFamily("Segoe UI, Noto Sans CJK SC, sans-serif"); FontSize = 13;
        QuickExcel.Runtime.ScriptExecutionPolicy.Current.UserSetting = settings.Value.DisableScripts;
        var services = _services = DesktopServices.Create(); _workspace = new(services, new WorkbookFileAccess(services.Storage));
        _files = new AvaloniaFileDialogs(this);
        _recentFiles = new(Path.Combine(userDataDirectory, "recent-files.json")); _recentFiles.Load();
        _settings = settings;
        QuickExcel.Runtime.ScriptCompilation.SourceMapMode = settings.Value.ScriptSourceMap.ToString().ToLowerInvariant();
        _userDataDirectory = userDataDirectory;
        BuildView(); InstallEditingCommands(); InstallFormulaEditing(); InstallRibbon(); InstallCalculation(); InstallSheetCommands(); InstallSearch(); InstallDataCommands(); InstallValidation(); InstallCleaning(); InstallObjects(); InstallFileCommands(); InstallScriptCommands(); InstallPrintCommands();
        ArrangeRibbonTabs(); InstallScriptEvents(); GroupMainMenus();
        _grid.ResolveStyle = ConditionalPresentation.GetStyle;
        _grid.ResolveConditionalVisual = ConditionalPresentation.GetVisual;
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
            if (Program.SmokeDirectory is { } noScriptsDirectory && Program.SmokeMode == "no-scripts") { await RunNoScriptsSmokeAsync(noScriptsDirectory); return; }
            if (Program.SmokeDirectory is { } closingDirectory && Program.SmokeMode == "closing") { await RunWindowClosingSmokeAsync(closingDirectory); return; }
            if (Program.SmokeDirectory is { } commandsDirectory && Program.SmokeMode == "commands") { await RunDesktopCommandsSmokeAsync(commandsDirectory); return; }
            if (Program.SmokeDirectory is { } licensesDirectory && Program.SmokeMode == "licenses") { await RunLicenseSmokeAsync(licensesDirectory); return; }
            if (Program.SmokeDirectory is { } pasteDirectory && Program.SmokeMode == "paste-options") { await RunPasteOptionsSmokeAsync(pasteDirectory); return; }
            if (Program.SmokeDirectory is { } cellStylesDirectory && Program.SmokeMode == "cell-styles") { await RunCellStylesSmokeAsync(cellStylesDirectory); return; }
            if (Program.SmokeDirectory is { } localizationDirectory && Program.SmokeMode == "localization") { await RunLocalizationSmokeAsync(localizationDirectory); return; }
            if (Program.SmokeDirectory is { } sheetsDirectory && Program.SmokeMode == "sheets") { await RunSheetsSmokeAsync(sheetsDirectory); return; }
            if (Program.SmokeDirectory is { } dataToolsDirectory && Program.SmokeMode == "data-tools") { await RunDataToolsSmokeAsync(dataToolsDirectory); return; }
            if (Program.SmokeDirectory is { } borderDirectory && Program.SmokeMode == "borders") { await RunBorderSmokeAsync(borderDirectory); return; }
            if (Program.SmokeDirectory is { } rotationDirectory && Program.SmokeMode == "text-rotation") { await RunTextRotationSmokeAsync(rotationDirectory); return; }
            if (Program.SmokeDirectory is { } numberDirectory && Program.SmokeMode == "number-format") { await RunNumberFormatSmokeAsync(numberDirectory); return; }
            if (Program.SmokeDirectory is { } menuDirectory && Program.SmokeMode == "menus") { await RunMenuSmokeAsync(menuDirectory); return; }
            if (Program.SmokeDirectory is { } colorDirectory && Program.SmokeMode == "colors") { await RunColorSmokeAsync(colorDirectory); return; }
            if (Program.SmokeDirectory is { } conditionalDirectory && Program.SmokeMode == "conditional-format") { await RunConditionalFormatSmokeAsync(conditionalDirectory); return; }
            if (Program.SmokeDirectory is { } executionDirectory && Program.SmokeMode == "execution") { await RunExecutionSmokeAsync(executionDirectory); return; }
            if (Program.SmokeDirectory is { } clipboardDirectory && Program.SmokeMode == "clipboard-exchange") { await RunClipboardExchangeAsync(clipboardDirectory); return; }
            if (Program.SmokeDirectory is { } automationDirectory && Program.SmokeMode == "automation") { await RunAutomationSmokeAsync(automationDirectory); return; }
            if (Program.SmokeDirectory is { } featureDirectory && Program.SmokeMode == "ribbon-filter") { await RunRibbonFilterSmokeAsync(featureDirectory); return; }
            if (Program.SmokeDirectory is { } drawingDirectory && Program.SmokeMode == "drawing") { await RunDrawingSmokeAsync(drawingDirectory); return; }
            if (Program.SmokeDirectory is { } pageBreakDirectory && Program.SmokeMode == "page-break") { await RunPageBreakSmokeAsync(pageBreakDirectory); return; }
            if (Program.SmokeDirectory is { } recoveryBenchmark && Program.SmokeMode == "recovery-benchmark") { await RunRecoveryBenchmarkAsync(recoveryBenchmark); return; }
            if (Program.SmokeDirectory is { } recoveryDirectory && Program.SmokeMode == "recovery") { await RunRecoverySmokeAsync(recoveryDirectory); return; }
            if (Program.SmokeDirectory is { } contentDirectory && Program.SmokeMode == "cell-content") { await RunCellContentSmokeAsync(contentDirectory); return; }
            if (Program.SmokeDirectory is { } auditDirectory && Program.SmokeMode == "formula-audit") { await RunFormulaAuditSmokeAsync(auditDirectory); return; }
            if (Program.SmokeDirectory is { } arithmeticDirectory && Program.SmokeMode == "arithmetic-paste") { await RunArithmeticPasteSmokeAsync(arithmeticDirectory); return; }
            if (Program.SmokeDirectory is { } advancedDirectory && Program.SmokeMode == "advanced-filters") { await RunAdvancedFiltersSmokeAsync(advancedDirectory); return; }
            if (Program.SmokeDirectory is { } visibleDirectory && Program.SmokeMode == "visible-transfer") { await RunVisibleTransferSmokeAsync(visibleDirectory); return; }
            if (Program.SmokeDirectory is { } specialDirectory && Program.SmokeMode == "special-selection") { await RunSpecialSelectionSmokeAsync(specialDirectory); return; }
            if (Program.SmokeDirectory is { } functionDirectory && Program.SmokeMode == "common-functions") { await RunCommonFunctionsSmokeAsync(functionDirectory); return; }
            if (Program.SmokeDirectory is { } textDirectory && Program.SmokeMode == "text-exchange") { await RunTextExchangeSmokeAsync(textDirectory); return; }
            if (Program.SmokeDirectory is { } compatibilityDirectory && Program.SmokeMode == "compatibility") { await RunCompatibilitySmokeAsync(compatibilityDirectory); return; }
            if (Program.SmokeDirectory is { } smoke && Program.SmokeMode != "manual") { if (Program.SmokeMode == "grid") await RunGridSmokeAsync(smoke); else if (Program.SmokeMode == "benchmark") await RunGridBenchmarkAsync(smoke); else if (Program.SmokeMode == "web") await RunWebSmokeAsync(smoke); else if (Program.SmokeMode == "web-benchmark") await RunWebBenchmarkAsync(smoke); else await RunWorkspaceSmokeAsync(smoke); return; }
#endif
            if (Program.InitialFile is { } path) await PerformAsync(() => OpenPathAsync(path));
            StartRecovery(); await PerformAsync(() => ShowRecoveryAsync(true));
            if (Program.SettingsMigrationError is { } migrationError) _status.Text = migrationError;
        };
        Synchronize(); ShowHome();
    }
    private sealed record SheetItem(SheetId Id, string Name, string? TabColor) { public override string ToString() => Name; }
    private void BuildView()
    {
        MenuItem Item(string text, Func<Task> action, bool requiresWorkbook = true, bool mutation = true)
        {
            var item = CommandMenu(text, action, requiresWorkbook, mutation);
            if (requiresWorkbook) _workbookMenuItems.Add(item);
            return item;
        }
        var exportMenu = new MenuItem { Header = UiText.Get("导出"), Items =
        {
            Item(UiText.Get("导出值快照"), () => ExportAsync(false), mutation: false),
            Item(UiText.Get("导出可移植公式"), () => ExportAsync(true), mutation: false),
            Item(UiText.Get("导出 CSV / TSV"), ExportTextAsync, mutation: false)
        } };
        _workbookMenuItems.Add(exportMenu);
        _menu.Items.Add(new MenuItem { Header = UiText.Get("文件"), Items =
        {
            Item(UiText.Get("新建  Ctrl+N"), async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } }, false),
            Item(UiText.Get("打开  Ctrl+O"), OpenAsync, false), Item(UiText.Get("最近文件"), async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowHome(); } }, false),
            Item(UiText.Get("以只读方式打开…"), () => OpenAsync(readOnly: true), false),
            Item(UiText.Get("打开恢复副本"), RecoveryAsync, false), new Separator(),
            Item(UiText.Get("保存  Ctrl+S"), async () => await SaveAsync(false), mutation: false), Item(UiText.Get("另存为"), async () => await SaveAsync(true), mutation: false), new Separator(),
            exportMenu,
            Item(UiText.Get("打印预览 / PDF"), OpenPrintPreviewAsync, mutation: false), new Separator(), Item(UiText.Get("兼容性报告"), CompatibilityAsync, mutation: false), new Separator(),
            Item(UiText.Get("设置"), OpenSettingsAsync, false), new Separator(),
            Item(UiText.Get("关闭工作簿"), async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowHome(); } }, mutation: false), Item(UiText.Get("退出"), () => { Close(); return Task.CompletedTask; }, false)
        } });
        _menu.Items.Add(new MenuItem { Header = UiText.Get("编辑"), Items =
        { Item(UiText.Get("撤销  Ctrl+Z"), () => History(false)), Item(UiText.Get("重做  Ctrl+Y"), () => History(true)) } });
        _menu.Items.Add(new MenuItem { Header = UiText.Get("工作表"), Items =
        { Item(UiText.Get("添加工作表"), AddSheetAsync), Item(UiText.Get("重命名"), RenameSheetAsync), Item(UiText.Get("删除工作表"), DeleteSheetAsync) } });
        _menu.Items.Add(new MenuItem { Header = UiText.Get("视图"), Items =
        { Item("100%", () => ZoomAsync(100)), Item("125%", () => ZoomAsync(125)), Item("150%", () => ZoomAsync(150)),
          Item(UiText.Get("标签在底部"), () => PlaceSheets(SheetNavigationPlacement.Bottom), mutation: false), Item(UiText.Get("标签在左侧"), () => PlaceSheets(SheetNavigationPlacement.Left), mutation: false), Item(UiText.Get("标签在右侧"), () => PlaceSheets(SheetNavigationPlacement.Right), mutation: false) } });
        DockPanel.SetDock(_readOnlyNotice, Dock.Top); _book.Children.Add(_readOnlyNotice);
        _book.Children.Add(BuildWorksheetLayout()); ApplySheetPlacement();
        BuildHome();
        var content = new global::Avalonia.Controls.Grid(); content.Children.Add(_book); content.Children.Add(_start);
        var root = new DockPanel(); DockPanel.SetDock(_menu, Dock.Top);
        var footer = new DockPanel();
        UpdateScriptExecutionPolicy();
        DockPanel.SetDock(_noScriptsStatus, Dock.Right); footer.Children.Add(_noScriptsStatus);
        DockPanel.SetDock(_selectionStatistics, Dock.Right); footer.Children.Add(_selectionStatistics); DockPanel.SetDock(_pivotStatus,Dock.Right); footer.Children.Add(_pivotStatus); DockPanel.SetDock(_recoveryStatus, Dock.Right); footer.Children.Add(_recoveryStatus); footer.Children.Add(_status);
        DockPanel.SetDock(footer, Dock.Bottom); root.Children.Add(_menu); root.Children.Add(footer); root.Children.Add(content);
        Content = _execution = new ExecutionOverlay(this, root, () => _grid.IsComposing || _formula.IsComposing);
    }
    private LayoutGrid BuildWorksheetLayout()
    {
        var layout = new LayoutGrid { RowDefinitions = new("38,5,*") };
        var formulaHeight = layout.RowDefinitions[0];
        formulaHeight.MinHeight = formulaHeight.Height.Value;
        const double minSheetHeight = 120;
        layout.SizeChanged += (_, e) => formulaHeight.MaxHeight = Math.Max(formulaHeight.MinHeight,
            e.NewSize.Height - layout.RowDefinitions[1].Height.Value - minSheetHeight);

        var formulaRow = new DockPanel { Margin = new(8, 4, 8, 2) };
        DockPanel.SetDock(_address, Dock.Left);
        ScrollViewer.SetHorizontalScrollBarVisibility(_formula, ScrollBarVisibility.Disabled);
        ScrollViewer.SetVerticalScrollBarVisibility(_formula, ScrollBarVisibility.Auto);
        formulaRow.Children.Add(_address);
        formulaRow.Children.Add(_formula);
        layout.Children.Add(formulaRow);

        var splitter = new GridSplitter { ResizeDirection = GridResizeDirection.Rows, ResizeBehavior = GridResizeBehavior.PreviousAndNext,
            HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch, Background = Brush.Parse("#E0E0E0") };
        AutomationProperties.SetName(splitter, UiText.Get("调整公式栏高度"));
        ToolTip.SetTip(splitter, UiText.Get("上下拖动可调整公式栏高度"));
        LayoutGrid.SetRow(splitter, 1);
        layout.Children.Add(splitter);

        var sheet = new DockPanel();
        sheet.Children.Add(BuildSheetStrip());
        DockPanel.SetDock(_validationPanel, Dock.Top);
        sheet.Children.Add(_validationPanel);
        sheet.Children.Add(_grid);
        LayoutGrid.SetRow(sheet, 2);
        layout.Children.Add(sheet);
        return layout;
    }
    private void ShowHome()
    {
        if (_book.IsVisible) Document.Session.RaiseScriptEvent("workbook.deactivate");
        _book.IsVisible = false; _start.IsVisible = true;
        UpdateWorkbookMenus();
        RefreshRecentFiles();
    }
    private void RefreshRecentFiles()
    {
        _recentFiles.Load();
        FilterRecentFiles();
    }
    private void FilterRecentFiles()
    {
        var query = _recentSearch.Text?.Trim() ?? "";
        _recent.Children.Clear();
        foreach (var entry in _recentFiles.Entries.Where(entry => entry.Path.Contains(query, StringComparison.OrdinalIgnoreCase)))
        {
            var path = entry.Path;
            var label = new StackPanel { Spacing = 5 };
            label.Children.Add(new TextBlock { Text = Path.GetFileName(path), FontSize = 16, FontWeight = FontWeight.Bold, Foreground = Brushes.Black, TextTrimming = TextTrimming.CharacterEllipsis });
            label.Children.Add(new TextBlock { Text = Path.GetDirectoryName(path), FontSize = 12, Foreground = Brushes.Gray, TextTrimming = TextTrimming.CharacterEllipsis });
            var button = new Button { Name = "OpenRecentFile", Tag = path, Content = label, HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Padding = new(16, 12), MinHeight = 70, Background = Brushes.White, BorderBrush = new SolidColorBrush(Color.Parse("#E1E5E3")), BorderThickness = new(1), CornerRadius = new(6) }; ToolTip.SetTip(button, path);
            AutomationProperties.SetName(button, Path.GetFileName(path) + "，" + Path.GetDirectoryName(path));
            button.Click += async (_, _) => await PerformAsync(() => OpenRecentFileAsync(path));
            var row = new LayoutGrid { ColumnDefinitions = new("*,Auto") };
            row.Children.Add(button);
            var pin = CreateRecentPinButton(entry);
            LayoutGrid.SetColumn(pin, 1); row.Children.Add(pin); _recent.Children.Add(row);
        }
        if (_recent.Children.Count == 0) _recent.Children.Add(new TextBlock
        {
            Text = _recentFiles.Paths.Count == 0 ? UiText.Get("暂无历史文件，打开工作簿后会显示在这里。") : UiText.Get("没有匹配的历史文件。"),
            Foreground = Brushes.Gray, TextWrapping = TextWrapping.Wrap, Margin = new(0, 12)
        });
    }
    private void ShowBook() { if (!_book.IsVisible) Document.Session.RaiseScriptEvent("workbook.activate"); _start.IsVisible = false; _book.IsVisible = true; UpdateWorkbookMenus(); }
    private void UpdateWorkbookMenus()
    {
        foreach (var item in _workbookMenuItems) item.IsEnabled = _book.IsVisible;
        foreach (var menu in _menu.Items.OfType<MenuItem>().Skip(1)) menu.IsEnabled = _book.IsVisible;
        UpdateReadOnlyMenus(_menu);
        if (_sheets.ContextMenu is { } tabs) UpdateReadOnlyMenus(tabs);
        if (Document.IsReadOnly)
        {
            ((MenuItem)_menu.Items[2]!).IsEnabled = false;
            var scripts = _menu.Items.OfType<MenuItem>().FirstOrDefault(menu => Equals(menu.Header, UiText.Get("脚本")));
            if (scripts is not null) scripts.IsEnabled = false;
        }
    }
    private void UpdateReadOnlyMenus(ItemsControl menu)
    {
        foreach (var item in menu.Items.OfType<MenuItem>())
        {
            if (item.Classes.Contains("workbook-mutation")) item.IsEnabled = _book.IsVisible && !Document.IsReadOnly;
            UpdateReadOnlyMenus(item);
        }
    }
    private void Synchronize()
    {
        if (_synchronizing) return;
        if (_sheetPress is { } press && !IsCurrentSheetPress(press)) EndSheetGesture();
        if (_filterContextCurrent?.Invoke() == false) CloseFilterPanel();
        if (_chartFilterContextCurrent?.Invoke() == false) CloseChartFilter();
        if (_lastPaste is not null && !IsPasteOptionsCurrent()) DismissPasteOptions();
        ObserveDocument();
        RefreshScriptApiMigrationButton();
        _synchronizing = true;
        try
        {
            _book.IsEnabled = !NonExecutionBusy; _menu.IsEnabled = !NonExecutionBusy;
            UpdateWorkbookMenus();
            if (_boundSession != Document.Session.SessionId || _boundSheet != Document.Snapshot.ActiveSheetId)
            {
                _formulaDirty = false; _formulaOwner = null;
                _boundSession = Document.Session.SessionId; _boundSheet = Document.Snapshot.ActiveSheetId;
                _grid.Bind(Document.Snapshot, SheetIndex, WorkbookLimits.MaxRows, WorkbookLimits.MaxColumns, Document.Session.SessionId, Document.IsReadOnly);
            }
            else _grid.Refresh(Document.Snapshot);
            var items = Document.Snapshot.Sheets.Where(s => s.Visibility == SheetVisibility.Visible).Select(s => new SheetItem(s.Id, s.Name, s.TabColor)).ToArray();
            if (_sheets.ItemsSource is not SheetItem[] prior || !prior.SequenceEqual(items)) _sheets.ItemsSource = items;
            _sheets.SelectedItem = items.FirstOrDefault(s => s.Id == Document.Snapshot.ActiveSheetId);
            _addSheetButton.IsEnabled = !Document.IsReadOnly && Document.Snapshot.FileState?.FormulaReadOnly != true && Document.Snapshot.Sheets.Length < WorkbookLimits.MaxSheets;
            _readOnlyNotice.IsVisible = Document.IsReadOnly;
        }
        finally { _synchronizing = false; }
        UpdateFormula(); UpdateTitle(); UpdateRibbon(); UpdateValidation(); QueueStatistics(); RefreshPrintStatus(); RefreshScriptConsole(); UpdatePivotStatus();
    }
    private void UpdateTitle()
    {
        Title = (Document.IsDirty || _scriptPage?.Drafts.IsDirty == true ? "* " : "") + (_workspace.FilePath is { } path ? Path.GetFileName(path) : UiText.Get("未命名.xlsq")) + (Document.IsReadOnly ? UiText.Get(" [只读]") : "") + " — QuickExcel";
        _address.Text = WorkbookGrid.ColumnName(_grid.Current.Column) + (_grid.Current.Row + 1);
        var sheet = Document.Snapshot.Sheets[SheetIndex];
        var value = Document.GetCellValue(SheetIndex, _grid.Current.Row, _grid.Current.Column);
        var readOnly = _grid.IsCellReadOnly(_grid.Current);
        AutomationProperties.SetName(_grid, UiText.Format($"工作表 {sheet.Name}，{WorkbookGrid.ColumnName(_grid.Current.Column)} 列，第 {_grid.Current.Row + 1} 行，值 {value}，选中 {_grid.SelectionRanges.Sum(r => r.Count)} 格，{(readOnly ? UiText.Get("只读") : UiText.Get("可编辑"))}"));
        AutomationProperties.SetName(_grid.CellEditor, UiText.Format($"单元格 {_address.Text}，{(readOnly ? UiText.Get("只读") : UiText.Get("编辑"))}"));
        AutomationProperties.SetName(_formula, UiText.Get("公式栏 ") + _address.Text); AutomationProperties.SetName(_sheets, UiText.Get("工作表列表"));
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
        if (Document.IsReadOnly) return false;
        if (_workspace.IsBusy || _execution.IsRunning) return false;
        try
        {
            var snapshot = Document.Snapshot; var sheet = snapshot.Sheets[index];
            var command = new SetCellsCommand([new(new(sheet.Id, owner.Row, owner.Column), CellInput.ParseForCell(text, snapshot, sheet, owner.Row, owner.Column))]);
            await _workspace.Editor.ExecuteAsync(command, ConfirmValidationAsync, snapshot.Revision);
            QueueCellEvent("cell.commit", sheet.Id, owner);
            Synchronize(); _status.Text = UiText.Get("已修改"); return true;
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
        if (Document.IsReadOnly) { _status.Text = UiText.Get(DocumentSession.ReadOnlyMessage); return false; }
        if (_workspace.IsBusy || _execution.IsRunning) return false;
        try { edit(); Synchronize(); _status.Text = UiText.Get("已修改"); return true; }
        catch (Exception ex) { Error(ex); return false; }
    }
    private void Error(Exception ex) => _status.Text = ex.Message;
    private Task PerformAsync(Func<Task> action)
    {
        if (_working || _execution.IsRunning) return Task.CompletedTask;
        // Closing waits for error reporting and cleanup, not the raw action task:
        // an already-reported open/save failure must not veto every later close.
        return _operation = PerformCoreAsync(action);
    }
    private async Task PerformCoreAsync(Func<Task> action)
    {
        _working = true;
        try { await using var hold = _scriptPage?.HoldEdits(); await action(); }
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
        if (accepted) { await CloseScriptWindowAsync(); await ClearRecoveryAsync(); }
        return accepted;
    }
    private Task OpenAsync() => OpenAsync(readOnly: false);
    private async Task OpenAsync(bool readOnly)
    {
        if (!await CommitEditorsAsync()) return;
        var selected = await _files.OpenAsync(new(UiText.Get("打开工作簿"), [new(UiText.Get("所有支持的文件"), ["*.xlsq", "*.xlsx", "*.csv"]), new(UiText.Get("CSV（UTF-8 文本）"), ["*.csv"])]));
        if (selected is not null && await ConfirmDiscardAsync()) await OpenPathAsync(selected.Path, readOnly);
    }
    private async Task OpenPathAsync(string path, bool readOnly = false)
    {
        await _workspace.OpenAsync(new(path), readOnly); _recentFiles.Remember(path); ShowBook();
        _status.Text = UiText.Get(Document.IsReadOnly ? "已以只读方式打开 " : "已打开 ") + Path.GetFileName(path);
        if (Path.GetExtension(path).Equals(".csv", StringComparison.OrdinalIgnoreCase))
            _status.Text = UiText.Get("已按文本导入 CSV。保存时将写入新的工作簿。");
        if (!Document.IsReadOnly && DesktopServices.SupportsScripts && !QuickExcel.Runtime.ScriptExecutionPolicy.Current.IsDisabled && Document.Snapshot.ScriptModules.Signature is not null)
        {
            var modules = Document.Snapshot.ScriptModules;
            try
            {
                var trust = await Task.Run(() => ScriptTrust.Verify(modules));
                if (ReferenceEquals(modules, Document.Snapshot.ScriptModules) && trust.CanAutoEnable)
                {
                    if (!modules.CanExecute) _status.Text = UiText.Get(ScriptModuleSet.MigrationMessage);
                    else if (trust.ExecutionPolicyAccepted) await EnableVerifiedScriptsAsync();
                    else _status.Text = UiText.Get("脚本签名有效；尚未确认当前引擎策略，请在脚本页显式应用一次。");
                }
            }
            catch (QuickExcel.Security.UserProtectionUnavailableException ex) { _status.Text = UiText.Get("文件已打开；无法读取持久信任，脚本保持停用：") + ex.Message; }
        }
    }
    private async Task<bool> SaveAsync(bool saveAs)
    {
        if (!_book.IsVisible) return false;
        await using var hold = _scriptPage?.HoldEdits();
        if (!await FlushEditsAsync()) return false;
        if (Document.IsReadOnly) saveAs = true;
        else
        {
            Document.Session.RaiseScriptEvent("workbook.beforeSave"); await DispatchEventsAsync();
            if (_settings.Value.Calculation == DesktopCalculationMode.Manual) _calculationPending = true;
            await CalculatePendingCoreAsync(_settings.Value.Calculation == DesktopCalculationMode.Manual);
        }
        if (_calculationPending && _settings.Value.Calculation != DesktopCalculationMode.Automatic)
        {
            _status.Text = UiText.Get("保存前重算未完成，本次保存已取消。") + _calculationFailure;
            return false;
        }
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
                canOverwrite: !Document.IsReadOnly && current is not null && _workspace.Version?.Sha256 is not null && !migration && !requiresNewContainer);
            if (choice == CompatibilitySaveDecision.Cancel) return false;
            overwrite = choice == CompatibilitySaveDecision.Overwrite;
            discard = true; saveAs = !overwrite;
        }
        if (migration || requiresNewContainer || string.Equals(Path.GetExtension(current), ".csv", StringComparison.OrdinalIgnoreCase)) saveAs = true;
        LocalFileSelection? selected = !saveAs && current is not null ? new(current) : null;
        if (selected is null)
        {
            var preferXlsx = !needsScripts && !migration && string.Equals(Path.GetExtension(current), ".xlsx", StringComparison.OrdinalIgnoreCase);
            var extension = preferXlsx ? ".xlsx" : ".xlsq";
            FileDialogType[] types = needsScripts ? [new(UiText.Get("QuickExcel 工作簿（含脚本）"), ["*.xlsq"])] : [new(UiText.Get("QuickExcel 工作簿"), ["*.xlsq"]), new(UiText.Get("Excel 工作簿"), ["*.xlsx"])];
            selected = await _files.SaveAsync(new(UiText.Get("保存工作簿"), types, Path.GetFileNameWithoutExtension(current ?? UiText.Get("未命名")) + (migration ? "-v2" : discard ? UiText.Get("-受支持副本") : Document.IsReadOnly ? UiText.Get("-副本") : "") + extension, extension, preferXlsx ? 1 : 0));
        }
        if (selected is null) return false;
        try { await _workspace.SaveAsync(selected, discard, overwrite); if (!Document.IsDirty) await ClearRecoveryAsync(); _recentFiles.Remember(selected.Path); _status.Text = UiText.Get("已保存 ") + Path.GetFileName(selected.Path) + (_calculationFailure is null ? "" : UiText.Get(" · 部分结果待重算：") + _calculationFailure); return true; }
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
    private async Task ZoomAsync(int value, bool preserveDraft = false)
    {
        if (_grid.IsComposing || _formula.IsComposing || _grid.IsCommitPending || _formulaCommit is not null) return;
        if (preserveDraft || await CommitEditorsAsync()) TryEdit(() => _workspace.Editor.Execute(new SetSheetViewCommand(Document.Snapshot.ActiveSheetId, Document.Snapshot.Sheets[SheetIndex].View with { Zoom = value })));
    }
    private Task PlaceSheets(SheetNavigationPlacement placement) { _settings.Save(_settings.Value with { Sheets = placement }); ApplySheetPlacement(); return Task.CompletedTask; }
    private async Task CalculateAsync()
    {
        if (Document.IsReadOnly) { _status.Text = UiText.Get(DocumentSession.ReadOnlyMessage); return; }
        if (!await CommitEditorsAsync()) return;
        _calculationPending = true; await CalculatePendingCoreAsync(true);
    }
    private async void Shortcuts(object? sender, KeyEventArgs e)
    {
        if (e.Handled || _grid.IsComposing || _formula.IsComposing) return;
        if (!_book.IsVisible)
        {
            if (e.Key == Key.F12 && e.KeyModifiers == KeyModifiers.Control) { e.Handled = true; await PerformAsync(OpenAsync); return; }
            if (!e.Handled && (e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta)))
            {
                if (e.Key == Key.O) { e.Handled = true; await PerformAsync(OpenAsync); }
                else if (e.Key == Key.N)
                {
                    e.Handled = true;
                    await PerformAsync(async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } });
                }
            }
            return;
        }
        if (e.Key == Key.Escape && _grid.IsSelectionDragging && !_grid.IsComposing && !_formula.IsComposing)
        { _grid.CancelInputGesture(); e.Handled = true; return; }
        if (e.Key == Key.Escape && _sheetPress is not null) { EndSheetGesture(); e.Handled = true; return; }
        if (e.Key == Key.Escape && _lastPaste is not null) { DismissPasteOptions(); e.Handled = true; return; }
        if (_pasteOptionsPopup?.IsOpen == true) return;
        if (_filterPopup?.IsOpen == true || _chartFilterPopup?.IsOpen == true) return;
        if (e.Handled || _grid.IsComposing || _formula.IsComposing) return;
        if (Document.IsReadOnly && ((e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta)) &&
                e.Key is Key.X or Key.V or Key.D or Key.R or Key.E or Key.Z or Key.Y or Key.D1 or Key.NumPad1 or Key.B or Key.I or Key.U or Key.D5 or Key.NumPad5 ||
                e.Key is Key.Delete or Key.Back or Key.F2 or Key.F9))
        {
            // Native text controls retain their own copy/select/navigation behavior.
            if (e.Source is TextBox text && text != _grid.CellEditor && !text.IsReadOnly) return;
            e.Handled = true; _status.Text = UiText.Get(DocumentSession.ReadOnlyMessage); return;
        }
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
        if (ExcelShortcut(e) is { } excelAction) { e.Handled = true; await PerformAsync(excelAction); return; }
        if (e.Key == Key.Escape && _formatBrush is not null) { _formatBrush = null; _status.Text = UiText.Get("已取消格式刷。"); e.Handled = true; return; }
        if (e.Key == Key.Escape && _clipboardCells?.IsCut == true && !_grid.IsEditing)
        { _clipboardCells = null; _status.Text = UiText.Get("已取消剪切。"); e.Handled = true; return; }
        if (e.Key == Key.Delete && !_grid.IsEditing && !_formula.IsKeyboardFocusWithin) { e.Handled = true; await PerformAsync(ClearAsync); return; }
        var command = e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta);
        if (command && e.Key == Key.E && e.KeyModifiers is KeyModifiers.Control or KeyModifiers.Meta && !_grid.IsEditing && !_formula.IsKeyboardFocusWithin &&
            (e.Source is not TextBox textInput || textInput == _grid.CellEditor))
        { e.Handled = true; await PerformAsync(FlashFillAsync); return; }
        if (command && e.Key is Key.Z or Key.Y && (_grid.IsEditing || _formula.IsKeyboardFocusWithin || e.Source is TextBox box && box != _grid.CellEditor)) return;
        if (command && !_grid.IsEditing && !_formula.IsKeyboardFocusWithin && e.Key is Key.C or Key.X or Key.V or Key.D or Key.R)
        {
            e.Handled = true;
            await PerformAsync(e.Key switch { Key.C => () => CopyAsync(false), Key.X => () => CopyAsync(true), Key.V => () => PasteAsync(e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? "formulas" : "all"), Key.D => () => FillDirectionAsync(false), _ => () => FillDirectionAsync(true) });
            return;
        }
        Func<Task>? action = command ? e.Key switch { Key.D1 or Key.NumPad1 => CellFormatAsync, Key.S => async () => await SaveAsync(e.KeyModifiers.HasFlag(KeyModifiers.Shift)), Key.O => OpenAsync, Key.F => SearchAsync, Key.H => SearchAsync, Key.G => NavigateAsync,
            Key.N => async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } }, Key.W => async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowHome(); } }, Key.Z => () => History(false), Key.Y => () => History(true), _ => null } : e.Key == Key.F9 ? CalculateAsync : null;
        if (action is null) return; e.Handled = true; await PerformAsync(action);
    }
    private async void OnClosing(object? sender, WindowClosingEventArgs e)
    {
        if (_allowClose) return;
        CloseFilterPanel();
        CloseChartFilter();
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
