using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow : Window
{
    private readonly WorkbookServices _services = DesktopComposition.CreateServices();
    private WorkbookDocument _document = null!;
    private DocumentViewModel _viewModel = null!;
    private EditorController _editor = null!;
    private readonly FastGridSurface _grid;
    private IGridSurface GridSurface => _grid;
    private readonly CellEditState _editState = new();
    private int _sheetIndex;
    private int _rowCount = 100;
    private int _columnCount = 12;
    private int _selectedRow;
    private int _selectedColumn;
    private bool _refreshing;
    private bool _busy;
    private CancellationTokenSource? _calculationCancellation;
    private bool _closingApproved;
    private bool _closeRequestPending;
    private bool _windowClosed;
    private bool _smokeMode;
    private MessageBoxResult? _smokeDiscardDecision;
    private bool _calculationPending;
    private bool _suppressAutoCalculation;
    private bool _cellCommitFailed;
    private bool _updatingFormula;
    private bool _formulaPending => _editState.Formula != null;
    private int _formulaSheet => _editState.Formula!.Sheet;
    private int _formulaRow => _editState.Formula!.Row;
    private int _formulaColumn => _editState.Formula!.Column;
    private string? _filePath;
    private string _displayName = "未命名.xlsq";

    public MainWindow()
    {
        InitializeComponent();
        _grid = new FastGridSurface(CommitGridValue) { HandlesKeyboard = false };
        SheetGrid.Children.Add(_grid.Element);
        _grid.BeforeNavigate = () => !_busy && CanCommitComposition() && CommitFormulaBar();
        _grid.BeforeBeginEdit = () => !_busy && CanCommitComposition() && CommitFormulaBar();
        _grid.LatestSnapshot = () => _document.Snapshot;
        _grid.SelectionChanged += Grid_CurrentCellChanged;
        _grid.EditingChanged += () => { if (_grid.IsEditing) _editState.BeginGrid(); else _editState.EndGrid(); };
        InstallInputHandlers();
        InstallWorkspaceHandlers();
        InstallDocument(WorkbookDocument.CreateBlank(_services), null, "未命名.xlsq", disposePrevious: false);
        _hasWorkbook = false;
        ShowStartPage();
    }

    public async Task StartAsync(string[] args)
    {
        if (args.Length == 2 && args[0] == "--ui-validation")
        {
            var folder = Path.GetFullPath(args[1]);
            _recentFiles = new(Path.Combine(folder, "recent-files.json"));
            _settings = new(Path.Combine(folder, "settings.json"));
            _scriptTrust = new(Path.Combine(folder, "trusted-script-certificates.json"));
            args = [];
        }
        if (args.Length > 0 && args[0] is "--smoke-test" or "--interaction-smoke" or "--workspace-smoke" or "--product-smoke" or "--ribbon-smoke" or "--signature-smoke" or "--grid-benchmark" or "--resize-benchmark" or "--data-operations-smoke" or "--horizontal-scroll-benchmark" or "--fast-grid-benchmark" or "--fast-grid-drag-benchmark" or "--grid-integration-smoke" or "--xlsx-rendering-smoke")
        {
            _smokeMode = true;
            _recentFiles = new(Path.Combine(args.Length > 1 ? args[1] : AppContext.BaseDirectory, "recent-files.json"));
            _settings = new(Path.Combine(args.Length > 1 ? args[1] : AppContext.BaseDirectory, "settings.json"));
            _scriptTrust = new(Path.Combine(args.Length > 1 ? args[1] : AppContext.BaseDirectory, "trusted-script-certificates.json"));
            ShowWorkbookPage();
            try
            {
                var folder = args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "smoke");
                if (args[0] == "--xlsx-rendering-smoke") await RunXlsxRenderingSmokeAsync(folder, args.ElementAtOrDefault(2));
                else if (args[0] == "--fast-grid-drag-benchmark") await RunFastGridDragBenchmarkAsync(folder);
                else if (args[0] == "--grid-integration-smoke") await RunGridIntegrationSmokeAsync(folder);
                else if (args[0] == "--horizontal-scroll-benchmark" && args.Length > 2) await RunHorizontalScrollStressAsync(folder, args[2]);
                else if (args[0] == "--horizontal-scroll-benchmark") await RunHorizontalScrollBenchmarkAsync(folder);
                else if (args[0] == "--fast-grid-benchmark") await RunFastGridBenchmarkAsync(folder);
                else if (args[0] == "--data-operations-smoke") await RunDataOperationsSmokeAsync(folder);
                else if (args[0] == "--signature-smoke") await RunSignatureSmokeAsync(folder);
                else if (args[0] == "--ribbon-smoke") await RunRibbonSmokeAsync(folder);
                else if (args[0] == "--resize-benchmark") await MeasureResizeAsync(folder);
                else if (args[0] == "--grid-benchmark") await RunGridBenchmarkAsync(folder);
                else if (args[0] == "--interaction-smoke") await RunInteractionSmokeAsync(folder);
                else if (args[0] == "--workspace-smoke") await RunWorkspaceSmokeAsync(folder);
                else if (args[0] == "--product-smoke") await RunProductSmokeAsync(folder);
                else
                {
                    await Task.Run(() => QuickExcel.Runtime.ScriptWorkerRuntimeFactory.RecoverAbandonedSandboxes());
                    await RunSmokeTestAsync(folder);
                }
                Environment.ExitCode = 0;
            }
            catch (Exception ex)
            {
                Environment.ExitCode = 1;
                var folder = args.Length > 1 ? args[1] : AppContext.BaseDirectory;
                Directory.CreateDirectory(folder);
                File.WriteAllText(Path.Combine(folder, "smoke-error.txt"), ex.ToString());
            }
            _closingApproved = true;
            Close();
            return;
        }
        _recentFiles.Load();
        _settings.Load();
        ApplyPreferences(_settings.Value);
        ShowStartPage();
        if (args.Length > 0) await OpenPathAsync(args[0]);
        try { await Task.Run(() => QuickExcel.Runtime.ScriptWorkerRuntimeFactory.RecoverAbandonedSandboxes()); }
        catch (ScriptRuntimeFailureException ex) { LogText.Text += "\n" + ex.Message; }
    }

    private void InstallDocument(WorkbookDocument document, string? path, string displayName, bool disposePrevious = true, FileVersion? fileVersion = null)
    {
        if (disposePrevious) _document.Dispose();
        _document = document;
        _verifiedScriptModules = null;
        _signatureVerification = null;
        _viewModel = new DocumentViewModel(document);
        _editor = new EditorController(document);
        DataContext = _viewModel;
        _filePath = path != null && Path.GetExtension(path).Equals(".xlsq", StringComparison.OrdinalIgnoreCase) ? path : null;
        _versionPath = _filePath is null ? null : Path.GetFullPath(_filePath);
        _fileVersion = fileVersion;
        _sheetViews.Clear();
        _displayName = displayName;
        _sheetIndex = _selectedRow = _selectedColumn = 0;
        _rowCount = 100;
        _columnCount = 12;
        _editState.Reset();
        _calculationPending = false;
        _refreshing = true;
        LoadScriptModules();
        ActionPicker.ItemsSource = null;
        SheetPicker.ItemsSource = document.Snapshot.Sheets.Select(sheet => sheet.Name).ToList();
        SheetPicker.SelectedIndex = 0;
        _refreshing = false;
        RefreshSheet(true);
        _hasWorkbook = true;
        ShowWorkbookPage();
        if (IsLoaded) GridSurface.Select(new(0, 0));
        UpdateChrome();
        ScriptState.Text = "● 脚本未启用";
        LogText.Text = "脚本未启用。";
        StatusText.Text = "就绪";
        if (document.Snapshot.FileState?.RequiresMigration == true)
            LogText.Text += "\n已导入 v1：历史文本不会被猜测转换类型。首次保存必须另存为 v2，原文件保留。";
        if (document.Snapshot.FileState is { CompatibilityIssues.Length: > 0 } fileState)
            LogText.Text += $"\n文件已打开。有 {fileState.CompatibilityIssues.Length} 项格式或附加信息尚未完整支持，部分显示可能与 Excel 不同。" +
                "\n原始内容已保留；修改后保存或导出前会提示可能丢失的内容。详情见文件菜单中的兼容性报告。";
        if (document.Snapshot.Sheets.Any(s => s.Cells.Values.Any(c => c.Content.Kind == CellContentKind.Formula)))
            LogText.Text += "\n文件中的公式缓存不等于已重算结果；内置不可计算的函数和语法可在兼容性报告中按单元格查看。";
        if (!document.Snapshot.ScriptModules.CanExecute)
            LogText.Text += "\n多模块源码已保留；执行需要显式声明脚本 API v2。";
        else if (document.Snapshot.ScriptModules.ScriptApiVersion == 2)
            LogText.Text += $"\n包内模块：{document.Snapshot.ScriptModules.Modules.Count} 个；编辑器显示首入口 {document.Snapshot.ScriptModules.EntryModules[0]}。应用时执行全部声明入口。";
        if (document.Snapshot.Sheets.Any(sheet => sheet.UsedColumnCount > 256))
            StatusText.Text = "部分列未显示，可在文件菜单查看兼容性报告。";
    }

    private void RefreshSheet(bool rebuild = false)
    {
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        _grid.View = CurrentView;
        if (rebuild)
        {
            var ranges = sheet.MergedCells.Select(CellRectangle.Parse).ToArray();
            var lastRow = Math.Max(sheet.UsedRowCount, Math.Max(sheet.Images.Select(i => i.RowIndex + i.RowSpan).DefaultIfEmpty().Max(), ranges.Select(m => m.Bottom + 1).DefaultIfEmpty().Max()));
            var lastColumn = Math.Max(sheet.UsedColumnCount, Math.Max(sheet.Images.Select(i => i.ColumnIndex + i.ColumnSpan).DefaultIfEmpty().Max(), ranges.Select(m => m.Right + 1).DefaultIfEmpty().Max()));
            _rowCount = Math.Clamp(Math.Max(_rowCount, lastRow + 20), 100, WorkbookLimits.MaxRows);
            _columnCount = Math.Clamp(Math.Max(_columnCount, lastColumn + 2), 12, 256);
            _refreshing = true;
            try { GridSurface.Bind(_document.Snapshot, _sheetIndex, _rowCount, _columnCount); }
            finally { _refreshing = false; }
        }
        else GridSurface.Refresh(_document.Snapshot);
        RefreshWorkspaceView();
        if (!_editState.GridEditing && !_formulaPending && CurrentView.FilterRange is { } filter &&
            CurrentView.IsFilteredOut(_document.Snapshot, _sheetIndex, _selectedRow))
            GridSurface.Select(new(filter.Top, _selectedColumn));
        UpdateFormulaBar();
        _viewModel.Refresh();
        var stale = _viewModel.StaleCacheCount;
        SelectionText.Text = $"{sheet.Cells.Count:N0} 格" + (stale > 0 ? $" · {stale} 个结果待重算" : "");
    }

    private void UpdateChrome()
    {
        RefreshSignatureState();
        _viewModel.Refresh();
        FileTitle.Text = _displayName + (_viewModel.IsDirty ? "  •" : "");
        Title = $"{(_viewModel.IsDirty ? "* " : "")}{_displayName} — QuickExcel";
        FileTitle.ToolTip = (_filePath ?? _displayName) + (_viewModel.DiscardedHistoryCount > 0 ? $" · 已释放 {_viewModel.DiscardedHistoryCount} 条较早撤销记录" : "");
        UndoButton.IsEnabled = _viewModel.CanUndo && !_busy;
        RedoButton.IsEnabled = _viewModel.CanRedo && !_busy;
        UpdateRibbonState();
        ActionsPanel.Visibility = _document.Snapshot.ScriptModules.HasActions ? Visibility.Visible : Visibility.Collapsed;
        RunActionButton.IsEnabled = !_busy && _document.ScriptsEnabled && ActionPicker.Items.Count > 0;
    }

    private void UpdateFormulaBar()
    {
        UpdateRibbonState();
        if (_formulaPending) return; // A selection/refresh must not overwrite an unsubmitted edit.
        CellAddress.Text = ColumnName(_selectedColumn) + (_selectedRow + 1);
        _updatingFormula = true;
        CellEditorText.InitializeText(FormulaBar, _document.GetCellInput(_sheetIndex, _selectedRow, _selectedColumn));
        _updatingFormula = false;
    }

    private void FormulaBar_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_updatingFormula || _refreshing || _document is null) return;
        _editState.BeginFormula(_sheetIndex, _selectedRow, _selectedColumn);
    }

    private bool CommitFormulaBar()
    {
        if (!CanCommitComposition()) return false;
        if (!_formulaPending) return true;
        try
        {
            var result = _editor.SetCell(_formulaSheet, _formulaRow, _formulaColumn, FormulaBar.Text);
            _calculationPending |= result.Changed && !result.CalculationUnchanged;
            _editState.EndFormula();
            UpdateChrome();
            return true;
        }
        catch (Exception ex) { ReportEditFailure(ex); return false; }
    }

    private Task<bool> WorkAsync(string message, Action operation) => WorkAsync(message, _ => operation(), cancellable: false);

    private async Task<bool> WorkAsync(string message, Action<CancellationToken> operation, bool cancellable)
    {
        if (_busy || _windowClosed) return false;
        var inputFocus = Keyboard.FocusedElement as UIElement;
        _busy = true;
        using var cancellation = cancellable ? new CancellationTokenSource() : null;
        _calculationCancellation = cancellation;
        StopCalculationButton.Visibility = cancellable ? Visibility.Visible : Visibility.Collapsed;
        // Keep enabled visuals and the attached native input context stable. Preview input
        // handlers enforce the busy gate without disabling/re-enabling the entire visual tree.
        StatusText.Text = message;
        var timer = Stopwatch.StartNew();
        try
        {
            await Task.Run(() => operation(cancellation?.Token ?? CancellationToken.None));
            StatusText.Text = $"{message.TrimEnd('…')}完成 · {timer.ElapsedMilliseconds:N0} ms";
            return true;
        }
        catch (Exception ex)
        {
            if (ex is OperationCanceledException) { StatusText.Text = "操作已停止，保留已提交的数据与结果。"; return false; }
            StatusText.Text = "操作未完成";
            AppendLog(ex.Message);
            if (!_smokeMode) MessageBox.Show(this, ex.Message, "QuickExcel", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        finally
        {
            _busy = false;
            _calculationCancellation = null;
            StopCalculationButton.Visibility = Visibility.Collapsed;
            UpdateChrome();
            // Restore focus if the user selected Stop while the operation was running.
            if (IsActive && inputFocus is { IsVisible: true, IsEnabled: true } &&
                PresentationSource.FromVisual(inputFocus) is not null)
                Keyboard.Focus(inputFocus);
        }
    }

    private void ShowDiagnostics()
    {
        ScriptState.Text = _document.ScriptsEnabled ? "● 脚本已启用" : "● 脚本未启用";
        if (_document.ScriptErrors.Count > 0) AppendLog(string.Join("\n", _document.ScriptErrors.Take(8)));
        RefreshDiagnostics();
        RefreshSheet();
    }

    private void AppendLog(string message)
    {
        var text = LogText.Text + $"\n[{DateTime.Now:HH:mm:ss}] {message}";
        LogText.Text = text.Length > 12000 ? text[^12000..] : text;
        LogText.ScrollToEnd();
    }

    private void ReportEditFailure(Exception exception)
    {
        StatusText.Text = exception.Message.Length <= 180 ? exception.Message : exception.Message[..180] + "…";
        AppendLog(exception.Message);
    }

    private async Task<bool> RecalculateAsync()
    {
        if (_busy) return false;
        _calculationPending = false;
        var committed = false;
        var ok = await WorkAsync("重新计算…", token => committed = _document.Calculate(token), cancellable: true);
        ok &= committed;
        if (!ok) _calculationPending = true;
        ShowDiagnostics();
        return ok;
    }

    private bool ApplyValues(IEnumerable<(int Row, int Column, string Value)> values)
    {
        var result = _editor.SetCells(_sheetIndex, values);
        _calculationPending |= result.Changed && !result.CalculationUnchanged;
        UpdateChrome();
        return result.Changed;
    }

    private void QueueRecalculation() => Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(async () =>
    {
        if (_calculationPending && !_busy && !_suppressAutoCalculation && _gridResize is null && !_editState.GridEditing && !_formulaPending && !_editState.IsComposing &&
            !_closeRequestPending && !_closingApproved && !_windowClosed) await RecalculateAsync();
    }));

    private async Task ApplyFormulaBarAsync()
    {
        // Finish the older grid editor before applying the newer formula-bar draft.
        if (_busy || !CommitGridEdit() || !CommitFormulaBar()) return;
        if (_calculationPending) await RecalculateAsync();
        else UpdateFormulaBar();
    }

    private async void ApplyCell_Click(object sender, RoutedEventArgs e) => await ApplyFormulaBarAsync();
    private bool CommitGridValue(CellPosition position, string text)
    {
        if (_busy || !CanCommitComposition()) return false;
        try
        {
            if (ApplyValues([(position.Row, position.Column, text)])) QueueRecalculation();
            return true;
        }
        catch (Exception ex) { _cellCommitFailed = true; ReportEditFailure(ex); return false; }
    }

    private void Grid_CurrentCellChanged()
    {
        if (_refreshing || _busy) return;
        _selectedRow = _grid.Current.Row;
        _selectedColumn = _grid.Current.Column;
        var merge = _document.Snapshot.Sheets[_sheetIndex].MergedCells.Select(CellRectangle.Parse)
            .Where(range => range.Contains(_selectedRow, _selectedColumn)).Cast<CellRectangle?>().FirstOrDefault();
        if (merge is { } merged && (merged.Top != _selectedRow || merged.Left != _selectedColumn) && _grid.Control.SelectedCellCount <= 1)
        {
            GridSurface.Select(new(merged.Top, merged.Left));
            return;
        }
        UpdateFormulaBar();
        QueueRecalculation();
    }

    private async void Undo_Click(object sender, RoutedEventArgs e) => await UndoAsync(false);
    private async void Redo_Click(object sender, RoutedEventArgs e) => await UndoAsync(true);
    private async Task UndoAsync(bool redo)
    {
        if (_busy || !CommitEditors()) return;
        var sheetId = _document.Snapshot.Sheets[_sheetIndex].Id;
        var result = redo ? _editor.Redo() : _editor.Undo();
        if (!result.Changed) return;
        if (result.CalculationUnchanged)
        {
            RefreshSheet();
            UpdateChrome();
            if (_calculationPending) await RecalculateAsync();
            return;
        }
        var sheets = _document.Snapshot.Sheets;
        var index = Array.FindIndex(sheets.ToArray(), sheet => sheet.Id == sheetId);
        _sheetIndex = index >= 0 ? index : Math.Min(_sheetIndex, sheets.Length - 1);
        if (CurrentView.Region is null)
        {
            CurrentView.FilterRange = null;
            CurrentView.FrozenRows = CurrentView.FrozenColumns = 0;
        }
        SynchronizeSheets();
        _calculationPending = true;
        await RecalculateAsync();
    }

    private void SynchronizeSheets()
    {
        _refreshing = true;
        SheetPicker.ItemsSource = _document.Snapshot.Sheets.Select(sheet => sheet.Name).ToList();
        SheetPicker.SelectedIndex = _sheetIndex;
        _refreshing = false;
        RefreshSheet(true);
        UpdateChrome();
    }

    private void SheetPicker_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_refreshing || SheetPicker.SelectedIndex < 0 || _busy) return;
        if (!CommitEdit())
        {
            _refreshing = true;
            SheetPicker.SelectedIndex = _sheetIndex;
            _refreshing = false;
            return;
        }
        _sheetIndex = SheetPicker.SelectedIndex;
        _selectedColumn = _selectedRow = 0;
        _columnCount = 12;
        _rowCount = 100;
        RefreshSheet(true);
        GridSurface.Select(new(0, 0));
        QueueRecalculation();
    }

    private void MoreRows_Click(object sender, RoutedEventArgs e) { if (!CommitEdit()) return; _rowCount = Math.Min(100000, _rowCount + 100); RefreshSheet(true); }
    private void MoreColumns_Click(object sender, RoutedEventArgs e) { if (!CommitEdit()) return; _columnCount = Math.Min(256, _columnCount + 8); RefreshSheet(true); }
    private void AddSheet_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEdit()) return;
        try
        {
            var number = _document.Snapshot.Sheets.Length + 1;
            while (_document.Snapshot.Sheets.Any(sheet => sheet.Name == $"工作表{number}")) number++;
            _editor.AddSheet($"工作表{number}");
            _sheetIndex = _document.Snapshot.Sheets.Length - 1;
            _selectedRow = _selectedColumn = 0;
            SynchronizeSheets();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private void ScriptEditor_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_refreshing || ScriptState is null || _editor is null) return;
        // Even rejected source edits must revoke the old source's authorization.
        _document.DisableScripts();
        var signedModules = _document.Snapshot.ScriptModules;
        if (signedModules.Signature is not null) _editor.Execute(new UpdateScriptSignatureCommand(signedModules, null));
        ActionPicker.ItemsSource = null;
        ScriptState.Text = "● 已修改，等待应用";
        try
        {
            CommitScriptModule();
            UpdateChrome();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private async void Run_Click(object sender, RoutedEventArgs e) => await RunScriptsAsync();
    private async Task RunScriptsAsync(bool automatic = false)
    {
        if (_busy) return;
        _suppressAutoCalculation = true;
        try { if (!CommitEdit()) return; }
        finally { _suppressAutoCalculation = false; }
        _calculationPending = false;
        if (!await CheckScriptTrustAsync(automatic)) return;
        IReadOnlyList<ScriptActionDescriptor> actions = [];
        ActionPicker.ItemsSource = null;
        var ok = await WorkAsync("应用 JavaScript 并计算…", token =>
        {
            _document.EnableScriptsAndCalculate(token);
            if (_document.Snapshot.ScriptModules.HasActions) actions = _document.ListActions(token);
        }, cancellable: true);
        if (ok) { ActionPicker.ItemsSource = actions; ActionPicker.SelectedIndex = actions.Count > 0 ? 0 : -1; }
        if (ok) AppendLog("脚本已应用，工作簿已重新计算。");
        ShowDiagnostics();
        UpdateChrome();
    }

    private void Disable_Click(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        _document.DisableScripts();
        ActionPicker.ItemsSource = null;
        ScriptState.Text = "● 脚本未启用";
        AppendLog("脚本已停用；自定义公式保留上次计算结果。");
        RefreshSheet();
        UpdateChrome();
    }

    private void StopCalculation_Click(object sender, RoutedEventArgs e)
    {
        _calculationCancellation?.Cancel();
        StatusText.Text = "正在停止…";
    }

    private Task _actionExecution = Task.CompletedTask;
    private async void RunAction_Click(object sender, RoutedEventArgs e) => await (_actionExecution = RunSelectedActionAsync());
    private async Task RunSelectedActionAsync()
    {
        if (_busy || !_document.ScriptsEnabled || ActionPicker.SelectedItem is not ScriptActionDescriptor action) return;
        _suppressAutoCalculation = true;
        try { if (!CommitEdit()) return; }
        finally { _suppressAutoCalculation = false; }
        CommandResult? result = null;
        var ok = await WorkAsync("运行操作脚本…", token => result = _document.RunAction(action.Name, token), cancellable: true);
        if (ok)
        {
            _calculationPending |= result!.Changed;
            AppendLog(result.Changed ? $"动作「{action.Name}」已提交，可一次撤销。公式缓存等待重算。" : $"动作「{action.Name}」未产生变更。");
        }
        ShowDiagnostics();
        UpdateChrome();
    }

    private bool CommitGridEdit()
    {
        _cellCommitFailed = false;
        return CanCommitComposition() && GridSurface.Commit() && !_cellCommitFailed;
    }

    private bool CommitEdit()
    {
        if (!CommitGridEdit() || !CommitFormulaBar()) return false;
        try { CommitScriptModule(); return true; }
        catch (Exception ex) { ReportEditFailure(ex); return false; }
    }

    private bool CommitEditors()
    {
        var previous = _suppressAutoCalculation;
        _suppressAutoCalculation = true;
        try { return CommitEdit(); }
        finally { _suppressAutoCalculation = previous; }
    }

    private async Task<bool> FlushEditsAsync()
    {
        if (!CommitEditors()) return false;
        if (_calculationPending && !await RecalculateAsync())
            AppendLog("重算未完成，将保存已提交的输入和上次缓存；部分结果需要重新计算。");
        return true;
    }

    private async Task<bool> ConfirmDiscardAsync()
    {
        if (!CommitEditors()) return false;
        if (!_document.IsDirty) return true;
        var result = _smokeMode && _smokeDiscardDecision is { } decision ? decision :
            MessageBox.Show(this, "保存当前工作簿和脚本后再继续？", "QuickExcel", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        return result == MessageBoxResult.No || result == MessageBoxResult.Yes && await SaveAsync(false);
    }

    private async void New_Click(object sender, RoutedEventArgs e)
    { if (!_busy && await ConfirmDiscardAsync()) InstallDocument(CreateProductWorkbook(), null, "未命名.xlsq"); }
    private async void Demo_Click(object sender, RoutedEventArgs e)
    { if (!_busy && await ConfirmDiscardAsync()) InstallDocument(WorkbookDocument.CreateDemo(_services), null, "示例工作簿.xlsq"); }
    private async void Open_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        var dialog = new OpenFileDialog { Filter = "工作簿 (*.xlsq;*.xlsx)|*.xlsq;*.xlsx|QuickExcel (*.xlsq)|*.xlsq|Excel (*.xlsx)|*.xlsx", Title = "打开工作簿" };
        if (dialog.ShowDialog(this) == true && await ConfirmDiscardAsync()) await OpenPathAsync(dialog.FileName);
    }

    private async Task OpenPathAsync(string path, bool allowAutomaticScripts = true)
    {
        WorkbookDocument? next = null;
        FileVersion? version = null;
        if (await WorkAsync("打开工作簿…", () =>
        {
            // Deny write/delete while both the fingerprint and model are read.
            using var lease = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            version = FileVersions.Read(lease);
            next = WorkbookDocument.Open(path, _services);
        }))
        {
            InstallDocument(next!, Path.GetFullPath(path), Path.GetFileName(path), fileVersion: version);
            RememberFile(Path.GetFullPath(path));
            if (allowAutomaticScripts && _document.Snapshot.ScriptModules.Signature is not null) await RunScriptsAsync(automatic: true);
        }
    }

    private async void Save_Click(object sender, RoutedEventArgs e) => await SaveAsync(false);
    private async void SaveAs_Click(object sender, RoutedEventArgs e) => await SaveAsync(true);
    private async Task<bool> SaveAsync(bool saveAs)
    {
        if (_busy || !await FlushEditsAsync()) return false;
        var snapshot = _document.Snapshot;
        var issues = TypedWorkbookStorage.SaveIssues(snapshot);
        var discardUnsupported = false;
        if (issues.Count > 0)
        {
            var choice = MessageBox.Show(this, "以下内容不能在编辑后可靠保留：\n" + string.Join("\n", issues.Take(15)) +
                "\n\n是否另存受支持副本？原文件将保留，副本会移除上述内容。选择“否”取消保存。", "另存兼容副本", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (choice != MessageBoxResult.Yes) return false;
            discardUnsupported = true;
            saveAs = true;
        }
        var migration = snapshot.FileState is { RequiresMigration: true } provenance &&
            (_filePath is null || string.Equals(Path.GetFullPath(_filePath), provenance.SourcePath, StringComparison.OrdinalIgnoreCase));
        string? path = saveAs || migration ? null : _filePath;
        if (path is null)
        {
            var dialog = new SaveFileDialog { Filter = "QuickExcel 工作簿 (*.xlsq)|*.xlsq", DefaultExt = ".xlsq", AddExtension = true,
                FileName = Path.GetFileNameWithoutExtension(_displayName) + (migration ? "-v2" : discardUnsupported ? "-受支持副本" : "") + ".xlsq", Title = "保存 XLSQ v2 工作簿和 JavaScript" };
            if (dialog.ShowDialog(this) != true) return false;
            path = dialog.FileName;
        }
        FileVersion? expected = null;
        if (!await WorkAsync("检查文件版本…", () => expected = _filePath is not null && Path.GetFullPath(path).Equals(Path.GetFullPath(_filePath), StringComparison.OrdinalIgnoreCase)
            ? (string.Equals(_versionPath, Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase) ? _fileVersion : null) ?? new FileVersion(null) : FileVersions.Read(path))) return false;
        if (!await SaveToPathAsync(path, discardUnsupported, expected!)) return false;
        if (_viewModel.StaleCacheCount > 0) StatusText.Text = "已保存输入和上次缓存；部分公式结果仍待重算。";
        return true;
    }

    private async void ExportPortable_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        PortableExportReport? report = null;
        if (!await WorkAsync("准备可移植公式…", () => report = _document.InspectPortableExport())) return;
        if (report!.MissingResults > 0)
        {
            MessageBox.Show(this, $"{report.MissingResults} 个脚本相关公式缺少结果，请先完成计算。", "不能导出", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        var losses = _document.Snapshot.FileState?.CompatibilityIssues.ToArray() ?? [];
        var message = $"保留 {report.RetainedFormulas} 个独立标准公式；将 {report.FrozenFormulas} 个脚本、循环或不支持公式及其传递依赖冻结为值。名称会展开为引用。";
        if (report.StaleResults > 0) message += $"\n其中 {report.StaleResults} 个结果可能过期，继续将使用已有缓存。";
        if (losses.Length > 0) message += "\n副本将移除以下未支持内容：\n" + string.Join("\n", losses.Take(15));
        if (MessageBox.Show(this, message + "\n\n是否导出？", "可移植公式", MessageBoxButton.YesNo,
            report.StaleResults > 0 || losses.Length > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information) != MessageBoxResult.Yes) return;
        var dialog = new SaveFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", DefaultExt = ".xlsx", AddExtension = true,
            FileName = Path.GetFileNameWithoutExtension(_displayName) + "-公式.xlsx", Title = "导出可移植公式" };
        if (dialog.ShowDialog(this) != true) return;
        if (await WorkAsync("导出可移植公式…", () => _document.ExportPortableXlsx(dialog.FileName, report.StaleResults > 0, losses.Length > 0)))
            AppendLog($"已导出 .xlsx：保留 {report.RetainedFormulas} 个公式，冻结 {report.FrozenFormulas} 个结果。");
    }

    private async void Export_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        var dialog = new SaveFileDialog { Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", DefaultExt = ".xlsx", AddExtension = true,
            FileName = Path.GetFileNameWithoutExtension(_displayName) + "-结果.xlsx", Title = "导出当前结果（所有公式转为值，不含脚本）" };
        if (dialog.ShowDialog(this) != true) return;
        var formulas = _document.Snapshot.Sheets.SelectMany(sheet => sheet.Cells.Values).Where(cell => cell.Content.Kind == CellContentKind.Formula).ToArray();
        var stale = formulas.Count(cell => cell.Calculation?.Value is not null &&
            (cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error) || cell.Calculation.InputRevision != _document.Snapshot.Revision));
        var losses = _document.Snapshot.FileState?.CompatibilityIssues.ToArray() ?? [];
        var hostErrors = formulas.Count(cell => cell.Value.Kind == CellValueKind.Error && cell.Value.ErrorCode is CellErrorCode.Script or CellErrorCode.Cycle);
        if (stale > 0 || losses.Length > 0 || hostErrors > 0)
        {
            var message = $"将导出已有结果：{stale} 个公式缓存可能过期，{hostErrors} 个宿主错误将映射为 #VALUE!。" +
                (losses.Length == 0 ? "" : "\n副本将移除以下未支持内容：\n" + string.Join("\n", losses.Take(15))) + "\n\n是否继续？";
            if (MessageBox.Show(this, message, "确认导出已有结果", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
        }
        if (await WorkAsync("导出当前结果…", () => _document.ExportXlsx(dialog.FileName, allowStaleCaches: stale > 0, discardUnsupported: losses.Length > 0)))
            AppendLog("已导出 .xlsx：公式转为类型化缓存值，原工作簿和脚本保持不变。");
    }

    private async void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (_busy)
        {
            if (e.Key == Key.Escape) _calculationCancellation?.Cancel();
            if (e.Key == Key.Tab && StopCalculationButton.IsVisible) StopCalculationButton.Focus();
            e.Handled = !StopCalculationButton.IsKeyboardFocused || e.Key is not (Key.Enter or Key.Space);
            return;
        }
        if (!_busy && !_editState.IsComposing && e.Key == Key.F6)
        {
            e.Handled = true;
            if (ScriptPanel.IsVisible) { if (ScriptEditor.IsKeyboardFocusWithin) ModulePicker.Focus(); else ScriptEditor.Focus(); }
            else if (WorkbookPage.IsVisible)
            {
                if (FormulaBar.IsKeyboardFocusWithin) SheetPicker.Focus();
                else if (SheetPicker.IsKeyboardFocusWithin) GridSurface.Select(new(_selectedRow, _selectedColumn));
                else if (CommitGridEdit()) FormulaBar.Focus();
            }
            return;
        }
        if (_busy || _editState.IsComposing || !Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) return;
        switch (e.Key)
        {
            case Key.PageDown: case Key.PageUp:
                e.Handled = true;
                if (CommitEditors()) { ShowWorkbookPage(); SheetPicker.SelectedIndex = Math.Clamp(_sheetIndex + (e.Key == Key.PageDown ? 1 : -1), 0, SheetPicker.Items.Count - 1); }
                break;
            case Key.F: case Key.H: e.Handled = true; Find_Click(sender, e); break;
            case Key.G: e.Handled = true; GoTo_Click(sender, e); break;
            case Key.Tab when ScriptEditor.IsKeyboardFocusWithin:
                e.Handled = true; RunButton.Focus(); break;
            case Key.S: e.Handled = true; await SaveAsync(Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)); break;
            case Key.O: e.Handled = true; Open_Click(sender, e); break;
            case Key.N: e.Handled = true; New_Click(sender, e); break;
            case Key.W: e.Handled = true; CloseDocument_Click(sender, e); break;
            case Key.Enter: e.Handled = true; await RunScriptsAsync(); break;
            case Key.Z when !ScriptEditor.IsKeyboardFocusWithin && !FormulaBar.IsKeyboardFocusWithin && (e.OriginalSource is not TextBox || _gridEditorArmed):
                e.Handled = true; await UndoAsync(false); break;
            case Key.Y when !ScriptEditor.IsKeyboardFocusWithin && !FormulaBar.IsKeyboardFocusWithin && (e.OriginalSource is not TextBox || _gridEditorArmed):
                e.Handled = true; await UndoAsync(true); break;
        }
    }

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        if (_closingApproved) return;
        e.Cancel = true;
        if (_closeRequestPending) return;
        if (_busy) { StatusText.Text = "请等待当前操作结束。"; return; }
        _closeRequestPending = true;
        // Even an async confirmation can complete synchronously. Leave Closing
        // before showing owned dialogs or requesting the approved second close.
        Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(async () => await CompleteCloseRequestAsync()));
    }

    private async Task CompleteCloseRequestAsync()
    {
        try
        {
            if (_busy || _windowClosed || !await ConfirmDiscardAsync()) return;
            _closingApproved = true;
            Close();
        }
        finally
        {
            _closeRequestPending = false;
            if (!_windowClosed)
            {
                _closingApproved = false;
                QueueRecalculation();
            }
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        _windowClosed = true;
        _document.Dispose();
        base.OnClosed(e);
    }

    public static string ColumnName(int index)
    {
        var result = "";
        do { result = (char)('A' + index % 26) + result; index = index / 26 - 1; } while (index >= 0);
        return result;
    }

    private async Task RunSmokeTestAsync(string folder)
    {
        Directory.CreateDirectory(folder);
        await RunCloseSmokeTestsAsync(folder);
        await RunWorkerSmokeTestsAsync(folder);
        // Drive the same editing, undo and formula-bar paths used by the window.
        InstallDocument(WorkbookDocument.CreateBlank(_services), null, "界面验收.xlsq");
        ScriptEditor.Text = "quickExcel.functions.register(\"DOUBLE\", x => x * 2);";
        ApplyValues([(0, 0, "21"), (0, 1, "=DOUBLE(A1)")]);
        if (!await WorkAsync("界面验收重算…", _document.EnableScriptsAndCalculate)) throw new Exception("Smoke calculation failed.");
        _calculationPending = false;
        if (_document.GetCellValue(0, 0, 1) != "42") throw new Exception("JS formula integration failed.");
        _selectedRow = 0; _selectedColumn = 0;
        FormulaBar.Text = "25";
        await ApplyFormulaBarAsync();
        if (_document.GetCellValue(0, 0, 1) != "50") throw new Exception("Formula-bar recalculation failed.");
        await UndoAsync(false);
        if (_document.GetCellValue(0, 0, 1) != "42") throw new Exception("Undo recalculation failed.");
        await UndoAsync(true);
        if (_document.GetCellValue(0, 0, 1) != "50") throw new Exception("Redo recalculation failed.");
        _document.DisableScripts();
        ApplyValues([(0, 1, "99")]);
        await UndoAsync(false);
        if (_document.GetCellValue(0, 0, 1) != "50") throw new Exception("Disabled formula cache was lost on undo.");
        var beforeRejected = _document.GetCellValue(0, 0, 0);
        try
        {
            ApplyValues([(0, 0, "999"), (1, 0, new string('x', 32768))]);
            throw new Exception("Oversized paste was accepted.");
        }
        catch (ArgumentException) { }
        if (_document.GetCellValue(0, 0, 0) != beforeRejected) throw new Exception("Rejected paste partially changed workbook.");
        if (!await WorkAsync("界面验收重算…", _document.EnableScriptsAndCalculate)) throw new Exception("Smoke calculation failed.");
        _calculationPending = false;
        RefreshSheet();
        var path = Path.Combine(folder, "smoke.xlsq");
        _filePath = path;
        GridSurface.Select(new(0, 0));
        SheetGrid.Focus();
        if (!GridSurface.BeginEdit()) throw new Exception("Grid could not enter edit mode.");
        var editBox = _grid.Editor;
        editBox.Text = "30";
        if (!await SaveAsync(false)) throw new Exception("Saving an active cell edit failed.");
        using (var reopened = WorkbookDocument.Open(path, _services))
        {
            if (reopened.ScriptsEnabled) throw new Exception("Opened scripts ran automatically.");
            if (reopened.GetCellValue(0, 0, 1) != "60") throw new Exception("Saved formula cache did not include the pending edit.");
            reopened.EnableScriptsAndCalculate();
            if (reopened.GetCellValue(0, 0, 1) != "60") throw new Exception("Package roundtrip failed.");
            reopened.ExportXlsx(Path.Combine(folder, "smoke.xlsx"));
        }
        // Saving without Enter must commit the formula bar to its original cell address.
        _selectedRow = 0; _selectedColumn = 0;
        UpdateFormulaBar();
        FormulaBar.Focus();
        FormulaBar.Text = "35";
        var beforeFormulaSaveRevision = _document.Snapshot.Revision;
        if (!await SaveAsync(false)) throw new Exception("Saving a pending formula bar failed.");
        if (_document.Snapshot.Revision != beforeFormulaSaveRevision + 1) throw new Exception("Formula save did not commit exactly one command revision.");
        using (var formulaSaved = WorkbookDocument.Open(path, _services))
            if (formulaSaved.GetCellValue(0, 0, 1) != "70") throw new Exception("Formula bar pending edit was omitted from saved cache.");
        if (_document.IsDirty) throw new Exception("Successful save did not mark the session revision saved.");
        // Entering the grid in the same cell must not let an older formula-bar draft overwrite it.
        GridSurface.Select(new(0, 0));
        FormulaBar.Text = "36";
        SheetGrid.Focus();
        if (!GridSurface.BeginEdit()) throw new Exception($"Same-cell editor switch failed (busy={_busy}, enabled={SheetGrid.IsEnabled}, current={_grid.Current}).");
        var switchedEditor = _grid.Editor;
        switchedEditor.Text = "37";
        if (!await SaveAsync(false)) throw new Exception("Saving the latest editor failed.");
        using (var latestEditor = WorkbookDocument.Open(path, _services))
            if (latestEditor.GetCellValue(0, 0, 0) != "37" || latestEditor.GetCellValue(0, 0, 1) != "74")
                throw new Exception("An older formula draft overwrote the latest grid edit.");
        FormulaBar.Text = new string('x', WorkbookLimits.MaxCellTextLength + 1);
        var invalidRevision = _document.Snapshot.Revision;
        GridSurface.Select(new(0, 1));
        // Rejected input must leave the visual selection at the original A1 address.
        if (_grid.Current != new CellPosition(0, 0) ||
            _selectedColumn != 0 || !_formulaPending || _formulaColumn != 0 || _formulaRow != 0)
            throw new Exception($"Rejected formula input left selection and edit address out of sync (grid={_grid.Current.Column}, selected={_selectedColumn}, formula={_formulaColumn}, pending={_formulaPending}, busy={_busy}).");
        if (await SaveAsync(false) || _document.Snapshot.Revision != invalidRevision)
            throw new Exception("Invalid pending input changed or saved the document.");
        using (var previousSave = WorkbookDocument.Open(path, _services))
            if (previousSave.GetCellValue(0, 0, 1) != "74") throw new Exception("Rejected edit changed the saved file.");
        FormulaBar.Text = "37";
        if (!CommitFormulaBar()) throw new Exception("Corrected formula input was rejected.");
        var authorizedRevision = _document.Snapshot.Revision;
        await RunScriptsAsync();
        if (_document.IsDirty || _document.Snapshot.Revision != authorizedRevision) throw new Exception("Script authorization changed document dirty state.");
        var acceptedSource = _document.ScriptSource;
        ScriptEditor.Text = new string('x', 1024 * 1024 + 1);
        if (_document.ScriptsEnabled || _document.ScriptSource != acceptedSource)
            throw new Exception("Rejected script edit retained authorization or changed the accepted source.");
        ApplyValues([(0, 0, "38")]);
        if (!await RecalculateAsync() || _document.GetCellValue(0, 0, 1) != "74" || _document.ScriptsEnabled)
            throw new Exception("Rejected script source allowed the previous authorized function to run.");
        ScriptEditor.Text = acceptedSource;
        await RunScriptsAsync();
        if (_document.GetCellValue(0, 0, 1) != "76") throw new Exception("Corrected script could not be explicitly reauthorized.");
        var beforeSheetCount = _document.Snapshot.Sheets.Length;
        AddSheet_Click(this, new RoutedEventArgs());
        if (_document.Snapshot.Sheets.Length != beforeSheetCount + 1) throw new Exception("Add sheet command failed.");
        await UndoAsync(false);
        if (_document.Snapshot.Sheets.Length != beforeSheetCount) throw new Exception("Add sheet undo failed.");
        await UndoAsync(true);
        if (_document.Snapshot.Sheets.Length != beforeSheetCount + 1) throw new Exception("Add sheet redo failed.");
        var reversePath = Path.Combine(folder, "reverse-editor.xlsq");
        InstallDocument(WorkbookDocument.CreateBlank(_services), reversePath, "reverse-editor.xlsq");
        GridSurface.Select(new(0, 0));
        SheetGrid.Focus();
        if (!GridSurface.BeginEdit())
            throw new Exception("Reverse editor switch could not start grid editing.");
        var olderGridEditor = _grid.Editor;
        olderGridEditor.Text = "11";
        FormulaBar.Focus();
        FormulaBar.Text = "=1+2";
        await ApplyFormulaBarAsync();
        if (!await SaveAsync(false)) throw new Exception("Reverse editor switch could not save.");
        using (var reverseSaved = WorkbookDocument.Open(reversePath, _services))
            if (reverseSaved.GetCellInput(0, 0, 0) != "=1+2" || reverseSaved.GetCellValue(0, 0, 0) != "3")
                throw new Exception("An older grid editor overwrote the newer formula-bar expression.");
        // A failed calculation must leave the user able to undo, save inputs, and close.
        var failingCalculator = new SmokeFailingCalculator();
        var failingServices = new WorkbookServices(_services.Storage, failingCalculator);
        var recoveryPath = Path.Combine(folder, "recovery.xlsq");
        InstallDocument(WorkbookDocument.CreateBlank(failingServices), recoveryPath, "recovery.xlsq");
        ApplyValues([(0, 0, "1"), (0, 1, "=A1*2")]);
        ApplyValues([(0, 0, "2")]);
        if (await RecalculateAsync()) throw new Exception("The failing calculator unexpectedly succeeded.");
        await UndoAsync(false);
        if (_document.GetCellValue(0, 0, 0) != "1") throw new Exception("Failed recalculation prevented undo.");
        if (!await SaveAsync(false) || _document.IsDirty) throw new Exception("Failed recalculation prevented saving inputs.");
        using (var recoverySaved = WorkbookDocument.Open(recoveryPath, _services))
            if (recoverySaved.GetCellValue(0, 0, 0) != "1" || recoverySaved.GetCellInput(0, 0, 1) != "=A1*2")
                throw new Exception("Recovery save did not preserve inputs and formulas.");
        var beforeCloseAttempts = failingCalculator.Attempts;
        if (!await ConfirmDiscardAsync() || failingCalculator.Attempts != beforeCloseAttempts)
            throw new Exception("Closing a saved document retried the failed calculation.");
        await RunStorageSmokeAsync(folder);
        await RunCalculationSmokeAsync(folder);
        await RunDateDisplaySmokeAsync(folder);
        await RunInteractionSmokeAsync(folder);
        await RunWorkspaceSmokeAsync(folder);
        await RunProductSmokeAsync(folder);
        InstallDocument(WorkbookDocument.CreateDemo(_services), null, "示例工作簿.xlsq");
        if (!await WorkAsync("界面验收重算…", _document.EnableScriptsAndCalculate)) throw new Exception("Smoke calculation failed.");
        _calculationPending = false;
        ShowDiagnostics();
        AppendLog("界面验收通过：JS 计算、公式栏编辑、撤销/重做、.xlsq 往返、.xlsx 导出。");
        _document.Save(Path.Combine(folder, "QuickExcel-示例.xlsq"));
        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
        UpdateLayout();
        var dpi = VisualTreeHelper.GetDpi(this);
        var bitmap = new RenderTargetBitmap((int)(ActualWidth * dpi.DpiScaleX), (int)(ActualHeight * dpi.DpiScaleY), 96 * dpi.DpiScaleX, 96 * dpi.DpiScaleY, PixelFormats.Pbgra32);
        bitmap.Render(this);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using (var file = File.Create(Path.Combine(folder, "QuickExcel-preview.png"))) encoder.Save(file);
        File.WriteAllText(Path.Combine(folder, "smoke-result.txt"), "PASS: actual WPF close lifecycle (8 scenarios, including cancel/save/failure/duplicate requests), native QuickJS, formula-bar edit/recalculate, undo/redo with cached formulas, atomic rejected paste, save active grid and formula-bar edits with fresh caches, bidirectional same-cell editor switch, rejected pending input preserves selection and saved file, session revisions and clean script authorization, rejected script edit revokes authorization, failed calculation allows undo/save/close, add-sheet undo/redo, XLSQ roundtrip, values-only XLSX export, WPF render.");
    }

    private sealed class SmokeFailingCalculator : IWorkbookCalculator
    {
        public int Attempts { get; private set; }
        public CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default)
        {
            Attempts++;
            throw new TimeoutException("Smoke: simulated calculation deadline.");
        }
    }

}
