using System.Collections.ObjectModel;
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
    private readonly ObservableCollection<SheetRow> _rows = [];
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
    private bool _formulaPending;
    private int _formulaSheet;
    private int _formulaRow;
    private int _formulaColumn;
    private string? _filePath;
    private string _displayName = "示例工作簿.xlsq";

    public MainWindow()
    {
        InitializeComponent();
        SheetGrid.ItemsSource = _rows;
        InstallDocument(WorkbookDocument.CreateDemo(_services), null, "示例工作簿.xlsq", disposePrevious: false);
    }

    public async Task StartAsync(string[] args)
    {
        if (args.Length > 0 && args[0] == "--smoke-test")
        {
            _smokeMode = true;
            try
            {
                await RunSmokeTestAsync(args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "smoke"));
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
        if (args.Length > 0 && File.Exists(args[0])) await OpenPathAsync(args[0]);
    }

    private void InstallDocument(WorkbookDocument document, string? path, string displayName, bool disposePrevious = true)
    {
        if (disposePrevious) _document.Dispose();
        _document = document;
        _viewModel = new DocumentViewModel(document);
        _editor = new EditorController(document);
        DataContext = _viewModel;
        _filePath = path != null && Path.GetExtension(path).Equals(".xlsq", StringComparison.OrdinalIgnoreCase) ? path : null;
        _displayName = displayName;
        _sheetIndex = _selectedRow = _selectedColumn = 0;
        _rowCount = 100;
        _columnCount = 12;
        _formulaPending = _calculationPending = false;
        _refreshing = true;
        ScriptEditor.Text = document.ScriptSource;
        ActionPicker.ItemsSource = null;
        SheetPicker.ItemsSource = document.Snapshot.Sheets.Select(sheet => sheet.Name).ToList();
        SheetPicker.SelectedIndex = 0;
        _refreshing = false;
        RefreshSheet(true);
        UpdateChrome();
        ScriptState.Text = "● 脚本未启用";
        LogText.Text = "工作簿已载入。脚本尚未启用。XLSQ v2 保存类型、日期系统和稳定表身份。";
        if (document.Snapshot.FileState?.RequiresMigration == true)
            LogText.Text += "\n已导入 v1：历史文本不会被猜测转换类型。首次保存必须另存为 v2，原文件保留。";
        if (document.Snapshot.FileState is { CompatibilityIssues.Length: > 0 } fileState)
            LogText.Text += "\n兼容性报告：\n" + string.Join("\n", fileState.CompatibilityIssues.Take(20));
        if (!document.Snapshot.ScriptModules.CanExecute)
            LogText.Text += "\n多模块源码已保留；执行需要显式声明脚本 API v2。";
        else if (document.Snapshot.ScriptModules.ScriptApiVersion == 2)
            LogText.Text += $"\n包内模块：{document.Snapshot.ScriptModules.Modules.Count} 个；编辑器显示首入口 {document.Snapshot.ScriptModules.EntryModules[0]}。应用时执行全部声明入口。";
        HintText.Text = "0.5 脚本进程预览版 · 计算中可点击「停止计算」；可信脚本需手动启用。";
        if (document.Snapshot.Sheets.Any(sheet => sheet.UsedColumnCount > 256))
            HintText.Text += " 此工作簿超过界面的 256 列显示范围；未显示的受支持数据仍会保存。";
    }

    private void RefreshSheet(bool rebuild = false)
    {
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        if (rebuild)
        {
            _rowCount = Math.Clamp(Math.Max(_rowCount, sheet.UsedRowCount + 20), 100, WorkbookLimits.MaxRows);
            _columnCount = Math.Clamp(Math.Max(_columnCount, sheet.UsedColumnCount + 2), 12, 256);
            _refreshing = true;
            SheetGrid.Columns.Clear();
            SheetGrid.GridLinesVisibility = sheet.ShowGridLines ? DataGridGridLinesVisibility.All : DataGridGridLinesVisibility.None;
            for (var c = 0; c < _columnCount; c++)
            {
                var textStyle = new Style(typeof(TextBlock));
                textStyle.Setters.Add(new Setter(TextBlock.PaddingProperty, new Thickness(8, 4, 5, 3)));
                textStyle.Setters.Add(new Setter(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis));
                textStyle.Setters.Add(new Setter(TextBlock.ForegroundProperty, new Binding { Converter = new AppearanceConverter(c, "foreground") }));
                textStyle.Setters.Add(new Setter(TextBlock.BackgroundProperty, new Binding { Converter = new AppearanceConverter(c, "background") }));
                textStyle.Setters.Add(new Setter(TextBlock.FontWeightProperty, new Binding { Converter = new AppearanceConverter(c, "weight") }));
                var width = c == 0 ? 160d : 112d;
                sheet.Columns.TryGetValue(c, out var column);
                if (column?.Width is { } storedWidth) width = Math.Clamp(storedWidth * 7 + 5, 55, 400);
                SheetGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = ColumnName(c), Width = width,
                    Visibility = column?.Hidden == true ? Visibility.Collapsed : Visibility.Visible,
                    Binding = new Binding($"[{c}]") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.Explicit },
                    ElementStyle = textStyle,
                    EditingElementStyle = new Style(typeof(TextBox))
                    { Setters = { new Setter(TextBox.PaddingProperty, new Thickness(5, 3, 4, 2)) } }
                });
            }
            _rows.Clear();
            for (var r = 0; r < _rowCount; r++) _rows.Add(new SheetRow(this, r));
            _refreshing = false;
        }
        else foreach (var row in _rows) row.Refresh();
        UpdateFormulaBar();
        _viewModel.Refresh();
        var stale = _viewModel.StaleCacheCount;
        SelectionText.Text = $"{sheet.Cells.Count:N0} 个单元格 · {(stale > 0 ? $"{stale} 个缓存待重算" : "结果已更新")} · 本地运行";
    }

    private void UpdateChrome()
    {
        _viewModel.Refresh();
        FileTitle.Text = _displayName + (_viewModel.IsDirty ? "  •" : "");
        Title = $"{(_viewModel.IsDirty ? "* " : "")}{_displayName} — QuickExcel";
        FileSubtitle.Text = _filePath ?? "本地工作簿 · 保存为 .xlsq 保留脚本";
        if (_viewModel.DiscardedHistoryCount > 0) FileSubtitle.Text += $" · 已释放 {_viewModel.DiscardedHistoryCount} 条较早撤销记录";
        UndoButton.IsEnabled = _viewModel.CanUndo && !_busy;
        RedoButton.IsEnabled = _viewModel.CanRedo && !_busy;
        ActionsPanel.Visibility = _document.Snapshot.ScriptModules.HasActions ? Visibility.Visible : Visibility.Collapsed;
        RunActionButton.IsEnabled = !_busy && _document.ScriptsEnabled && ActionPicker.Items.Count > 0;
    }

    private void UpdateFormulaBar()
    {
        if (_formulaPending) return; // A selection/refresh must not overwrite an unsubmitted edit.
        CellAddress.Text = ColumnName(_selectedColumn) + (_selectedRow + 1);
        _updatingFormula = true;
        FormulaBar.Text = _document.GetCellInput(_sheetIndex, _selectedRow, _selectedColumn);
        _updatingFormula = false;
    }

    private void FormulaBar_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_updatingFormula || _refreshing || _document is null) return;
        if (!_formulaPending)
        {
            _formulaSheet = _sheetIndex;
            _formulaRow = _selectedRow;
            _formulaColumn = _selectedColumn;
        }
        _formulaPending = true;
    }

    private bool CommitFormulaBar()
    {
        if (!_formulaPending) return true;
        try
        {
            var result = _editor.SetCell(_formulaSheet, _formulaRow, _formulaColumn, FormulaBar.Text);
            _calculationPending |= result.Changed;
            _formulaPending = false;
            UpdateChrome();
            return true;
        }
        catch (Exception ex) { ReportEditFailure(ex); return false; }
    }

    private Task<bool> WorkAsync(string message, Action operation) => WorkAsync(message, _ => operation(), cancellable: false);

    private async Task<bool> WorkAsync(string message, Action<CancellationToken> operation, bool cancellable)
    {
        if (_busy || _windowClosed) return false;
        _busy = true;
        using var cancellation = cancellable ? new CancellationTokenSource() : null;
        _calculationCancellation = cancellation;
        StopCalculationButton.Visibility = cancellable ? Visibility.Visible : Visibility.Collapsed;
        Toolbar.IsEnabled = Workspace.IsEnabled = false;
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
            Toolbar.IsEnabled = Workspace.IsEnabled = true;
            UpdateChrome();
        }
    }

    private void ShowDiagnostics()
    {
        ScriptState.Text = _document.ScriptsEnabled ? "● 脚本已启用" : "● 脚本未启用";
        if (_document.ScriptErrors.Count > 0) AppendLog(string.Join("\n", _document.ScriptErrors.Take(8)));
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
        StatusText.Text = "编辑尚未提交，请修正输入后重试。";
        AppendLog(exception.Message);
    }

    private async Task<bool> RecalculateAsync()
    {
        if (_busy) return false;
        _calculationPending = false;
        var ok = await WorkAsync("重新计算…", token => _document.Calculate(token), cancellable: true);
        if (!ok) _calculationPending = true;
        ShowDiagnostics();
        return ok;
    }

    private bool ApplyValues(IEnumerable<(int Row, int Column, string Value)> values)
    {
        var result = _editor.SetCells(_sheetIndex, values);
        _calculationPending |= result.Changed;
        UpdateChrome();
        return result.Changed;
    }

    private void QueueRecalculation() => Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(async () =>
    {
        if (_calculationPending && !_busy && !_suppressAutoCalculation &&
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
    private async void FormulaBar_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) { e.Handled = true; await ApplyFormulaBarAsync(); if (!_formulaPending) SheetGrid.Focus(); }
        else if (e.Key == Key.Escape) { e.Handled = true; _formulaPending = false; UpdateFormulaBar(); SheetGrid.Focus(); }
    }

    private void Grid_LoadingRow(object sender, DataGridRowEventArgs e)
    {
        if (e.Row.Item is not SheetRow row) return;
        e.Row.Header = (row.Index + 1).ToString(CultureInfo.InvariantCulture);
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        sheet.Rows.TryGetValue(row.Index, out var metadata);
        e.Row.Height = metadata?.Hidden == true ? 0 : Math.Clamp(metadata?.Height ?? sheet.DefaultRowHeight ?? 31, 0, 400);
    }

    private void Grid_CurrentCellChanged(object? sender, EventArgs e)
    {
        if (_refreshing || _busy || SheetGrid.CurrentCell.Item is not SheetRow row || SheetGrid.CurrentCell.Column is null) return;
        if (!CommitFormulaBar())
        {
            // Keep the actual grid selection aligned with the pending edit's address.
            _refreshing = true;
            try
            {
                var previous = new DataGridCellInfo(_rows[_selectedRow], SheetGrid.Columns[_selectedColumn]);
                SheetGrid.CurrentCell = previous;
                SheetGrid.SelectedCells.Clear();
                SheetGrid.SelectedCells.Add(previous);
            }
            finally { _refreshing = false; }
            FormulaBar.Focus();
            return;
        }
        _selectedRow = row.Index;
        _selectedColumn = SheetGrid.CurrentCell.Column.DisplayIndex;
        UpdateFormulaBar();
        QueueRecalculation();
    }

    private void Grid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
    {
        // Switching editors in the same cell does not raise CurrentCellChanged.
        // Commit the older formula-bar input before loading the newer grid editor.
        if (_busy || !CommitFormulaBar()) { e.Cancel = true; return; }
        if (e.Row.Item is SheetRow row)
        {
            _selectedRow = row.Index;
            _selectedColumn = e.Column.DisplayIndex;
        }
    }

    private void Grid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
    {
        if (e.Row.Item is SheetRow row && e.EditingElement is TextBox editor)
        {
            editor.Text = _document.GetCellInput(_sheetIndex, row.Index, e.Column.DisplayIndex);
            editor.SelectAll();
        }
    }

    private void Grid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (_busy || e.EditAction != DataGridEditAction.Commit || e.Row.Item is not SheetRow row || e.EditingElement is not TextBox editor) return;
        try
        {
            if (ApplyValues([(row.Index, e.Column.DisplayIndex, editor.Text)])) QueueRecalculation();
        }
        catch (Exception ex) { _cellCommitFailed = true; e.Cancel = true; ReportEditFailure(ex); }
    }

    private async void Grid_KeyDown(object sender, KeyEventArgs e)
    {
        if (_busy || e.OriginalSource is TextBox) return;
        try
        {
            var ctrl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
            if (ctrl && e.Key == Key.V)
            {
                e.Handled = true;
                if (!Clipboard.ContainsText() || !CommitEdit()) return;
                var input = Clipboard.GetText().Replace("\r\n", "\n").TrimEnd('\n');
                if (input.Length > 4 * 1024 * 1024) throw new ArgumentException("单次粘贴文本不能超过 4 MiB。");
                var values = new List<(int Row, int Column, string Value)>();
                var lines = input.Split('\n');
                for (var r = 0; r < lines.Length; r++)
                {
                    var fields = lines[r].Split('\t');
                    if (values.Count + fields.Length > 20_000) throw new ArgumentException("单次粘贴上限为 20,000 个单元格。");
                    for (var c = 0; c < fields.Length; c++) values.Add((_selectedRow + r, _selectedColumn + c, fields[c]));
                }
                ApplyValues(values);
                RefreshSheet(true);
                await RecalculateAsync();
            }
            else if (ctrl && e.Key == Key.C)
            {
                e.Handled = true;
                var cells = SheetGrid.SelectedCells.Where(cell => cell.Item is SheetRow).ToList();
                if (cells.Count == 0) return;
                var minRow = cells.Min(cell => ((SheetRow)cell.Item).Index);
                var maxRow = cells.Max(cell => ((SheetRow)cell.Item).Index);
                var minCol = cells.Min(cell => cell.Column.DisplayIndex);
                var maxCol = cells.Max(cell => cell.Column.DisplayIndex);
                if ((long)(maxRow - minRow + 1) * (maxCol - minCol + 1) > 100000)
                { StatusText.Text = "复制范围过大，请缩小选择范围。"; return; }
                var text = new StringBuilder();
                for (var row = minRow; row <= maxRow; row++)
                {
                    if (row > minRow) text.AppendLine();
                    for (var column = minCol; column <= maxCol; column++)
                    {
                        if (column > minCol) text.Append('\t');
                        text.Append(_document.GetCellValue(_sheetIndex, row, column));
                    }
                }
                Clipboard.SetText(text.ToString());
                StatusText.Text = "已复制计算结果。";
            }
            else if (e.Key == Key.Delete)
            {
                e.Handled = true;
                if (!CommitEdit()) return;
                var result = _editor.ClearCells(_sheetIndex, SheetGrid.SelectedCells.Where(cell => cell.Item is SheetRow)
                    .Select(cell => (((SheetRow)cell.Item).Index, cell.Column.DisplayIndex)));
                _calculationPending |= result.Changed;
                UpdateChrome();
                if (_calculationPending) await RecalculateAsync();
            }
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private async void Undo_Click(object sender, RoutedEventArgs e) => await UndoAsync(false);
    private async void Redo_Click(object sender, RoutedEventArgs e) => await UndoAsync(true);
    private async Task UndoAsync(bool redo)
    {
        if (_busy || !CommitEditors()) return;
        var sheetId = _document.Snapshot.Sheets[_sheetIndex].Id;
        if (!(redo ? _editor.Redo() : _editor.Undo())) return;
        var sheets = _document.Snapshot.Sheets;
        var index = Array.FindIndex(sheets.ToArray(), sheet => sheet.Id == sheetId);
        _sheetIndex = index >= 0 ? index : Math.Min(_sheetIndex, sheets.Length - 1);
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
        ActionPicker.ItemsSource = null;
        ScriptState.Text = "● 已修改，等待应用";
        try
        {
            _editor.UpdateScriptSource(ScriptEditor.Text);
            UpdateChrome();
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private async void Run_Click(object sender, RoutedEventArgs e) => await RunScriptsAsync();
    private async Task RunScriptsAsync()
    {
        if (_busy) return;
        _suppressAutoCalculation = true;
        try { if (!CommitEdit()) return; }
        finally { _suppressAutoCalculation = false; }
        _calculationPending = false;
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

    private async void RunAction_Click(object sender, RoutedEventArgs e) => await RunSelectedActionAsync();
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
        return SheetGrid.CommitEdit(DataGridEditingUnit.Cell, true) && !_cellCommitFailed &&
            SheetGrid.CommitEdit(DataGridEditingUnit.Row, true) && !_cellCommitFailed;
    }

    private bool CommitEdit()
    {
        if (!CommitGridEdit() || !CommitFormulaBar()) return false;
        try { _editor.UpdateScriptSource(ScriptEditor.Text); return true; }
        catch (Exception ex) { ReportEditFailure(ex); return false; }
    }

    private bool CommitEditors()
    {
        _suppressAutoCalculation = true;
        try { return CommitEdit(); }
        finally { _suppressAutoCalculation = false; }
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
    { if (!_busy && await ConfirmDiscardAsync()) InstallDocument(WorkbookDocument.CreateBlank(_services), null, "未命名.xlsq"); }
    private async void Demo_Click(object sender, RoutedEventArgs e)
    { if (!_busy && await ConfirmDiscardAsync()) InstallDocument(WorkbookDocument.CreateDemo(_services), null, "示例工作簿.xlsq"); }
    private async void Open_Click(object sender, RoutedEventArgs e)
    {
        if (_busy) return;
        var dialog = new OpenFileDialog { Filter = "工作簿 (*.xlsq;*.xlsx)|*.xlsq;*.xlsx|QuickExcel (*.xlsq)|*.xlsq|Excel (*.xlsx)|*.xlsx", Title = "打开工作簿" };
        if (dialog.ShowDialog(this) == true && await ConfirmDiscardAsync()) await OpenPathAsync(dialog.FileName);
    }

    private async Task OpenPathAsync(string path)
    {
        WorkbookDocument? next = null;
        if (await WorkAsync("打开工作簿…", () => next = WorkbookDocument.Open(path, _services)))
            InstallDocument(next!, path, Path.GetFileName(path));
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
        if (!await WorkAsync("保存 .xlsq…", () => _document.Save(path, discardUnsupported))) return false;
        _filePath = path;
        _displayName = Path.GetFileName(path);
        UpdateChrome();
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
        if (_busy || !Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) return;
        switch (e.Key)
        {
            case Key.S: e.Handled = true; await SaveAsync(Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)); break;
            case Key.O: e.Handled = true; Open_Click(sender, e); break;
            case Key.N: e.Handled = true; New_Click(sender, e); break;
            case Key.Enter: e.Handled = true; await RunScriptsAsync(); break;
            case Key.Z when !ScriptEditor.IsKeyboardFocusWithin && !FormulaBar.IsKeyboardFocusWithin && e.OriginalSource is not TextBox:
                e.Handled = true; await UndoAsync(false); break;
            case Key.Y when !ScriptEditor.IsKeyboardFocusWithin && !FormulaBar.IsKeyboardFocusWithin && e.OriginalSource is not TextBox:
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
        SheetGrid.CurrentCell = new DataGridCellInfo(_rows[0], SheetGrid.Columns[0]);
        SheetGrid.ScrollIntoView(_rows[0], SheetGrid.Columns[0]);
        SheetGrid.Focus();
        if (!SheetGrid.BeginEdit()) throw new Exception("DataGrid could not enter edit mode.");
        if (SheetGrid.Columns[0].GetCellContent(_rows[0]) is not TextBox editBox) throw new Exception("DataGrid editor missing.");
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
        SheetGrid.CurrentCell = new DataGridCellInfo(_rows[0], SheetGrid.Columns[0]);
        SheetGrid.SelectedCells.Clear();
        SheetGrid.SelectedCells.Add(SheetGrid.CurrentCell);
        SheetGrid.ScrollIntoView(_rows[0], SheetGrid.Columns[0]);
        FormulaBar.Text = "36";
        SheetGrid.Focus();
        if (!SheetGrid.BeginEdit()) throw new Exception($"Same-cell editor switch failed (busy={_busy}, enabled={SheetGrid.IsEnabled}, valid={SheetGrid.CurrentCell.IsValid}).");
        if (SheetGrid.Columns[0].GetCellContent(_rows[0]) is not TextBox switchedEditor) throw new Exception("Switched grid editor missing.");
        switchedEditor.Text = "37";
        if (!await SaveAsync(false)) throw new Exception("Saving the latest editor failed.");
        using (var latestEditor = WorkbookDocument.Open(path, _services))
            if (latestEditor.GetCellValue(0, 0, 0) != "37" || latestEditor.GetCellValue(0, 0, 1) != "74")
                throw new Exception("An older formula draft overwrote the latest grid edit.");
        FormulaBar.Text = new string('x', WorkbookLimits.MaxCellTextLength + 1);
        var invalidRevision = _document.Snapshot.Revision;
        SheetGrid.CurrentCell = new DataGridCellInfo(_rows[0], SheetGrid.Columns[1]);
        // Moving focus back to the formula editor can clear WPF's CurrentCell.
        // It must never leave another valid grid address beside the pending A1 edit.
        if ((SheetGrid.CurrentCell.IsValid && SheetGrid.CurrentCell.Column != SheetGrid.Columns[0]) ||
            _selectedColumn != 0 || !_formulaPending || _formulaColumn != 0 || _formulaRow != 0)
            throw new Exception($"Rejected formula input left selection and edit address out of sync (grid={SheetGrid.CurrentCell.Column?.DisplayIndex}, selected={_selectedColumn}, formula={_formulaColumn}, pending={_formulaPending}, busy={_busy}).");
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
        SheetGrid.CurrentCell = new DataGridCellInfo(_rows[0], SheetGrid.Columns[0]);
        SheetGrid.SelectedCells.Clear();
        SheetGrid.SelectedCells.Add(SheetGrid.CurrentCell);
        SheetGrid.ScrollIntoView(_rows[0], SheetGrid.Columns[0]);
        SheetGrid.Focus();
        if (!SheetGrid.BeginEdit() || SheetGrid.Columns[0].GetCellContent(_rows[0]) is not TextBox olderGridEditor)
            throw new Exception("Reverse editor switch could not start grid editing.");
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
        File.WriteAllText(Path.Combine(folder, "smoke-result.txt"), "PASS: actual WPF close lifecycle (8 scenarios, including cancel/save/failure/duplicate requests), native QuickJS, formula-bar edit/recalculate, undo/redo with cached formulas, atomic rejected paste, save active DataGrid and formula-bar edits with fresh caches, bidirectional same-cell editor switch, rejected pending input preserves selection and saved file, session revisions and clean script authorization, rejected script edit revokes authorization, failed calculation allows undo/save/close, add-sheet undo/redo, XLSQ roundtrip, values-only XLSX export, WPF render.");
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

    private sealed class SheetRow(MainWindow owner, int index) : INotifyPropertyChanged
    {
        public int Index { get; } = index;
        public string this[int column] { get => owner._document.GetCellValue(owner._sheetIndex, Index, column); set { } }
        public StyleDefinition? GetStyle(int column)
        {
            var snapshot = owner._document.Snapshot;
            var sheet = snapshot.Sheets[owner._sheetIndex];
            var id = sheet.GetStyleId(Index, column);
            return snapshot.Styles[id];
        }
        public event PropertyChangedEventHandler? PropertyChanged;
        public void Refresh() => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Item[]"));
    }

    private sealed class AppearanceConverter(int column, string property) : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var style = (value as SheetRow)?.GetStyle(column);
            return property switch
            {
                "weight" => style?.Bold == true ? FontWeights.SemiBold : FontWeights.Normal,
                "background" => ColorBrush(style?.FillColor, Brushes.Transparent),
                _ => ColorBrush(style?.FontColor, new SolidColorBrush(Color.FromRgb(34, 53, 46)))
            };
        }
        private static Brush ColorBrush(string? color, Brush fallback)
        {
            try { return string.IsNullOrEmpty(color) ? fallback : (Brush)new BrushConverter().ConvertFromString("#" + color.TrimStart('#'))!; }
            catch { return fallback; }
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
    }
}
