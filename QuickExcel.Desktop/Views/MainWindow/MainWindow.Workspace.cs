using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Documents;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private readonly Dictionary<SheetId, SheetViewState> _sheetViews = [];
    private RecentFileStore _recentFiles = new(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QuickExcel", "recent-files.json"));
    private FileVersion? _fileVersion;
    private string? _versionPath;
    private string? _smokeTextDecision;
    private Task _workspaceInput = Task.CompletedTask;
    private static T? VisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T match) return match;
            if (VisualChild<T>(child) is { } found) return found;
        }
        return null;
    }
    private void InstallWorkspaceHandlers()
    {
        InstallGridResizeHandlers();
        _grid.HeaderRequested += OpenHeaderFilter;
        _grid.Control.ViewportChanged += (_, _) => _filterPopup?.Close();
        Deactivated += (_, _) => _filterPopup?.Close();
    }
    private void RefreshWorkspaceView()
    {
        var view = CurrentView;
        view.RefreshFilter(_document.Snapshot);
        ViewStatus.Text = (view.FilterRange is null ? "" : view.Region is null ? $"筛选 {ColumnName(view.FilterColumn)} · {view.FilterText}  " :
            view.FilterError is not null ? $"筛选待更新：{view.FilterError}（暂显示全部行）  " : $"筛选区域 · {view.Filters.Count} 列条件  ") +
            (view.FrozenRows + view.FrozenColumns == 0 ? "" : $"冻结 {view.FrozenRows} 行 / {view.FrozenColumns} 列");
        ViewStatus.ToolTip = "筛选首行为表头；视图设置仅本次工作区有效。";
        ViewStatus.Visibility = view.FilterRange is not null || view.FrozenRows + view.FrozenColumns > 0 ? Visibility.Visible : Visibility.Collapsed;
    }
    private SheetViewState CurrentView
    {
        get
        {
            var id = _document.Snapshot.Sheets[_sheetIndex].Id;
            if (!_sheetViews.TryGetValue(id, out var view)) _sheetViews[id] = view = new();
            return view;
        }
    }
    private GridRange SelectedRange()
    {
        var ranges = GridSurface.SelectionRanges;
        if (ranges.Count == 0) return new(_selectedRow, _selectedColumn, _selectedRow, _selectedColumn);
        if (ranges.Count == 1) return ranges[0];
        var rectangle = new GridRange(ranges.Min(r => r.Top), ranges.Min(r => r.Left), ranges.Max(r => r.Bottom), ranges.Max(r => r.Right));
        var sheet = _document.Snapshot.Sheets[_sheetIndex];
        var visibleRows = Enumerable.Range(rectangle.Top, rectangle.Rows).Count(r => sheet.Rows.GetValueOrDefault(r)?.Hidden != true);
        var visibleColumns = Enumerable.Range(rectangle.Left, rectangle.Columns).Count(c => sheet.Columns.GetValueOrDefault(c)?.Hidden != true);
        if (ranges.Sum(r => r.Count) == (long)visibleRows * visibleColumns) return rectangle;
        throw new ArgumentException("请使用连续矩形选区。");
    }

    private string? PromptText(string title, string label, string initial)
    {
        if (_smokeMode) return _smokeTextDecision;
        var input = new TextBox { Text = initial, Margin = new(0, 8, 0, 12), MinWidth = 320 };
        var ok = new Button { Content = "确定", IsDefault = true, MinWidth = 80 };
        var cancel = new Button { Content = "取消", IsCancel = true, MinWidth = 80 };
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(ok); buttons.Children.Add(cancel);
        var panel = new StackPanel { Margin = new(20) };
        panel.Children.Add(new TextBlock { Text = label, TextWrapping = TextWrapping.Wrap, MaxWidth = 460 });
        panel.Children.Add(input); panel.Children.Add(buttons);
        var dialog = new Window { Owner = this, Title = title, Content = panel, SizeToContent = SizeToContent.WidthAndHeight,
            ResizeMode = ResizeMode.NoResize, WindowStartupLocation = WindowStartupLocation.CenterOwner, ShowInTaskbar = false };
        ok.Click += (_, _) => dialog.DialogResult = true;
        dialog.Loaded += (_, _) => { input.Focus(); input.SelectAll(); };
        return dialog.ShowDialog() == true ? input.Text : null;
    }

    private bool ExecuteWorkbookCommand(DocumentCommand command, CellPosition? target = null, bool clearView = false)
    {
        if (_busy || !CommitEditors()) return false;
        try
        {
            var result = _editor.Execute(command);
            if (!result.Changed) return true;
            if (result.CalculationUnchanged)
            {
                if (command is SetAxisMetadataCommand)
                {
                    _grid.RefreshDimensions(_document.Snapshot);
                }
                else RefreshSheet();
                UpdateChrome();
                QueueRecalculation();
                return true;
            }
            _sheetIndex = Math.Min(_sheetIndex, _document.Snapshot.Sheets.Length - 1);
            if (clearView) { CurrentView.FilterRange = null; CurrentView.FrozenRows = CurrentView.FrozenColumns = 0; }
            SynchronizeSheets();
            var position = target ?? new CellPosition(_selectedRow, _selectedColumn);
            GridSurface.Select(new(Math.Clamp(position.Row, 0, _rowCount - 1), Math.Clamp(position.Column, 0, _columnCount - 1)));
            _calculationPending = true;
            QueueRecalculation();
            return true;
        }
        catch (Exception ex) { ReportEditFailure(ex); return false; }
    }

    private void Structure_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        try
        {
            if (CurrentView.FilterRange is not null) throw new InvalidOperationException("请先清除筛选，再进行结构操作。");
            var operation = ((FrameworkElement)sender).Tag.ToString()!;
            var sheet = _document.Snapshot.Sheets[_sheetIndex];
            if (operation == "rename")
            {
                var name = PromptText("重命名工作表", "工作表名称（同步更新公式引用）", sheet.Name);
                if (name is not null) ExecuteWorkbookCommand(new RenameSheetCommand(sheet.Id, name));
                return;
            }
            if (operation == "deleteSheet")
            {
                var decision = _smokeMode ? _smokeDiscardDecision ?? MessageBoxResult.Cancel : MessageBox.Show(this, $"删除“{sheet.Name}”？引用此表的公式将变为 #REF!，可撤销。", "删除工作表", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
                if (decision == MessageBoxResult.OK)
                    ExecuteWorkbookCommand(new DeleteSheetCommand(sheet.Id), new(0, 0));
                return;
            }
            var range = SelectedRange();
            var axis = operation.EndsWith("Row", StringComparison.Ordinal) ? SheetAxis.Row : SheetAxis.Column;
            var start = axis == SheetAxis.Row ? range.Top : range.Left;
            var count = axis == SheetAxis.Row ? range.Rows : range.Columns;
            if (operation.StartsWith("show", StringComparison.Ordinal))
            {
                // Hidden dimensions cannot own a native selection. Include contiguous hidden
                // neighbors so a visible header can restore them, including at sheet edges.
                bool Hidden(int index) => axis == SheetAxis.Row ? sheet.Rows.GetValueOrDefault(index)?.Hidden == true : sheet.Columns.GetValueOrDefault(index)?.Hidden == true;
                var end = start + count;
                while (start > 0 && Hidden(start - 1)) start--;
                var limit = axis == SheetAxis.Row ? _rowCount : _columnCount;
                while (end < limit && Hidden(end)) end++;
                count = end - start;
            }
            DocumentCommand command;
            var target = new CellPosition(range.Top, range.Left);
            if (operation.StartsWith("move", StringComparison.Ordinal))
            {
                var text = PromptText("移动行列", "输入移动后的起始序号（从 1 开始；引用随结构移动）", (start + 1).ToString());
                if (text is null) return;
                if (!int.TryParse(text, out var destination) || destination < 1 || destination > (axis == SheetAxis.Row ? _rowCount : _columnCount) - count + 1)
                    throw new ArgumentException("目标超出当前显示范围。");
                command = new MoveAxisCommand(sheet.Id, axis, start, count, destination - 1);
                target = axis == SheetAxis.Row ? target with { Row = destination - 1 } : target with { Column = destination - 1 };
            }
            else if (operation.StartsWith("size", StringComparison.Ordinal))
            {
                var text = PromptText("设置尺寸", axis == SheetAxis.Row ? "行高（1–400）" : "列宽（字符单位，1–56）", axis == SheetAxis.Row ? "31" : "16");
                if (text is null) return;
                if (!double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out var size) || !double.IsFinite(size) || size < 1 || size > (axis == SheetAxis.Row ? 400 : 56))
                    throw new ArgumentException("尺寸超出范围。");
                command = new SetAxisMetadataCommand(sheet.Id, axis, start, count, size);
            }
            else if (operation.StartsWith("hide", StringComparison.Ordinal) || operation.StartsWith("show", StringComparison.Ordinal))
                command = new SetAxisMetadataCommand(sheet.Id, axis, start, count, hidden: operation.StartsWith("hide", StringComparison.Ordinal));
            else command = new ShiftAxisCommand(sheet.Id, axis, start, count, operation.StartsWith("delete", StringComparison.Ordinal));
            ExecuteWorkbookCommand(command, target, clearView: true);
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private async void Sort_Click(object sender, RoutedEventArgs e)
        => await (_workspaceInput = SortSelectionAsync(((FrameworkElement)sender).Tag?.ToString() == "desc"));
    private async Task SortSelectionAsync(bool descending)
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            if (CurrentView.Region is { } region)
            {
                var column = Math.Clamp(_selectedColumn, region.Range.Left, region.Range.Right);
                await SortDataRegionAsync(region, [new(column, descending)]); return;
            }
            if (CurrentView.FilterRange is not null) throw new InvalidOperationException("请清除筛选后排序，避免隐藏行的参与范围不明确。");
            var range = SelectedRange();
            if (range.Rows < 2) throw new ArgumentException("请选中至少两行数据（不含表头）；按当前单元格所在列排序。");
            ExecuteWorkbookCommand(new SortRangeCommand(_document.Snapshot.Sheets[_sheetIndex].Id, new(range.Top, range.Left), new(range.Bottom, range.Right),
                _selectedColumn, descending), new(range.Top, _selectedColumn));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private void ApplyFilter(GridRange range, int column, string text)
    {
        CurrentView.FilterRange = range; CurrentView.FilterColumn = column; CurrentView.FilterText = text;
        RefreshSheet();
        GridSurface.Select(new(range.Top, column));
        StatusText.Text = "筛选已应用，首行为表头；复制仅含筛选可见行，粘贴不允许跨隐藏行。";
    }
    private async void Filter_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await FlushEditsAsync()) return;
        try
        {
            var range = SelectedRange();
            if (range.Rows >= 2) EstablishDataRegion(range);
            else if (CurrentView.Region is null) throw new ArgumentException("请选择含表头和数据的矩形，再建立筛选区域。");
            else await OpenHeaderFilterAsync(_selectedColumn, sender as FrameworkElement ?? SheetGrid);
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }
    private void ClearFilter_Click(object sender, RoutedEventArgs e)
    { if (!CommitEditors()) return; CurrentView.FilterRange = null; RefreshSheet(); GridSurface.Select(new(_selectedRow, _selectedColumn)); }

    private void Freeze_Click(object sender, RoutedEventArgs e)
    {
        if (!CommitEditors()) return;
        var clear = ((FrameworkElement)sender).Tag?.ToString() == "clear";
        if (!clear && _selectedRow > 10) { ReportEditFailure(new ArgumentException("冻结上方行最多 10 行，请选择第 11 行以内的单元格。")); return; }
        CurrentView.FrozenRows = clear ? 0 : _selectedRow;
        CurrentView.FrozenColumns = clear ? 0 : _selectedColumn;
        RefreshSheet();
        StatusText.Text = clear ? "已取消冻结。" : "已冻结左侧列和上方行；点击冻结行可定位到原单元格编辑。";
    }

    private void Style_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) { UpdateRibbonState(); return; }
        try
        {
            var snapshot = _document.Snapshot; var sheet = snapshot.Sheets[_sheetIndex];
            var style = snapshot.Styles[sheet.GetStyleId(_selectedRow, _selectedColumn)];
            var action = ((FrameworkElement)sender).Tag.ToString();
            CellStylePatch patch = action switch
            {
                "bold" => new(Bold: !style.Bold), "italic" => new(Italic: !style.Italic),
                "fill" => new(FillColor: "FFFFE599"), "font" => new(FontColor: "FF187553"),
                "left" => new(Horizontal: 1), "center" => new(Horizontal: 2), "right" => new(Horizontal: 3),
                "top" => new(Vertical: 0), "middle" => new(Vertical: 1), "bottom" => new(Vertical: 2),
                "wrap" => new(WrapText: !style.WrapText),
                "border" => new(Border: true), "noBorder" => new(Border: false),
                "number" => new(NumberFormatId: 2), "percent" => new(NumberFormatId: 10),
                "date" => new(NumberFormatId: 0, NumberFormatCode: "yyyy-mm-dd"),
                "general" => new(NumberFormatId: 0), "reset" => new(Reset: true), _ => throw new ArgumentException("未知样式。")
            };
            ApplyRibbonStyle(patch);
        }
        catch (Exception ex) { ReportEditFailure(ex); }
        finally { UpdateRibbonState(); }
    }

    private void Merge_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CommitEditors()) return;
        try
        {
            var range = SelectedRange();
            ExecuteWorkbookCommand(new MergeCellsCommand(_document.Snapshot.Sheets[_sheetIndex].Id,
                new(range.Top, range.Left, range.Bottom, range.Right), (sender as FrameworkElement)?.Tag as string == "unmerge"));
        }
        catch (Exception ex) { ReportEditFailure(ex); }
    }

    private void RecentFiles_Opened(object sender, RoutedEventArgs e)
    {
        RecentFilesMenu.Items.Clear();
        foreach (var path in _recentFiles.Paths)
        {
            var item = new MenuItem { Header = path.Replace("_", "__"), ToolTip = path };
            item.Click += async (_, _) => { if (!_busy && await ConfirmDiscardAsync()) await OpenPathAsync(path); };
            RecentFilesMenu.Items.Add(item);
        }
        if (_recentFiles.Paths.Count == 0) RecentFilesMenu.Items.Add(new MenuItem { Header = "暂无最近文件", IsEnabled = false });
        else
        {
            RecentFilesMenu.Items.Add(new Separator());
            var clear = new MenuItem { Header = "清空记录" };
            clear.Click += (_, _) => _recentFiles.Clear();
            RecentFilesMenu.Items.Add(clear);
        }
    }
    private void RememberFile(string path)
    {
        _recentFiles.Remember(path);
        if (_recentFiles.Error is { } error) AppendLog("最近文件记录未保存：" + error);
    }
    private async void CloseDocument_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !await ConfirmDiscardAsync()) return;
        InstallDocument(WorkbookDocument.CreateBlank(_services), null, "未命名.xlsq");
        _hasWorkbook = false; ShowStartPage();
    }
    private void ToggleScriptPanel_Click(object sender, RoutedEventArgs e)
    {
        if (!CommitEditors()) return;
        ShowWorkbookPage(ScriptPanel.Visibility != Visibility.Visible);
    }

    private async Task<bool> SaveToPathAsync(string path, bool discardUnsupported, FileVersion expected)
    {
        FileVersion? savedVersion = null;
        if (!await WorkAsync("保存 " + Path.GetExtension(path).ToLowerInvariant() + "…", () => savedVersion = _document.SaveChecked(path, expected, discardUnsupported))) return false;
        _filePath = Path.GetFullPath(path); _fileVersion = savedVersion; _versionPath = _filePath;
        _displayName = Path.GetFileName(path);
        RememberFile(_filePath);
        UpdateChrome();
        return true;
    }
}
