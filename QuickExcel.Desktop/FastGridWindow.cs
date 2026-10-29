using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>Opt-in migration workspace. The parent is modal-disabled; the same document session owns every write.</summary>
internal sealed class FastGridWindow : Window
{
    private readonly WorkbookDocument _document;
    private readonly EditorController _controller;
    private readonly int _sheetIndex;
    private readonly SheetViewState _view;
    internal FastGridSurface Surface { get; }
    private readonly TextBlock _status = new() { TextWrapping = TextWrapping.Wrap, Margin = new(8) };
    private bool _busy, _pending, _closed;
    private FilterPopup? _popup;

    public FastGridWindow(Window owner, WorkbookDocument document, int sheetIndex, SheetViewState view, int rows, int columns)
    {
        Owner = owner; Title = "QuickExcel · FastWpfGrid 试验工作区"; Width = 1150; Height = 760; WindowStartupLocation = WindowStartupLocation.CenterOwner;
        _document = document; _controller = new(document); _sheetIndex = sheetIndex; _view = view;
        Surface = new((position, text) =>
        {
            try
            {
                _pending |= _controller.SetCell(_sheetIndex, position.Row, position.Column, text).Changed;
                Dispatcher.BeginInvoke(new Action(async () =>
                {
                    if (_closed || _busy) return;
                    try { await RefreshAsync(); } catch (Exception ex) { _status.Text = ex.Message; }
                })); return true;
            }
            catch (Exception ex) { _status.Text = ex.Message; return false; }
        }) { View = view };
        var panel = new DockPanel(); var toolbar = new WrapPanel { Margin = new(6) };
        DockPanel.SetDock(toolbar, Dock.Top); panel.Children.Add(toolbar);
        void Button(string text, Func<Task> action)
        {
            var button = new Button { Content = text, Margin = new(3), Padding = new(8, 5, 8, 5) };
            button.Click += async (_, _) => { if (!_busy) try { await action(); } catch (Exception ex) { _status.Text = ex.Message; } };
            toolbar.Children.Add(button);
        }
        Button("建立筛选区域", async () =>
        {
            // Capture before refresh, which may reduce the view selection to its current cell.
            var range = Rectangle();
            if (!await FlushAsync()) return;
            _view.SetRegion(new(Guid.NewGuid(), _document.Snapshot.Sheets[_sheetIndex].Id, range)); Surface.Refresh(_document.Snapshot);
        });
        Button("当前列筛选", () => FilterAsync(Surface.Current.Column, Surface.Editor));
        Button("多级排序", SortAsync);
        Button("清除条件", async () => { if (!await FlushAsync()) return; _view.ClearConditions(); Surface.Refresh(_document.Snapshot); });
        Button("撤销", () => HistoryAsync(false)); Button("重做", () => HistoryAsync(true));
        Button("冻结到当前格", async () =>
        {
            if (!await FlushAsync()) return;
            if (Surface.Current.Row > 10) throw new ArgumentException("冻结上方行最多 10 行。");
            _view.FrozenRows = Surface.Current.Row; _view.FrozenColumns = Surface.Current.Column; Surface.Refresh(_document.Snapshot);
        });
        Button("取消冻结", async () => { if (!await FlushAsync()) return; _view.FrozenRows = _view.FrozenColumns = 0; Surface.Refresh(_document.Snapshot); });
        var zoom = new ComboBox { ItemsSource = new[] { "75%", "100%", "125%", "150%", "175%" }, SelectedIndex = 1, Margin = new(6), MinWidth = 75 };
        zoom.SelectionChanged += (_, _) => { if (Surface.Commit()) Surface.Element.LayoutTransform = new System.Windows.Media.ScaleTransform(new[] { .75, 1, 1.25, 1.5, 1.75 }[zoom.SelectedIndex], new[] { .75, 1, 1.25, 1.5, 1.75 }[zoom.SelectedIndex]); };
        toolbar.Children.Add(zoom);
        var note = new TextBlock { Text = "试验工作区：改动进入当前文档，可撤销；关闭后回到主界面保存。复杂样式、合并/图片、跨屏 DPI 和辅助功能尚未完成对等验收。", Margin = new(8), TextWrapping = TextWrapping.Wrap };
        DockPanel.SetDock(note, Dock.Top); panel.Children.Add(note);
        DockPanel.SetDock(_status, Dock.Bottom); panel.Children.Add(_status); panel.Children.Add(Surface.Element); Content = panel;
        Surface.HeaderRequested += async (column, target) => { if (!_busy) try { await FilterAsync(column, target); } catch (Exception ex) { _status.Text = ex.Message; } };
        Surface.CommandRequested += async (key, modifiers) => { if (!_busy) try { await CommandAsync(key, modifiers); } catch (Exception ex) { _status.Text = ex.Message; } };
        Surface.SelectionChanged += () => _status.Text = $"{MainWindow.ColumnName(Surface.Current.Column)}{Surface.Current.Row + 1} · {Surface.Control.SelectedCellCount:N0} 格";
        Surface.Control.ViewportChanged += (_, _) => _popup?.Close();
        Deactivated += (_, _) => _popup?.Close();
        PreviewMouseDown += (_, e) =>
        {
            if (_busy || Surface.IsComposing && e.OriginalSource is DependencyObject target && !ReferenceEquals(target, Surface.Editor) && !Surface.Editor.IsAncestorOf(target)) e.Handled = true;
        };
        Closing += (_, e) =>
        {
            if (_busy || !Surface.Commit()) { e.Cancel = true; _status.Text = "请先完成或取消当前输入。"; return; }
            _popup?.Close();
        };
        Closed += (_, _) => _closed = true;
        Loaded += (_, _) => Surface.Bind(_document.Snapshot, _sheetIndex, rows, columns);
    }
    private GridRange Rectangle()
    {
        var ranges = Surface.SelectionRanges;
        if (ranges.Count != 1) throw new ArgumentException("请选择连续矩形；筛选隐藏行会打断显示选区。已有区域请从标题按钮操作。");
        return ranges[0];
    }
    private async Task<bool> FlushAsync()
    {
        if (_busy || !Surface.Commit()) return false;
        await RefreshAsync(); return true;
    }
    private async Task RefreshAsync()
    {
        if (_busy || _closed) return;
        _busy = true;
        try
        {
            if (_pending) { _status.Text = "重算并更新筛选…"; Surface.Element.IsEnabled = false; await Task.Run(() => _document.Calculate()); _pending = false; }
            Surface.Refresh(_document.Snapshot);
            if (_view.FilterError is not null) _status.Text = _view.FilterError + "（暂显示全部行）";
        }
        finally { Surface.Element.IsEnabled = true; _busy = false; }
    }
    private async Task HistoryAsync(bool redo)
    {
        if (!await FlushAsync()) return;
        _pending |= (redo ? _controller.Redo() : _controller.Undo()).Changed; await RefreshAsync();
    }
    private async Task FilterAsync(int column, FrameworkElement target)
    {
        if (!await FlushAsync()) return;
        var region = _view.Region ?? throw new InvalidOperationException("请先选择含标题行的区域并建立筛选区域。");
        var snapshot = _document.Snapshot;
        var candidates = DataFilterEngine.Candidates(snapshot, region, _view.Filters, column);
        _popup?.Close(); _popup = new(target, new Rect(target.RenderSize), $"{MainWindow.ColumnName(column)} 列", candidates, _view.Filters.GetValueOrDefault(column));
        var result = await _popup.ShowAsync(); _popup = null;
        if (result is null) return;
        if (!ReferenceEquals(snapshot, _document.Snapshot) || _view.Region != region) throw new InvalidOperationException("数据已变化，请重新打开筛选面板。");
        if (result.Action == FilterPopupAction.AdvancedSort) { await SortAsync(); return; }
        if (result.Action is FilterPopupAction.SortAscending or FilterPopupAction.SortDescending)
        { await SortRegionAsync(region, [new(column, result.Action == FilterPopupAction.SortDescending)]); return; }
        if (result.Action == FilterPopupAction.ClearAll) _view.ClearConditions();
        else _view.Apply(snapshot, column, result.Action == FilterPopupAction.ClearColumn ? null : result.Filter);
        Surface.Refresh(_document.Snapshot);
    }
    private async Task SortAsync()
    {
        var selected = _view.Region?.Range ?? Rectangle();
        if (!await FlushAsync()) return;
        var region = _view.Region;
        var dialog = new SortDialog(this, _document.Snapshot, _document.Snapshot.Sheets[_sheetIndex].Id, selected, _view.SortKeys, region?.HasHeader ?? true, region is not null);
        if (dialog.ShowDialog() != true) return;
        await SortRegionAsync(region ?? new(Guid.NewGuid(), _document.Snapshot.Sheets[_sheetIndex].Id, selected, dialog.HasHeader), dialog.Keys);
    }
    private async Task SortRegionAsync(DataRegion region, IReadOnlyList<SortKey> keys)
    {
        var range = region.Range;
        _pending |= _controller.Execute(new SortRangeCommand(new(region.SheetId, new(range.Top, range.Left), new(range.Bottom, range.Right), keys, region.HasHeader))).Changed;
        _view.SortKeys = keys.ToArray(); await RefreshAsync();
    }
    private async Task CommandAsync(Key key, ModifierKeys modifiers)
    {
        if (key is Key.Z or Key.Y) { await HistoryAsync(key == Key.Y); return; }
        if (key == Key.C)
        {
            _view.RequireCurrentFilter(_document.Snapshot);
            var selected = Surface.Selection;
            bool Excluded(int row) => _view.IsFilteredOut(_document.Snapshot, _sheetIndex, row) || _document.Snapshot.Sheets[_sheetIndex].Rows.GetValueOrDefault(row)?.Hidden == true;
            var range = GridRange.Rectangle(selected, Excluded);
            Clipboard.SetText(GridClipboard.Write(range, (r, c) => _document.GetCellValue(_sheetIndex, r, c), Excluded)); return;
        }
        var ranges = Surface.SelectionRanges;
        if (key == Key.Delete)
        {
            if (ranges.Sum(r => r.Count) > GridClipboard.MaxEditCells) throw new ArgumentException("单次清空上限为 20,000 格。");
            _view.RequireVisibleTargets(_document.Snapshot, _sheetIndex, Surface.Selection.Select(p => p.Row));
            _pending |= _controller.ClearCells(_sheetIndex, Surface.Selection.Select(p => (p.Row, p.Column))).Changed;
        }
        else if (key == Key.V && Clipboard.ContainsText())
        {
            var edits = GridClipboard.Prepare(Clipboard.GetText(), Rectangle(), modifiers.HasFlag(ModifierKeys.Shift));
            _view.RequireVisibleTargets(_document.Snapshot, _sheetIndex, edits.Select(edit => edit.Row));
            _pending |= _controller.SetCells(_sheetIndex, edits).Changed;
        }
        await RefreshAsync();
    }
}
