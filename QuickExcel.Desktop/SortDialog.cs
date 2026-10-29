using System.Windows;
using System.Windows.Controls;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal sealed class SortDialog : Window
{
    private readonly StackPanel _levels = new();
    private readonly List<(ComboBox Column, ComboBox Direction, ComboBox Basis, CheckBox Case, TextBox Option)> _keys = [];
    public IReadOnlyList<SortKey> Keys { get; private set; } = [];
    public bool HasHeader => _header.IsChecked == true;
    private readonly CheckBox _header = new() { Content = "首行为标题，不参与排序", Margin = new(0, 6, 0, 6) };

    public SortDialog(Window owner, WorkbookSnapshot snapshot, SheetId sheetId, GridRange range, IReadOnlyList<SortKey> existing, bool header, bool fixedHeader)
    {
        Owner = owner; Title = "多级排序"; Width = 830; Height = 500; WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var sheet = snapshot.GetSheet(sheetId);
        var columns = Enumerable.Range(range.Left, range.Columns).Select(c => $"{MainWindow.ColumnName(c)} · {CellInput.DisplayText(sheet.GetCell(range.Top, c)?.Value ?? CellValue.Blank)}").ToArray();
        var panel = new DockPanel { Margin = new(16) };
        var top = new StackPanel(); DockPanel.SetDock(top, Dock.Top); panel.Children.Add(top);
        top.Children.Add(new TextBlock { Text = "排序移动区域内完整记录，筛选隐藏和手动隐藏的记录均参与。行高、隐藏状态留在原地址。", TextWrapping = TextWrapping.Wrap });
        _header.IsChecked = header; _header.IsEnabled = !fixedHeader; top.Children.Add(_header);
        top.Children.Add(new TextBlock { Text = "顺序：列 / 方向 / 依据 / 大小写 / 选项。值排序可填自定义文本列表（用 | 分隔）；颜色填 ARGB。", TextWrapping = TextWrapping.Wrap });
        var add = new Button { Content = "添加级别", HorizontalAlignment = HorizontalAlignment.Left, Margin = new(0, 6, 0, 6) };
        top.Children.Add(add);
        void Add(SortKey key)
        {
            if (_keys.Count >= 32) return;
            var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new(0, 4, 0, 4) };
            var column = new ComboBox { ItemsSource = columns, SelectedIndex = Math.Clamp(key.Column - range.Left, 0, columns.Length - 1), Width = 155 };
            var direction = new ComboBox { ItemsSource = new[] { "升序 / 置顶", "降序 / 置底" }, SelectedIndex = key.Descending ? 1 : 0, Width = 95 };
            var basis = new ComboBox { ItemsSource = new[] { "值", "填充色", "字体色" }, SelectedIndex = (int)key.Basis, Width = 75 };
            var sensitive = new CheckBox { Content = "区分大小写", IsChecked = key.CaseSensitive, VerticalAlignment = VerticalAlignment.Center, Margin = new(5, 0, 5, 0) };
            var option = new TextBox { Text = key.PreferredColor ?? string.Join("|", key.CustomOrder), Width = 150 };
            var remove = new Button { Content = "删除", Margin = new(4, 0, 0, 0) };
            var up = new Button { Content = "↑", Margin = new(4, 0, 0, 0) };
            var down = new Button { Content = "↓", Margin = new(4, 0, 0, 0) };
            var controls = (column, direction, basis, sensitive, option);
            remove.Click += (_, _) => { _keys.Remove(controls); _levels.Children.Remove(row); };
            void Move(int delta)
            {
                var index = _keys.IndexOf(controls); var next = index + delta;
                if (next < 0 || next >= _keys.Count) return;
                _keys.RemoveAt(index); _keys.Insert(next, controls); _levels.Children.RemoveAt(index); _levels.Children.Insert(next, row);
            }
            up.Click += (_, _) => Move(-1); down.Click += (_, _) => Move(1);
            foreach (var item in new UIElement[] { column, direction, basis, sensitive, option, up, down, remove }) row.Children.Add(item);
            _keys.Add(controls); _levels.Children.Add(row);
        }
        add.Click += (_, _) => Add(new(range.Left + Math.Min(_keys.Count, range.Columns - 1)));
        foreach (var key in existing.Count > 0 ? existing : new[] { new SortKey(range.Left) }) Add(key);
        var bottom = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        DockPanel.SetDock(bottom, Dock.Bottom); panel.Children.Add(bottom);
        var error = new TextBlock { Foreground = System.Windows.Media.Brushes.Firebrick, MaxWidth = 500, TextWrapping = TextWrapping.Wrap }; bottom.Children.Add(error);
        var ok = new Button { Content = "排序", MinWidth = 80, IsDefault = true, Margin = new(6) };
        var cancel = new Button { Content = "取消", MinWidth = 80, IsCancel = true, Margin = new(6) };
        bottom.Children.Add(ok); bottom.Children.Add(cancel);
        ok.Click += (_, _) =>
        {
            if (_keys.Count == 0) { error.Text = "请添加至少一个排序级别。"; return; }
            Keys = _keys.Select(k => new SortKey(range.Left + k.Column.SelectedIndex, k.Direction.SelectedIndex == 1, k.Case.IsChecked == true,
                (SortBasis)k.Basis.SelectedIndex, k.Basis.SelectedIndex == 0 ? null : k.Option.Text.Trim(),
                k.Basis.SelectedIndex == 0 && k.Option.Text.Length > 0 ? k.Option.Text.Split('|') : null)).ToArray();
            DialogResult = true;
        };
        panel.Children.Add(new ScrollViewer { Content = _levels, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }); Content = panel;
    }
}
