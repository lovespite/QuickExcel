using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Grid.Avalonia;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class SortDialog : Window
{
    public SortDescriptor? Descriptor { get; private set; }
    public SortDialog(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, CellRectangle range, bool hasHeader,
        IEnumerable<SortKey> initial, bool fixedHeader, bool expandColumns)
    {
        Title = "多级排序"; Width = 850; Height = 560; MinWidth = 650; WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var panel = new StackPanel { Spacing = 10, Margin = new(20) };
        panel.Children.Add(new TextBlock { Text = $"区域：{range}。级别从上到下依次比较；整个记录随排序移动。", TextWrapping = TextWrapping.Wrap });
        var header = new CheckBox { Content = "首行为标题", IsChecked = hasHeader, IsEnabled = !fixedHeader };
        var expand = new CheckBox { Content = "扩展到相邻数据列", IsChecked = expandColumns, IsEnabled = expandColumns };
        var matchCase = new CheckBox { Content = "区分大小写", IsChecked = initial.FirstOrDefault()?.CaseSensitive == true };
        var settings = new WrapPanel(); settings.Children.Add(header); settings.Children.Add(expand); settings.Children.Add(matchCase); panel.Children.Add(settings);
        var levels = new StackPanel { Spacing = 8 }; panel.Children.Add(new ScrollViewer { Content = levels, Height = 340 });
        var error = new TextBlock { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap }; panel.Children.Add(error);
        var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 }; panel.Children.Add(actions);
        var columns = Enumerable.Range(range.Left, range.Columns).Select(c => WorkbookGrid.ColumnName(c) + " · " + CellInput.DisplayText(sheet.GetCell(range.Top, c)?.Value ?? CellValue.Blank)).ToArray();
        var rows = new List<(Border View, Func<SortKey> Read)>();
        void Add(SortKey? key = null)
        {
            if (rows.Count >= 32) { error.Text = "最多 32 个排序级别。"; return; }
            ComboBox Choice(string label, string[] choices, int index, double width)
            { var box = new ComboBox { ItemsSource = choices, SelectedIndex = index, Width = width }; AutomationProperties.SetName(box, label); return box; }
            var column = Choice("排序列", columns, Math.Clamp((key?.Column ?? range.Left) - range.Left, 0, columns.Length - 1), 170);
            var order = Choice("顺序", ["升序 / 指定颜色在前", "降序 / 指定颜色在后"], key?.Descending == true ? 1 : 0, 180);
            var basis = Choice("依据", ["值", "填充颜色", "字体颜色"], (int)(key?.Basis ?? SortBasis.Value), 130);
            var color = new TextBox { PlaceholderText = "ARGB 色值", Text = key?.PreferredColor ?? "FF000000", Width = 110 }; AutomationProperties.SetName(color, "首选颜色 ARGB");
            var custom = new TextBox { PlaceholderText = "自定义文本顺序（用逗号分隔）", Text = key is null ? "" : string.Join(',', key.CustomOrder), Width = 500 }; AutomationProperties.SetName(custom, "自定义文本顺序");
            var row = new WrapPanel { Margin = new(6) }; row.Children.Add(column); row.Children.Add(order); row.Children.Add(basis); row.Children.Add(color); row.Children.Add(custom);
            var border = new Border { Child = row, BorderBrush = Brushes.LightGray, BorderThickness = new(1) };
            var remove = new Button { Content = "删除级别" }; row.Children.Add(remove);
            var up = new Button { Content = "上移" }; row.Children.Add(up);
            remove.Click += (_, _) => { rows.RemoveAll(r => r.View == border); levels.Children.Remove(border); };
            up.Click += (_, _) =>
            { var index = rows.FindIndex(r => r.View == border); if (index <= 0) return; var item = rows[index]; rows.RemoveAt(index); rows.Insert(index - 1, item); levels.Children.Remove(border); levels.Children.Insert(index - 1, border); };
            rows.Add((border, () => new(range.Left + column.SelectedIndex, order.SelectedIndex == 1, matchCase.IsChecked == true,
                (SortBasis)basis.SelectedIndex, basis.SelectedIndex == 0 ? null : color.Text?.Trim(),
                basis.SelectedIndex == 0 && !string.IsNullOrWhiteSpace(custom.Text) ? custom.Text.Split(',') : null)));
            levels.Children.Add(border);
        }
        foreach (var key in initial) Add(key); if (rows.Count == 0) Add();
        var add = new Button { Content = "添加级别" }; add.Click += (_, _) => Add(); actions.Children.Add(add);
        var ok = new Button { Content = "排序", IsDefault = true }; actions.Children.Add(ok);
        var cancel = new Button { Content = "取消", IsCancel = true }; actions.Children.Add(cancel); cancel.Click += (_, _) => Close(false);
        ok.Click += (_, _) =>
        {
            try
            {
                var keys = rows.Select(r => r.Read()).ToArray();
                if (keys.Length == 0 || keys.Select(k => k.Column).Distinct().Count() != keys.Length) throw new ArgumentException("请选择至少一个不重复的排序列。");
                // Validate persistent sort choices without altering the document.
                _ = new WorksheetSortState(range with { Top = range.Top + (header.IsChecked == true ? 1 : 0) }, keys);
                Descriptor = new(sheet.Id, new(range.Top, range.Left), new(range.Bottom, range.Right), keys, header.IsChecked == true, expand.IsChecked == true);
                Close(true);
            }
            catch (Exception ex) { error.Text = ex.Message; }
        };
        Content = panel;
    }
}
