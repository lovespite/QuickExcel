using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal sealed class SubtotalDialog : Window
{
    internal TextBox GroupColumn { get; } = new() { Text = "1" };
    internal TextBox TotalColumns { get; } = new() { Text = "2" };
    internal ComboBox Function { get; } = new() { ItemsSource = new[] { "平均值", "数值计数", "非空计数", "最大值", "最小值", "乘积", "样本标准差", "总体标准差", "求和", "样本方差", "总体方差" }, SelectedIndex = 8 };
    internal Button PreviewButton { get; } = new() { Content = "生成预览", MinWidth = 110, Margin = new(4) };
    internal Button ApplyButton { get; } = new() { Content = "插入分类汇总", MinWidth = 130, Margin = new(4), IsEnabled = false, IsDefault = true };
    internal SubtotalPreview? Preview { get; private set; }
    private sealed record Row(string Detail, string Summary, string Label);
    public SubtotalDialog(Window owner, WorkbookSnapshot book, SheetId id, CellRectangle range)
    {
        Owner = owner; Title = "分类汇总"; Width = 680; Height = 610; WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var root = new DockPanel { Margin = new(16) }; var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(PreviewButton); buttons.Children.Add(ApplyButton); buttons.Children.Add(new Button { Content = "取消", IsCancel = true, MinWidth = 85, Margin = new(4) });
        DockPanel.SetDock(buttons, Dock.Bottom); root.Children.Add(buttons); var panel = new StackPanel(); root.Children.Add(new ScrollViewer { Content = panel, VerticalScrollBarVisibility = ScrollBarVisibility.Auto });
        panel.Children.Add(new TextBlock { Text = $"{book.GetSheet(id).Name} · {range}\n首行为标题。按分类列连续变化插入汇总和总计，不自动排序。插入整行会移动下方所有列及其引用，应用后可一次撤销。", TextWrapping = TextWrapping.Wrap });
        void Field(string label, Control control) { panel.Children.Add(new TextBlock { Text = label, Margin = new(0, 8, 0, 3) }); panel.Children.Add(control); }
        Field("分类列（相对选区，从 1 开始）", GroupColumn); Field("汇总列（相对选区，逗号分隔）", TotalColumns); Field("汇总函数", Function);
        var summary = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new(0, 10, 0, 6) };
        var grid = new DataGrid { IsReadOnly = true, AutoGenerateColumns = false, CanUserAddRows = false, Height = 220 };
        foreach (var (name, property) in new[] { ("插入后的明细行", "Detail"), ("汇总行", "Summary"), ("分类", "Label") })
            grid.Columns.Add(new DataGridTextColumn { Header = name, Binding = new Binding(property), Width = 170 });
        panel.Children.Add(summary); panel.Children.Add(grid);
        void Reset() { Preview = null; ApplyButton.IsEnabled = false; grid.ItemsSource = null; summary.Text = "请生成预览。"; }
        GroupColumn.TextChanged += (_, _) => Reset(); TotalColumns.TextChanged += (_, _) => Reset(); Function.SelectionChanged += (_, _) => Reset();
        PreviewButton.Click += (_, _) =>
        {
            Reset();
            try
            {
                Preview = SubtotalCommands.Preview(book, id, range, range.Left + int.Parse(GroupColumn.Text) - 1,
                    TotalColumns.Text.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).Select(s => range.Left + int.Parse(s) - 1), Function.SelectedIndex + 1);
                summary.Text = $"新增 {Preview.InsertedRows} 整行，{Preview.Groups.Length} 个分类；总计在第 {Preview.GrandTotalRow + 1} 行。明细为第 2 层分组，分类汇总为第 1 层。";
                grid.ItemsSource = Preview.Groups.Select(g => new Row($"{g.FirstRow + 1}–{g.LastRow + 1}", (g.SummaryRow + 1).ToString(), g.Label)).ToArray(); ApplyButton.IsEnabled = true;
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or NotSupportedException or FormatException or OverflowException)
            { summary.Text = ex.Message; }
        };
        ApplyButton.Click += (_, _) => { if (Preview is not null) DialogResult = true; }; Content = root;
    }
}
