using System.Collections.Immutable;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

internal sealed record FilterChoice(string Action, ColumnFilter? Filter = null);
internal sealed class FilterDialog : Window
{
    public FilterDialog(string title, IReadOnlyList<FilterCandidate> candidates, ColumnFilter? current)
    {
        Title = title; Width = 570; Height = 650; MinWidth = 500; WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var selected = candidates.Where(c => c.Selected).Select(c => c.Key).ToHashSet();
        var panel = new StackPanel { Margin = new(20), Spacing = 10 };
        var mode = new ComboBox { ItemsSource = new[] { "候选值", "自定义条件" }, SelectedIndex = current?.First is not null ? 1 : 0 }; AutomationProperties.SetName(mode, "筛选方式"); panel.Children.Add(mode);
        var values = new StackPanel { Spacing = 8 }; panel.Children.Add(values);
        var search = new TextBox { PlaceholderText = "搜索候选值（保留其他勾选）" }; AutomationProperties.SetName(search, "搜索候选值"); values.Children.Add(search);
        var list = new ListBox { Height = 310, ItemTemplate = new FuncDataTemplate<FilterCandidate>((candidate, _) =>
        {
            var box = new CheckBox { Content = $"{candidate!.Label} ({candidate.Count})", IsChecked = selected.Contains(candidate.Key) };
            box.IsCheckedChanged += (_, _) => { if (box.IsChecked == true) selected.Add(candidate.Key); else selected.Remove(candidate.Key); }; return box;
        }) }; values.Children.Add(list);
        FilterCandidate[] Shown() => candidates.Where(c => c.Label.Contains(search.Text ?? "", StringComparison.OrdinalIgnoreCase)).ToArray();
        void Refresh() => list.ItemsSource = Shown();
        search.PropertyChanged += (_, e) => { if (e.Property == TextBox.TextProperty) Refresh(); }; Refresh();
        var selectButtons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 }; values.Children.Add(selectButtons);
        foreach (var select in new[] { true, false })
        { var button = new Button { Content = select ? "勾选搜索结果" : "取消搜索结果" }; button.Click += (_, _) => { foreach (var c in Shown()) { if (select) selected.Add(c.Key); else selected.Remove(c.Key); } list.ItemsSource = null; Refresh(); }; selectButtons.Children.Add(button); }
        var conditions = new StackPanel { Spacing = 10 }; panel.Children.Add(conditions);
        string[] operators = ["文本等于", "文本包含", "文本开头", "文本结尾", "数字等于", "大于", "大于等于", "小于", "小于等于", "日期不早于", "日期不晚于", "Excel 等于", "Excel 不等于", "Excel 小于", "Excel 小于等于", "Excel 大于", "Excel 大于等于"];
        ComboBox Operator(string name, FilterTerm? term) { var box = new ComboBox { ItemsSource = operators, SelectedIndex = (int)(term?.Operator ?? FilterOperator.TextContains) }; AutomationProperties.SetName(box, name); conditions.Children.Add(box); return box; }
        var first = Operator("第一个条件", current?.First); var value1 = new TextBox { Text = current?.First?.Operand ?? "" }; AutomationProperties.SetName(value1, "第一个条件值"); conditions.Children.Add(value1);
        var combine = new ComboBox { ItemsSource = new[] { "仅第一个条件", "并且", "或者" }, SelectedIndex = current?.Second is null ? 0 : current.UseOr ? 2 : 1 }; AutomationProperties.SetName(combine, "组合条件"); conditions.Children.Add(combine);
        var second = Operator("第二个条件", current?.Second); var value2 = new TextBox { Text = current?.Second?.Operand ?? "" }; AutomationProperties.SetName(value2, "第二个条件值"); conditions.Children.Add(value2);
        conditions.Children.Add(new TextBlock { Text = "数字使用小数点；日期使用 yyyy-MM-dd。", TextWrapping = TextWrapping.Wrap });
        void Mode() { values.IsVisible = mode.SelectedIndex == 0; conditions.IsVisible = mode.SelectedIndex == 1; }
        mode.SelectionChanged += (_, _) => Mode(); Mode();
        var error = new TextBlock { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap }; panel.Children.Add(error);
        var buttons = new WrapPanel(); panel.Children.Add(buttons);
        void Button(string label, Action action) { var b = new Button { Content = label, Margin = new(3) }; b.Click += (_, _) => { try { action(); } catch (Exception ex) { error.Text = ex.Message; } }; buttons.Children.Add(b); }
        Button("应用", () =>
        {
            var filter = mode.SelectedIndex == 0 ? new ColumnFilter(Values: selected.ToImmutableHashSet()) : new(First: new((FilterOperator)first.SelectedIndex, value1.Text ?? ""),
                Second: combine.SelectedIndex == 0 ? null : new((FilterOperator)second.SelectedIndex, value2.Text ?? ""), UseOr: combine.SelectedIndex == 2);
            filter.Validate(); if (filter.Values is { Count: 0 }) throw new ArgumentException("请至少选择一个值或空白。"); Close(new FilterChoice("apply", filter));
        });
        Button("清除此列", () => Close(new FilterChoice("clear"))); Button("清除全部条件", () => Close(new FilterChoice("clear-all")));
        Button("升序", () => Close(new FilterChoice("ascending"))); Button("降序", () => Close(new FilterChoice("descending"))); Button("取消", () => Close(null));
        Content = new ScrollViewer { Content = panel };
    }
}
