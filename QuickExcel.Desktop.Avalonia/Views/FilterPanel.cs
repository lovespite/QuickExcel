using System.Collections.Immutable;
using System.ComponentModel;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

internal sealed record FilterChoice(string Action, ColumnFilter? Filter = null);

/// <summary>A disposable column menu. Candidate state belongs to models, never recycled visual containers.</summary>
internal sealed class FilterPanel : Border
{
    private sealed class CandidateItem(FilterCandidate candidate) : INotifyPropertyChanged
    {
        public FilterCandidate Candidate { get; } = candidate;
        private bool _selected = candidate.Selected;
        public bool Selected
        {
            get => _selected;
            set { if (_selected == value) return; _selected = value; PropertyChanged?.Invoke(this, new(nameof(Selected))); }
        }
        public event PropertyChangedEventHandler? PropertyChanged;
    }

    public event Action<FilterChoice?>? Completed;
    private readonly TextBlock _error = new() { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap, IsVisible = false };
    private readonly TextBox _search = new() { Name = "FilterSearch", PlaceholderText = "搜索" };
    private TextBox? _firstValue;
    public void FocusSearch() { if (_search.IsEffectivelyVisible && _search.IsEffectivelyEnabled) _search.Focus(); else _firstValue?.Focus(); }

    public FilterPanel(string title, IReadOnlyList<FilterCandidate> candidates, ColumnFilter? current, string? candidateError = null)
    {
        Width = 340; Padding = new(10); Background = Brushes.White; BorderBrush = Brushes.LightGray; BorderThickness = new(1);
        CornerRadius = new(4); AutomationProperties.SetName(this, title);
        var items = candidates.Select(c => new CandidateItem(c)).ToArray();
        var panel = new StackPanel { Spacing = 5 };
        panel.Children.Add(new TextBlock { Text = title, FontWeight = FontWeight.SemiBold, Margin = new(5, 0, 5, 4) });
        Button Command(string name, string label, Action action, bool menu = true)
        {
            var button = new Button { Name = name, Content = label, HorizontalAlignment = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Left, Padding = new(10, 5) };
            if (menu) { button.Background = Brushes.Transparent; button.BorderThickness = new(0); }
            AutomationProperties.SetName(button, label); button.Click += (_, _) => Guard(action); return button;
        }
        panel.Children.Add(Command("FilterAscending", "↑  升序", () => Completed?.Invoke(new("ascending"))));
        panel.Children.Add(Command("FilterDescending", "↓  降序", () => Completed?.Invoke(new("descending"))));
        panel.Children.Add(new Separator());
        var clear = Command("FilterClear", "清除此列的筛选", () => Completed?.Invoke(new("clear"))); clear.IsEnabled = current is not null; panel.Children.Add(clear);
        panel.Children.Add(Command("FilterClearAll", "清除全部筛选条件", () => Completed?.Invoke(new("clear-all"))));
        var custom = new ToggleButton { Name = "FilterCustom", Content = "自定义条件筛选", IsChecked = current?.First is not null,
            Background = Brushes.Transparent, BorderThickness = new(0),
            HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Left, Padding = new(10, 5) };
        panel.Children.Add(custom); panel.Children.Add(new Separator());
        var values = new StackPanel { Spacing = 4 }; panel.Children.Add(values);
        AutomationProperties.SetName(_search, "搜索筛选值"); values.Children.Add(_search);
        // Mixed is a display state; user toggles must alternate checked/unchecked rather than cycle through null.
        var all = new CheckBox { Name = "FilterSelectAll", Content = "(全选)", IsThreeState = false, Margin = new(8, 0) }; values.Children.Add(all);
        var list = new ListBox { Name = "FilterCandidates", Height = 220, SelectionMode = SelectionMode.Single,
            Background = Brushes.White, BorderBrush = Brushes.LightGray, BorderThickness = new(1),
            ItemTemplate = new FuncDataTemplate<CandidateItem>((item, _) =>
            {
                // Avalonia may ask a template to build with null while probing or clearing its data context.
                if (item is null) return null;
                var box = new CheckBox { Content = item.Candidate.Label, Margin = new(0), MinHeight = 26 };
                ToolTip.SetTip(box, $"{item.Candidate.Label}（{item.Candidate.Count:N0} 项）");
                box.Bind(ToggleButton.IsCheckedProperty, new Binding(nameof(CandidateItem.Selected)) { Source = item, Mode = BindingMode.TwoWay });
                return box;
            }) };
        ScrollViewer.SetHorizontalScrollBarVisibility(list, ScrollBarVisibility.Disabled); values.Children.Add(list);
        list.Styles.Add(new Style(s => s.OfType<ListBoxItem>()) { Setters = { new Setter(PaddingProperty, new Thickness(6, 0)), new Setter(MinHeightProperty, 28d) } });
        var empty = new TextBlock { Text = "没有匹配的值", Margin = new(8), IsVisible = false }; values.Children.Add(empty);
        CandidateItem[] shown = []; var updatingAll = false;
        void UpdateAll()
        {
            updatingAll = true;
            try { var count = shown.Count(c => c.Selected); all.IsChecked = count == 0 ? false : count == shown.Length ? true : null; all.IsEnabled = shown.Length > 0; }
            finally { updatingAll = false; }
        }
        void Refresh()
        {
            shown = items.Where(c => c.Candidate.Label.Contains(_search.Text ?? "", StringComparison.OrdinalIgnoreCase)).ToArray();
            list.ItemsSource = shown; empty.IsVisible = shown.Length == 0; UpdateAll();
        }
        foreach (var item in items) item.PropertyChanged += (_, _) => { if (!updatingAll) UpdateAll(); };
        _search.TextChanged += (_, _) => Guard(Refresh);
        all.IsCheckedChanged += (_, _) =>
        {
            if (updatingAll) return;
            Guard(() =>
            {
                // Search never changes candidates outside the visible results.
                var select = all.IsChecked != false; updatingAll = true;
                try { foreach (var item in shown) item.Selected = select; }
                finally { updatingAll = false; }
                UpdateAll();
            });
        };
        Refresh();

        var conditions = new StackPanel { Spacing = 6 }; panel.Children.Add(conditions);
        string[] operators = ["文本等于", "文本包含", "文本开头", "文本结尾", "数字等于", "大于", "大于等于", "小于", "小于等于", "日期不早于", "日期不晚于", "等于", "不等于", "小于", "小于等于", "大于", "大于等于"];
        ComboBox Operator(string name, string label, FilterTerm? term)
        {
            var box = new ComboBox { Name = name, ItemsSource = operators, SelectedIndex = (int)(term?.Operator ?? FilterOperator.TextContains), HorizontalAlignment = HorizontalAlignment.Stretch };
            AutomationProperties.SetName(box, label); conditions.Children.Add(box); return box;
        }
        var first = Operator("FilterFirstOperator", "第一个条件", current?.First);
        var value1 = new TextBox { Name = "FilterFirstValue", Text = current?.First?.Operand ?? "", PlaceholderText = "条件值" }; conditions.Children.Add(value1);
        _firstValue = value1;
        AutomationProperties.SetName(value1, "第一个条件值");
        var combine = new ComboBox { Name = "FilterCombine", ItemsSource = new[] { "仅第一个条件", "并且", "或者" }, SelectedIndex = current?.Second is null ? 0 : current.UseOr ? 2 : 1, HorizontalAlignment = HorizontalAlignment.Stretch };
        conditions.Children.Add(combine); AutomationProperties.SetName(combine, "组合条件");
        var second = Operator("FilterSecondOperator", "第二个条件", current?.Second);
        var value2 = new TextBox { Name = "FilterSecondValue", Text = current?.Second?.Operand ?? "", PlaceholderText = "第二个条件值" }; conditions.Children.Add(value2);
        AutomationProperties.SetName(value2, "第二个条件值");
        void Combine() { second.IsEnabled = value2.IsEnabled = combine.SelectedIndex > 0; }
        combine.SelectionChanged += (_, _) => Combine(); Combine();
        conditions.Children.Add(new TextBlock { Text = "数字使用小数点；日期使用 yyyy-MM-dd。", FontSize = 12, TextWrapping = TextWrapping.Wrap });
        void Mode() { values.IsVisible = custom.IsChecked != true; conditions.IsVisible = custom.IsChecked == true; }
        custom.IsCheckedChanged += (_, _) => Mode(); Mode();
        if (candidateError is not null) { values.IsEnabled = false; custom.IsChecked = true; ShowError(candidateError); }

        void Apply()
        {
            if (custom.IsChecked != true && candidateError is not null) throw new InvalidOperationException(candidateError);
            var filter = custom.IsChecked != true ? new ColumnFilter(Values: items.Where(c => c.Selected).Select(c => c.Candidate.Key).ToImmutableHashSet()) :
                new(First: new((FilterOperator)first.SelectedIndex, value1.Text ?? ""),
                    Second: combine.SelectedIndex == 0 ? null : new((FilterOperator)second.SelectedIndex, value2.Text ?? ""), UseOr: combine.SelectedIndex == 2);
            filter.Validate();
            if (filter.Values is { Count: 0 }) throw new ArgumentException("请至少选择一个值或空白。");
            Completed?.Invoke(new("apply", filter));
        }
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, HorizontalAlignment = HorizontalAlignment.Right, Margin = new(0, 8, 0, 0) };
        buttons.Children.Add(Command("FilterApply", "确定", Apply, false)); buttons.Children.Add(Command("FilterCancel", "取消", () => Completed?.Invoke(null), false));
        var footer = new StackPanel { Spacing = 4 }; footer.Children.Add(_error); footer.Children.Add(buttons);
        var root = new DockPanel(); DockPanel.SetDock(footer, Dock.Bottom); root.Children.Add(footer);
        root.Children.Add(new ScrollViewer { Content = panel, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled }); Child = root;
        KeyboardNavigation.SetTabNavigation(this, KeyboardNavigationMode.Cycle);
        AddHandler(KeyDownEvent, (_, e) =>
        {
            if (e.Key == Key.Escape) { e.Handled = true; Completed?.Invoke(null); }
            else if (e.Key == Key.Enter && e.Source is not ComboBox && !first.IsDropDownOpen && !second.IsDropDownOpen && !combine.IsDropDownOpen)
            { e.Handled = true; Guard(Apply); }
        }, RoutingStrategies.Bubble);
    }
    private void ShowError(string message) { _error.Text = message; _error.IsVisible = true; }
    private void Guard(Action action)
    {
        try { _error.IsVisible = false; action(); }
        catch (Exception ex) { ShowError(ex.Message); }
    }
}
