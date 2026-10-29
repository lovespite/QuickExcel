using System.Collections.Immutable;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal enum FilterPopupAction { Apply, ClearColumn, ClearAll, SortAscending, SortDescending, AdvancedSort }
internal sealed record FilterPopupResult(FilterPopupAction Action, ColumnFilter? Filter = null);

/// <summary>Native popup with a private draft. Closing, Escape and search never mutate the applied filters.</summary>
internal sealed class FilterPopup
{
    internal sealed class Choice(FilterCandidate candidate) : INotifyPropertyChanged
    {
        public CellValue Key => candidate.Key;
        public string Label => candidate.Label + $"  ({candidate.Count:N0})";
        private bool _selected = candidate.Selected;
        public bool Selected { get => _selected; set { _selected = value; PropertyChanged?.Invoke(this, new(nameof(Selected))); } }
        public event PropertyChangedEventHandler? PropertyChanged;
    }

    private readonly Popup _popup;
    private readonly List<Choice> _choices;
    private readonly ICollectionView _items;
    private readonly TaskCompletionSource<FilterPopupResult?> _completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
    internal TextBox Search { get; } = new() { Margin = new(0, 6, 0, 6), ToolTip = "搜索候选值；不会改变未显示项的选择" };
    internal ListBox Choices { get; } = new() { Height = 230, BorderThickness = new(1) };
    internal Button ApplyButton { get; } = new() { Content = "确定", MinWidth = 70, Margin = new(3) };
    internal Button CancelButton { get; } = new() { Content = "取消", MinWidth = 70, Margin = new(3) };
    internal IReadOnlyList<Choice> Draft => _choices;
    private readonly CheckBox _conditions = new() { Content = "使用条件筛选", Margin = new(0, 8, 0, 5) };
    private readonly ComboBox _firstOperator = OperatorPicker();
    private readonly ComboBox _secondOperator = OperatorPicker();
    private readonly TextBox _firstValue = new() { Margin = new(4, 2, 0, 2), MinWidth = 145 };
    private readonly TextBox _secondValue = new() { Margin = new(4, 2, 0, 2), MinWidth = 145 };
    private readonly CheckBox _second = new() { Content = "第二条件", Margin = new(0, 5, 4, 5) };
    private readonly ComboBox _join = new() { ItemsSource = new[] { "并且 AND", "或者 OR" }, SelectedIndex = 0, MinWidth = 110 };
    private readonly TextBlock _error = new() { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap };

    private static readonly string[] OperatorNames = ["文本等于", "文本包含", "文本开头", "文本结尾", "数字等于", "数字大于", "数字大于等于", "数字小于", "数字小于等于", "日期不早于", "日期不晚于"];
    private static ComboBox OperatorPicker() => new() { ItemsSource = OperatorNames, SelectedIndex = 0, MinWidth = 125, Margin = new(0, 2, 0, 2) };
    private static StackPanel Line(params UIElement[] children)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal };
        foreach (var child in children) panel.Children.Add(child);
        return panel;
    }

    public FilterPopup(FrameworkElement target, Rect placement, string title, IReadOnlyList<FilterCandidate> candidates, ColumnFilter? existing)
    {
        _choices = candidates.Select(candidate => new Choice(candidate)).ToList();
        _items = CollectionViewSource.GetDefaultView(_choices);
        _items.Filter = value => ((Choice)value).Label.Contains(Search.Text, StringComparison.OrdinalIgnoreCase);
        Search.TextChanged += (_, _) => _items.Refresh();
        Choices.ItemsSource = _items;
        VirtualizingPanel.SetIsVirtualizing(Choices, true);
        VirtualizingPanel.SetVirtualizationMode(Choices, VirtualizationMode.Recycling);
        ScrollViewer.SetCanContentScroll(Choices, true);
        var template = new FrameworkElementFactory(typeof(CheckBox));
        template.SetBinding(ContentControl.ContentProperty, new Binding(nameof(Choice.Label)));
        template.SetBinding(ToggleButton.IsCheckedProperty, new Binding(nameof(Choice.Selected)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
        template.SetValue(FrameworkElement.MarginProperty, new Thickness(3));
        Choices.ItemTemplate = new DataTemplate { VisualTree = template };
        var panel = new StackPanel { Margin = new(12), Width = 355 };
        panel.Children.Add(new TextBlock { Text = title, FontWeight = FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap });
        void Command(string label, FilterPopupAction action)
        {
            var button = new Button { Content = label, HorizontalContentAlignment = HorizontalAlignment.Left, Margin = new(0, 2, 0, 0) };
            button.Click += (_, _) => Complete(new(action)); panel.Children.Add(button);
        }
        Command("升序排列（包括隐藏记录）", FilterPopupAction.SortAscending);
        Command("降序排列（包括隐藏记录）", FilterPopupAction.SortDescending);
        Command("多级排序…", FilterPopupAction.AdvancedSort);
        Command("清除此列条件", FilterPopupAction.ClearColumn);
        Command("清除全部条件", FilterPopupAction.ClearAll);
        panel.Children.Add(Search);
        var all = new Button { Content = "选中搜索结果", Margin = new(0, 0, 4, 4) };
        var none = new Button { Content = "取消搜索结果选择", Margin = new(0, 0, 0, 4) };
        all.Click += (_, _) => { foreach (Choice choice in _items) choice.Selected = true; };
        none.Click += (_, _) => { foreach (Choice choice in _items) choice.Selected = false; };
        panel.Children.Add(Line(all, none)); panel.Children.Add(Choices);
        panel.Children.Add(_conditions);
        panel.Children.Add(Line(_firstOperator, _firstValue));
        panel.Children.Add(Line(_second, _join));
        panel.Children.Add(Line(_secondOperator, _secondValue));
        panel.Children.Add(new TextBlock { Text = "日期 yyyy-MM-dd；数字小数点为 .。文本条件忽略大小写。", FontSize = 11, TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(_error);
        panel.Children.Add(Line(ApplyButton, CancelButton));
        if (existing?.First is { } first)
        {
            _conditions.IsChecked = true; _firstOperator.SelectedIndex = (int)first.Operator; _firstValue.Text = first.Operand;
            if (existing.Second is { } second) { _second.IsChecked = true; _secondOperator.SelectedIndex = (int)second.Operator; _secondValue.Text = second.Operand; }
            _join.SelectedIndex = existing.UseOr ? 1 : 0;
        }
        void EnableInputs()
        {
            var conditions = _conditions.IsChecked == true;
            Search.IsEnabled = Choices.IsEnabled = all.IsEnabled = none.IsEnabled = !conditions;
            _firstOperator.IsEnabled = _firstValue.IsEnabled = _second.IsEnabled = conditions;
            _secondOperator.IsEnabled = _secondValue.IsEnabled = _join.IsEnabled = conditions && _second.IsChecked == true;
        }
        _conditions.Checked += (_, _) => EnableInputs(); _conditions.Unchecked += (_, _) => EnableInputs();
        _second.Checked += (_, _) => EnableInputs(); _second.Unchecked += (_, _) => EnableInputs(); EnableInputs();
        var border = new Border { Background = Brushes.White, BorderBrush = Brushes.Gray, BorderThickness = new(1), Child = new ScrollViewer { Content = panel, MaxHeight = Math.Max(250, SystemParameters.WorkArea.Height - 60), VerticalScrollBarVisibility = ScrollBarVisibility.Auto } };
        _popup = new() { PlacementTarget = target, PlacementRectangle = placement, Placement = PlacementMode.Bottom,
            StaysOpen = false, AllowsTransparency = true, Child = border };
        var composing = false;
        border.AddHandler(TextCompositionManager.PreviewTextInputStartEvent, new TextCompositionEventHandler((_, _) => composing = true), true);
        border.AddHandler(TextCompositionManager.PreviewTextInputUpdateEvent, new TextCompositionEventHandler((_, _) => composing = true), true);
        border.AddHandler(TextCompositionManager.PreviewTextInputEvent, new TextCompositionEventHandler((_, _) => border.Dispatcher.BeginInvoke(new Action(() => composing = false))), true);
        border.PreviewKeyDown += (_, e) =>
        {
            if (composing || e.Key == Key.ImeProcessed) return;
            if (e.Key == Key.Escape) { e.Handled = true; Complete(null); }
            if (e.Key == Key.Enter && e.OriginalSource is TextBox) { e.Handled = true; Apply(); }
        };
        ApplyButton.Click += (_, _) => { if (!composing) Apply(); };
        CancelButton.Click += (_, _) => Complete(null);
        _popup.Closed += (_, _) => { _completion.TrySetResult(null); target.Focus(); };
    }

    private void Apply()
    {
        try
        {
            var filter = _conditions.IsChecked == true
                ? new ColumnFilter(First: new((FilterOperator)_firstOperator.SelectedIndex, _firstValue.Text),
                    Second: _second.IsChecked == true ? new((FilterOperator)_secondOperator.SelectedIndex, _secondValue.Text) : null, UseOr: _join.SelectedIndex == 1)
                : new ColumnFilter(_choices.Where(choice => choice.Selected).Select(choice => choice.Key).ToImmutableHashSet());
            filter.Validate(); Complete(new(FilterPopupAction.Apply, filter));
        }
        catch (ArgumentException ex) { _error.Text = ex.Message; }
    }
    private void Complete(FilterPopupResult? result) { _completion.TrySetResult(result); _popup.IsOpen = false; }
    public Task<FilterPopupResult?> ShowAsync() { _popup.IsOpen = true; Search.Focus(); return _completion.Task; }
    public void Close() => Complete(null);
}
