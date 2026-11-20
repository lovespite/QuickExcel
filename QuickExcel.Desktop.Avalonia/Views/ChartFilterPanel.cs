using System.ComponentModel;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class ChartFilterPanel : Border
{
    internal sealed class Item(int index, string label, bool? selected, string? color = null) : INotifyPropertyChanged
    {
        public int Index { get; } = index;
        public string Label { get; } = label;
        public string? Color { get; } = color;
        private bool? _selected = selected;
        public bool? Selected { get => _selected; set { if (_selected == value) return; _selected = value; PropertyChanged?.Invoke(this, new(nameof(Selected))); } }
        public event PropertyChangedEventHandler? PropertyChanged;
    }
    public event Action<WorksheetChart?>? Completed;
    public event Action? SelectDataRequested;
    private readonly TextBox _search = new() { Name = "ChartFilterSearch", PlaceholderText = UiText.Get("搜索类别"), Margin = new(0, 6) };
    internal Item[] SeriesItems { get; }
    internal Item[] CategoryItems { get; }
    internal Button ApplyButton { get; }
    public void FocusSearch() => _search.Focus();
    public ChartFilterPanel(WorkbookSnapshot book, WorksheetChart chart, bool canEdit)
    {
        Width = 320; Height = 520; Padding = new(14); Background = Brushes.White; BorderBrush = Brushes.LightGray; BorderThickness = new(1); CornerRadius = new(4);
        AutomationProperties.SetName(this, UiText.Get("图表筛选器"));
        var data = ChartData.Read(book, chart with { VisibleOnly = false }, includeFiltered: true);
        SeriesItems = data.Select((s, i) => new Item(i, string.IsNullOrEmpty(s.Name) ? UiText.Format($"系列 {i + 1}") : s.Name, !chart.Series[i].IsFiltered, s.Color)).ToArray();
        CategoryItems = Enumerable.Range(0, data.Max(s => s.Categories.Length)).Select(i =>
        {
            var included = chart.Series.Where(s => i < WorksheetChart.Count(WorksheetChart.ParseReference(s.Values))).Select(s => !s.HiddenPoints.Contains(i)).ToArray();
            var labels = data.Where(s => i < s.Categories.Length).Select(s => string.IsNullOrEmpty(s.Categories[i]) ? UiText.Get("(空白)") : s.Categories[i]).Distinct().ToArray();
            return new Item(i, string.Join(" / ", labels), included.All(v => v) ? true : included.All(v => !v) ? false : null);
        }).ToArray();
        var layout = new global::Avalonia.Controls.Grid { RowDefinitions = new("Auto,Auto,100,Auto,Auto,*,Auto,Auto") };
        void Add(Control control, int row) { global::Avalonia.Controls.Grid.SetRow(control, row); layout.Children.Add(control); }
        Add(new TextBlock { Text = UiText.Get("图表筛选器"), FontSize = 16, FontWeight = FontWeight.SemiBold, Margin = new(0, 0, 0, 8) }, 0);
        ListBox List(Item[] items, bool categories)
        {
            var list = new ListBox { ItemsSource = items, Background = Brushes.White, BorderThickness = new(0),
                ItemTemplate = new FuncDataTemplate<Item>((item, _) =>
                {
                    if (item is null) return null;
                    var label = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 7 };
                    if (item.Color is { } color) label.Children.Add(new Border { Width = 9, Height = 9, Background = new SolidColorBrush(global::Avalonia.Media.Color.Parse("#" + color)), VerticalAlignment = VerticalAlignment.Center });
                    label.Children.Add(new TextBlock { Text = item.Label, TextTrimming = TextTrimming.CharacterEllipsis, MaxWidth = 225, VerticalAlignment = VerticalAlignment.Center });
                    var box = new CheckBox { Content = label, MinHeight = 28, IsThreeState = false, IsEnabled = canEdit };
                    AutomationProperties.SetName(box, item.Label); ToolTip.SetTip(box, categories ? UiText.Format($"第 {item.Index + 1} 项 · {item.Label}") : item.Label);
                    box.Bind(ToggleButton.IsCheckedProperty,
                        CompiledBinding.Create((Item value) => value.Selected, source: item, mode: BindingMode.TwoWay)); return box;
                }) };
            list.Styles.Add(new Style(s => s.OfType<ListBoxItem>()) { Setters = { new Setter(PaddingProperty, new Thickness(5, 0)), new Setter(MinHeightProperty, 28d) } });
            ScrollViewer.SetHorizontalScrollBarVisibility(list, ScrollBarVisibility.Disabled); return list;
        }
        var seriesList = List(SeriesItems, false); var categoryList = List(CategoryItems, true);
        var seriesAll = new CheckBox { Content = UiText.Get("系列（全选）"), Name = "ChartSeriesAll", IsEnabled = canEdit };
        var categoryAll = new CheckBox { Content = UiText.Get("类别（全选）"), Name = "ChartCategoriesAll", IsEnabled = canEdit };
        Add(seriesAll, 1); Add(seriesList, 2); Add(_search, 3); Add(categoryAll, 4); Add(categoryList, 5);
        Item[] shown = CategoryItems; var updating = false;
        void UpdateAll()
        {
            updating = true;
            try
            {
                static bool? State(Item[] items) => items.Length == 0 || items.All(i => i.Selected == false) ? false : items.All(i => i.Selected == true) ? true : null;
                seriesAll.IsChecked = State(SeriesItems); categoryAll.IsChecked = State(shown); categoryAll.IsEnabled = canEdit && shown.Length > 0;
            }
            finally { updating = false; }
        }
        foreach (var item in SeriesItems.Concat(CategoryItems)) item.PropertyChanged += (_, _) => { if (!updating) UpdateAll(); };
        void SetAll(CheckBox box, Func<Item[]> items)
        {
            box.IsCheckedChanged += (_, _) => { if (updating) return; updating = true; try { foreach (var item in items()) item.Selected = box.IsChecked != false; } finally { updating = false; } UpdateAll(); };
        }
        SetAll(seriesAll, () => SeriesItems); SetAll(categoryAll, () => shown);
        _search.TextChanged += (_, _) => { shown = CategoryItems.Where(i => i.Label.Contains(_search.Text ?? "", StringComparison.OrdinalIgnoreCase)).ToArray(); categoryList.ItemsSource = shown; UpdateAll(); };
        UpdateAll();
        Add(new TextBlock { Text = canEdit ? UiText.Get("筛选随文件保存，仅影响此图表。") : UiText.Get("当前图表只读，无法修改筛选。"), Foreground = Brushes.Gray, FontSize = 12, TextWrapping = TextWrapping.Wrap, Margin = new(0, 8) }, 6);
        var footer = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        ApplyButton = new Button { Name = "ChartFilterApply", Content = UiText.Get("应用"), IsEnabled = canEdit };
        ApplyButton.Click += (_, _) => Completed?.Invoke(chart with { Series = [.. chart.Series.Select((s, i) => s with { IsFiltered = SeriesItems[i].Selected == false,
            HiddenPoints = [.. Enumerable.Range(0, WorksheetChart.Count(WorksheetChart.ParseReference(s.Values))).Where(p => CategoryItems[p].Selected == false || CategoryItems[p].Selected is null && s.HiddenPoints.Contains(p))] })] });
        var reset = new Button { Name = "ChartFilterReset", Content = UiText.Get("全部显示"), IsEnabled = canEdit };
        reset.Click += (_, _) => { updating = true; try { foreach (var item in SeriesItems.Concat(CategoryItems)) item.Selected = true; } finally { updating = false; } UpdateAll(); };
        var dataButton = new Button { Content = UiText.Get("选择数据…"), IsEnabled = canEdit }; dataButton.Click += (_, _) => SelectDataRequested?.Invoke();
        footer.Children.Add(ApplyButton); footer.Children.Add(reset); footer.Children.Add(dataButton); Add(footer, 7); Child = layout;
        AddHandler(KeyDownEvent, (_, e) => { if (e.Key == Key.Escape) { e.Handled = true; Completed?.Invoke(null); } });
    }
}
