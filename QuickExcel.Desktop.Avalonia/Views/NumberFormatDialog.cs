using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class NumberFormatDialog : Window
{
    private readonly ListBox _categories = new() { Name = "FormatCategories", ItemsSource = NumberFormatOptions.Categories };
    private readonly ListBox _types = new() { Name = "FormatTypes", MinHeight = 100 };
    private readonly TextBox _code = new() { Name = "FormatCode", MaxLength = 1024 };
    private readonly TextBlock _preview = new() { Name = "FormatPreview", FontSize = 20, TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock _error = new() { Name = "FormatError", Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock _hint = new() { TextWrapping = TextWrapping.Wrap, Foreground = Brushes.Gray };
    private readonly NumericUpDown _decimals = new() { Name = "FormatDecimals", Minimum = 0, Maximum = 30, Increment = 1, FormatString = "0", Value = 2, Width = 110 };
    private readonly ComboBox _symbol = new() { Name = "FormatSymbol", ItemsSource = NumberFormatOptions.Symbols.Select(UiText.Get).ToArray(), SelectedIndex = 0, MinWidth = 130 };
    private readonly ComboBox _negative = new() { Name = "FormatNegative", ItemsSource = new[] { "-1234.10", UiText.Get("红色 -1234.10"), "(1234.10)", UiText.Get("红色 (1234.10)") }, SelectedIndex = 0, MinWidth = 200 };
    private readonly CheckBox _grouping = new() { Name = "FormatGrouping", Content = UiText.Get("使用千位分隔符（,）") };
    private readonly StackPanel _numeric = new() { Spacing = 8 };
    private readonly Button _ok = new() { Name = "FormatConfirm", Content = UiText.Get("确定"), IsDefault = true, MinWidth = 80 };
    private readonly StyleDefinition _original;
    private readonly CellValue _sample;
    private readonly ExcelDateSystem _dateSystem;
    private readonly string[] _customCodes;
    private Func<CellStylePatch, CellStylePatch>? _appearance;
    private BorderEditor? _borders;
    public CellBorderPatch? BorderPatch => _borders?.Patch;
    private bool _loading;
    private bool _numberChanged;
    private NumberFormatCategory Category => (NumberFormatCategory)_categories.SelectedIndex;

    public NumberFormatDialog(StyleDefinition style, CellValue sample, ExcelDateSystem dateSystem, IEnumerable<StyleDefinition> styles, bool appearance = false, BorderStroke?[]? borders = null, bool borderTab = false)
    {
        _original = style; _sample = sample.Kind == CellValueKind.Blank ? CellValue.FromNumber(1234.56) : sample; _dateSystem = dateSystem;
        _customCodes = new[] { NumberFormatOptions.Code(style) }.Concat(styles.Select(NumberFormatOptions.Code)).Concat(NumberFormatOptions.CustomCodes).Where(c => c.Length > 0).Distinct().ToArray();
        Title = appearance ? UiText.Get("设置单元格格式") : UiText.Get("数字格式"); Width = 760; Height = 620; MinWidth = 650; MinHeight = appearance ? 620 : 540;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var root = new global::Avalonia.Controls.Grid { RowDefinitions = new("*,Auto,Auto"), Margin = new(20) };
        var body = new global::Avalonia.Controls.Grid { ColumnDefinitions = new("145,20,*") };
        var categories = new global::Avalonia.Controls.Grid { RowDefinitions = new("Auto,*") };
        categories.Children.Add(new TextBlock { Text = UiText.Get("分类"), Margin = new(0, 0, 0, 8) });
        global::Avalonia.Controls.Grid.SetRow(_categories, 1); categories.Children.Add(_categories); body.Children.Add(categories);
        var detail = new global::Avalonia.Controls.Grid { RowDefinitions = new("Auto,Auto,Auto,Auto,*,Auto") };
        global::Avalonia.Controls.Grid.SetColumn(detail, 2); body.Children.Add(detail);
        detail.Children.Add(new TextBlock { Text = sample.Kind == CellValueKind.Blank ? UiText.Get("示例（空单元格使用 1234.56）") : UiText.Get("示例"), Margin = new(0, 0, 0, 6) });
        var previewBorder = new Border { Child = _preview, Padding = new(12), MinHeight = 60, BorderThickness = new(1), BorderBrush = Brushes.LightGray, Margin = new(0, 0, 0, 12) };
        global::Avalonia.Controls.Grid.SetRow(previewBorder, 1); detail.Children.Add(previewBorder);
        var decimalsRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };
        decimalsRow.Children.Add(new TextBlock { Text = UiText.Get("小数位数"), VerticalAlignment = VerticalAlignment.Center }); decimalsRow.Children.Add(_decimals); decimalsRow.Children.Add(_grouping);
        _numeric.Children.Add(decimalsRow);
        var currencyRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };
        currencyRow.Children.Add(new TextBlock { Text = UiText.Get("货币符号"), VerticalAlignment = VerticalAlignment.Center }); currencyRow.Children.Add(_symbol); _numeric.Children.Add(currencyRow);
        var negativeRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };
        negativeRow.Children.Add(new TextBlock { Text = UiText.Get("负数显示"), VerticalAlignment = VerticalAlignment.Center }); negativeRow.Children.Add(_negative); _numeric.Children.Add(negativeRow);
        _numeric.Margin = new(0, 0, 0, 10); global::Avalonia.Controls.Grid.SetRow(_numeric, 2); detail.Children.Add(_numeric);
        var codePanel = new StackPanel { Spacing = 6, Margin = new(0, 0, 0, 8) };
        codePanel.Children.Add(new TextBlock { Text = UiText.Get("类型 / 格式代码") }); codePanel.Children.Add(_code);
        global::Avalonia.Controls.Grid.SetRow(codePanel, 3); detail.Children.Add(codePanel);
        global::Avalonia.Controls.Grid.SetRow(_types, 4); detail.Children.Add(_types);
        _hint.Margin = new(0, 8, 0, 0); global::Avalonia.Controls.Grid.SetRow(_hint, 5); detail.Children.Add(_hint);
        if (appearance)
        {
            var tabs = new TabControl(); tabs.Items.Add(new TabItem { Header = UiText.Get("数字"), Content = body });
            tabs.Items.Add(new TabItem { Header = UiText.Get("对齐与字体"), Content = Appearance(style) });
            _borders = new BorderEditor(borders ?? new BorderStroke?[8]);
            tabs.Name = "CellFormatTabs"; tabs.Items.Add(new TabItem { Header = UiText.Get("边框"), Content = _borders });
            if (borderTab) tabs.SelectedIndex = 2; root.Children.Add(tabs);
        }
        else root.Children.Add(body);
        _error.Margin = new(0, 10, 0, 10); global::Avalonia.Controls.Grid.SetRow(_error, 1); root.Children.Add(_error);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, HorizontalAlignment = HorizontalAlignment.Right };
        var cancel = new Button { Name = "FormatCancel", Content = UiText.Get("取消"), IsCancel = true, MinWidth = 80 };
        cancel.Click += (_, _) => Close(null); _ok.Click += (_, _) => Accept(); buttons.Children.Add(_ok); buttons.Children.Add(cancel);
        global::Avalonia.Controls.Grid.SetRow(buttons, 2); root.Children.Add(buttons); Content = root;
        foreach (var (control, name) in new (Control, string)[] { (_categories, UiText.Get("数字格式分类")), (_types, UiText.Get("格式类型")), (_code, UiText.Get("格式代码")), (_preview, UiText.Get("格式示例")), (_decimals, UiText.Get("小数位数")), (_symbol, UiText.Get("货币符号")), (_negative, UiText.Get("负数显示")) })
            AutomationProperties.SetName(control, name);
        _categories.SelectionChanged += (_, _) => { if (!_loading) SelectCategory(); };
        _types.SelectionChanged += (_, _) => { if (!_loading && _types.SelectedItem is NumberFormatPreset preset) SetCode(preset.Code); };
        _decimals.ValueChanged += (_, _) => NumericChanged(); _symbol.SelectionChanged += (_, _) => NumericChanged(); _negative.SelectionChanged += (_, _) => NumericChanged();
        _grouping.IsCheckedChanged += (_, _) => NumericChanged();
        _code.TextChanged += (_, _) =>
        {
            if (_loading) return;
            _loading = true; _types.SelectedItem = _types.Items.OfType<NumberFormatPreset>().FirstOrDefault(p => p.Code == _code.Text); _loading = false;
            _numberChanged = true; Preview();
        };
        var initial = NumberFormatOptions.Code(style);
        _loading = true; _categories.SelectedIndex = (int)NumberFormatOptions.Category(initial); _loading = false;
        SelectCategory(initial); _numberChanged = false; Preview();
    }

    private void SelectCategory(string? initial = null)
    {
        if (_categories.SelectedIndex < 0) return;
        _loading = true;
        var numeric = Category is NumberFormatCategory.Number or NumberFormatCategory.Currency or NumberFormatCategory.Accounting or NumberFormatCategory.Percentage or NumberFormatCategory.Scientific;
        _numeric.IsVisible = numeric;
        _numeric.Children[1].IsVisible = Category is NumberFormatCategory.Currency or NumberFormatCategory.Accounting;
        _numeric.Children[2].IsVisible = Category is NumberFormatCategory.Number or NumberFormatCategory.Currency;
        _grouping.IsVisible = Category == NumberFormatCategory.Number;
        _code.IsReadOnly = Category != NumberFormatCategory.Custom;
        _types.IsVisible = !numeric;
        var choices = Category == NumberFormatCategory.Custom
            ? new[] { _code.Text ?? NumberFormatOptions.Code(_original) }.Concat(_customCodes).Distinct().Select(c => new NumberFormatPreset(c, c, Category, "")).ToArray() : NumberFormatOptions.Types(Category);
        _types.ItemsSource = choices;
        if (initial is not null && numeric && NumberFormatOptions.ReadNumeric(initial) is { } settings)
        {
            _decimals.Value = settings.Decimals; _grouping.IsChecked = settings.Grouping; _negative.SelectedIndex = settings.Negative;
            _symbol.SelectedItem = settings.Symbol == NumberFormatOptions.NoSymbol && Category is not (NumberFormatCategory.Currency or NumberFormatCategory.Accounting) ? "¥" : UiText.Get(settings.Symbol);
        }
        _code.Text = initial ?? (numeric ? NumericCode() : choices.FirstOrDefault()?.Code ?? "General");
        _types.SelectedItem = choices.FirstOrDefault(c => c.Code == _code.Text);
        _hint.Text = Category switch
        {
            NumberFormatCategory.Custom => UiText.Get("使用 0、#、?、日期时间代码；分号分隔正数、负数、零和文本。支持 [红色]、[Red] 和 G/通用格式。"),
            NumberFormatCategory.Text => UiText.Get("文本格式会保留以后输入的前导零。已有数值和公式不会转换为文本。"),
            NumberFormatCategory.Accounting => UiText.Get("对齐货币符号与数值；零值显示为短横线。"),
            NumberFormatCategory.Fraction => UiText.Get("将数值显示为最接近的分数，保留原始精度。"),
            _ => UiText.Get("格式只改变显示方式，不改变单元格中的值。")
        };
        _loading = false; _numberChanged = initial is null; Preview();
    }

    private string NumericCode() => NumberFormatOptions.Numeric(Category, (int)(_decimals.Value ?? 0), Category is NumberFormatCategory.Currency or NumberFormatCategory.Accounting || Category == NumberFormatCategory.Number && _grouping.IsChecked == true, NumberFormatOptions.Symbols[Math.Max(0, _symbol.SelectedIndex)], Math.Max(0, _negative.SelectedIndex));
    private void NumericChanged() { if (!_loading && _numeric.IsVisible) SetCode(NumericCode()); }
    private void SetCode(string code) { _code.Text = code; _numberChanged = true; Preview(); }
    private void Preview()
    {
        var valid = CellDisplayFormatter.TryNormalizeNumberFormat(_code.Text ?? "", out var code);
        _ok.IsEnabled = valid || !_numberChanged;
        _error.Text = valid ? "" : UiText.Get("此格式代码无效或暂不支持预览，请选择其他类型或修改代码。");
        var display = CellDisplayFormatter.FormatDisplay(_sample, valid && _numberChanged ? new() { NumberFormatCode = code } : _original, _dateSystem);
        _preview.Text = display.Text; _preview.Foreground = display.Color is { } color ? new SolidColorBrush(Color.Parse("#" + color)) : Brushes.Black;
    }

    private void Accept()
    {
        if (!_ok.IsEnabled) return;
        try
        {
            var patch = new CellStylePatch();
            if (_numberChanged)
            {
                if (!CellDisplayFormatter.TryNormalizeNumberFormat(_code.Text ?? "", out var code)) return;
                patch = patch with { NumberFormatId = 0, NumberFormatCode = code.Equals("General", StringComparison.OrdinalIgnoreCase) ? null : code };
            }
            Close(_appearance?.Invoke(patch) ?? patch);
        }
        catch (ArgumentException ex) { _error.Text = ex.Message; }
    }

    private Control Appearance(StyleDefinition style)
    {
        var panel = new StackPanel { Spacing = 10, Margin = new(0, 10) };
        ComboBox Choice(string label, string[] choices, int selected)
        {
            panel.Children.Add(new TextBlock { Text = label });
            var box = new ComboBox { ItemsSource = choices, SelectedIndex = selected, HorizontalAlignment = HorizontalAlignment.Stretch };
            AutomationProperties.SetName(box, label); panel.Children.Add(box); return box;
        }
        CheckBox Check(string label, bool value) { var box = new CheckBox { Content = label, IsChecked = value }; panel.Children.Add(box); return box; }
        var horizontal = Choice(UiText.Get("水平对齐"), [UiText.Get("常规"), UiText.Get("左对齐"), UiText.Get("居中"), UiText.Get("右对齐")], Math.Clamp(style.Horizontal, 0, 3));
        var vertical = Choice(UiText.Get("垂直对齐"), [UiText.Get("顶端"), UiText.Get("居中"), UiText.Get("底端"), UiText.Get("两端对齐")], Math.Clamp(style.Vertical, 0, 3));
        var wrap = Check(UiText.Get("自动换行"), style.WrapText); var bold = Check(UiText.Get("加粗"), style.Bold); var italic = Check(UiText.Get("倾斜"), style.Italic);
        panel.Children.Add(new TextBlock { Text = UiText.Get("字号") });
        var size = new NumericUpDown { Minimum = .1m, Maximum = 409, Value = (decimal)(style.FontSize ?? 11), Width = 120 }; AutomationProperties.SetName(size, UiText.Get("字号")); panel.Children.Add(size);
        _appearance = patch => patch with
        {
            Horizontal = horizontal.SelectedIndex == Math.Clamp(style.Horizontal, 0, 3) ? null : horizontal.SelectedIndex,
            Vertical = vertical.SelectedIndex == Math.Clamp(style.Vertical, 0, 3) ? null : vertical.SelectedIndex,
            WrapText = (wrap.IsChecked == true) == style.WrapText ? null : wrap.IsChecked == true,
            Bold = (bold.IsChecked == true) == style.Bold ? null : bold.IsChecked == true,
            Italic = (italic.IsChecked == true) == style.Italic ? null : italic.IsChecked == true,
            FontSize = size.Value is not { } value ? throw new ArgumentException(UiText.Get("请输入字号。")) : (double)value == (style.FontSize ?? 11) ? null : (double)value
        };
        return new ScrollViewer { Content = panel };
    }

    public Task<CellStylePatch?> ShowAsync(MainWindow owner) => owner.ShowOwnedAsync<CellStylePatch?>(this);
}
