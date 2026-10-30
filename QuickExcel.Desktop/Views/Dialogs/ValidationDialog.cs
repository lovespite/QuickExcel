using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal sealed class ValidationDialog : Window
{
    internal ComboBox KindPicker { get; } = new() { ItemsSource = new[] { "任何值（仅提示）", "整数", "小数", "列表", "日期", "文本长度", "自定义公式" }, SelectedIndex = 1 };
    internal ComboBox OperatorPicker { get; } = new() { ItemsSource = new[] { "介于", "不介于", "等于", "不等于", "小于", "小于或等于", "大于", "大于或等于" }, SelectedIndex = 0 };
    internal TextBox First { get; } = new() { Text = "1" };
    internal TextBox Second { get; } = new() { Text = "100" };
    internal CheckBox Blank { get; } = new() { Content = "忽略空白", IsChecked = true };
    internal CheckBox DropDown { get; } = new() { Content = "显示列表下拉选择", IsChecked = true };
    internal CheckBox Input { get; } = new() { Content = "选中单元格时显示录入提示", IsChecked = true };
    internal TextBox InputTitle { get; } = new() { MaxLength = 32 };
    internal TextBox InputMessage { get; } = new() { MaxLength = 255 };
    internal CheckBox Error { get; } = new() { Content = "无效输入时显示错误提示", IsChecked = true };
    internal ComboBox AlertPicker { get; } = new() { ItemsSource = new[] { "停止（不能继续）", "警告（可确认继续）", "信息（可确认继续）" }, SelectedIndex = 0 };
    internal TextBox ErrorTitle { get; } = new() { MaxLength = 32 };
    internal TextBox ErrorMessage { get; } = new() { MaxLength = 255 };
    internal Button ApplyButton { get; } = new() { Content = "应用", IsDefault = true, MinWidth = 90 };
    public WorksheetValidation? Result { get; private set; }
    public ValidationDialog(Window owner, IReadOnlyList<CellRectangle> ranges, WorksheetValidation? existing = null)
    {
        Owner = owner; Title = "数据验证"; Width = 560; Height = 680; MinWidth = 470; MinHeight = 480;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var dock = new DockPanel { Margin = new(20) }; var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new(0, 12, 0, 0) };
        buttons.Children.Add(ApplyButton); buttons.Children.Add(new Button { Content = "取消", IsCancel = true, MinWidth = 90, Margin = new(8, 0, 0, 0) });
        DockPanel.SetDock(buttons, Dock.Bottom); dock.Children.Add(buttons);
        var panel = new StackPanel(); dock.Children.Add(new ScrollViewer { Content = panel, VerticalScrollBarVisibility = ScrollBarVisibility.Auto });
        panel.Children.Add(new TextBlock { Text = "应用于：" + string.Join(" ", ranges), TextWrapping = TextWrapping.Wrap, FontWeight = FontWeights.SemiBold });
        void Field(string label, Control control) { panel.Children.Add(new TextBlock { Text = label, Margin = new(0, 8, 0, 3) }); panel.Children.Add(control); }
        Field("允许", KindPicker); Field("比较", OperatorPicker); Field("条件 1 / 列表来源", First); Field("条件 2（介于 / 不介于）", Second);
        panel.Children.Add(new TextBlock { Text = "列表：用逗号分隔项目，或输入 =$D$1:$D$10 / =定义名称。\n日期：yyyy-MM-dd 或 =DATE(2026,10,31)。自定义公式：例如 =AND(ISNUMBER(A1),A1>0)，以区域首格为基准。",
            TextWrapping = TextWrapping.Wrap, Margin = new(0, 6, 0, 8) });
        panel.Children.Add(Blank); panel.Children.Add(DropDown); panel.Children.Add(Input);
        Field("录入标题 / 提示", InputTitle); panel.Children.Add(InputMessage);
        panel.Children.Add(Error); Field("错误样式", AlertPicker); Field("错误标题 / 提示", ErrorTitle); panel.Children.Add(ErrorMessage);
        var error = new TextBlock { Foreground = System.Windows.Media.Brushes.Firebrick, TextWrapping = TextWrapping.Wrap, Margin = new(0, 10, 0, 0) }; panel.Children.Add(error);
        if (existing is { } v)
        {
            KindPicker.SelectedIndex = (int)v.Kind; OperatorPicker.SelectedIndex = (int)v.Comparison;
            First.Text = v.Formula1 ?? ""; Second.Text = v.Formula2 ?? ""; Blank.IsChecked = v.AllowBlank; DropDown.IsChecked = v.ShowDropDown;
            Input.IsChecked = v.ShowInput; Error.IsChecked = v.ShowError; AlertPicker.SelectedIndex = (int)v.Alert;
            InputTitle.Text = v.InputTitle ?? ""; InputMessage.Text = v.InputMessage ?? ""; ErrorTitle.Text = v.ErrorTitle ?? ""; ErrorMessage.Text = v.ErrorMessage ?? "";
        }
        ApplyButton.Click += (_, _) =>
        {
            try
            {
                var kind = (ValidationKind)KindPicker.SelectedIndex;
                string Formula(string value)
                {
                    if (kind == ValidationKind.Date && DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                        return $"=DATE({date.Year},{date.Month},{date.Day})";
                    return value;
                }
                var first = Formula(First.Text.Trim());
                if (kind == ValidationKind.List && !first.StartsWith('=') && !first.StartsWith('"')) first = "=\"" + first.Replace("\"", "\"\"") + "\"";
                Result = new(ranges, kind, kind == ValidationKind.None ? null : first,
                    kind is ValidationKind.None or ValidationKind.List or ValidationKind.Custom || OperatorPicker.SelectedIndex > 1 ? null : Formula(Second.Text.Trim()),
                    (ValidationOperator)OperatorPicker.SelectedIndex, Blank.IsChecked == true, DropDown.IsChecked == true, Input.IsChecked == true, Error.IsChecked == true,
                    (ValidationAlert)AlertPicker.SelectedIndex, Empty(InputTitle.Text), Empty(InputMessage.Text), Empty(ErrorTitle.Text), Empty(ErrorMessage.Text));
                DialogResult = true;
            }
            catch (Exception ex) when (ex is ArgumentException or FormatException or NotSupportedException) { error.Text = ex.Message; }
        };
        Content = dock;
    }
    private static string? Empty(string value) => value.Length == 0 ? null : value;
}
