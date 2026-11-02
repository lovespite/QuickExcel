using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using QuickExcel.Core;

namespace QuickExcel.Desktop;

internal sealed class SeriesDialog : Window
{
    internal ComboBox KindPicker { get; } = new() { ItemsSource = new[] { "数字（等差）", "日期：日", "日期：月", "日期：年", "文本：编号（XX1、XX2…）" }, SelectedIndex = 0, Margin = new(0, 8, 0, 8) };
    internal TextBox StepInput { get; } = new() { Text = "1", Margin = new(0, 6, 0, 12) };
    internal Button ApplyButton { get; } = new() { Content = "填充", IsDefault = true, MinWidth = 80 };
    public CellSeriesKind Kind => (CellSeriesKind)KindPicker.SelectedIndex;
    public double Step { get; private set; }
    public SeriesDialog(Window owner)
    {
        Owner = owner; Title = "数字与日期序列"; SizeToContent = SizeToContent.WidthAndHeight; ResizeMode = ResizeMode.NoResize;
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var panel = new StackPanel { Margin = new(20), Width = 360 };
        panel.Children.Add(new TextBlock { Text = "选择单行或单列，以首格数字/日期/文本编号为起点，填满当前选区。首格格式随序列复制；文本编号保留前缀和前导零（XX1→XX2、A001→A002）。", TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(KindPicker); panel.Children.Add(new TextBlock { Text = "步长（日期/文本编号使用整数，可为负数）" }); panel.Children.Add(StepInput);
        var error = new TextBlock { Foreground = System.Windows.Media.Brushes.Firebrick, TextWrapping = TextWrapping.Wrap }; panel.Children.Add(error);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        buttons.Children.Add(ApplyButton); buttons.Children.Add(new Button { Content = "取消", IsCancel = true, MinWidth = 80, Margin = new(8, 0, 0, 0) }); panel.Children.Add(buttons);
        ApplyButton.Click += (_, _) =>
        {
            if (KindPicker.SelectedIndex < 0 || !double.TryParse(StepInput.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var step) || !double.IsFinite(step) ||
                Kind != CellSeriesKind.Numbers && (step != Math.Truncate(step) || Math.Abs(step) > 100_000))
            { error.Text = "请输入有限数字；日期或文本编号步长须为 -100,000 至 100,000 的整数。"; return; }
            Step = step; DialogResult = true;
        };
        Content = panel; Loaded += (_, _) => { StepInput.Focus(); StepInput.SelectAll(); };
    }
}
