using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

internal sealed partial class NumberFormatDialog
{
    private static Func<CellStylePatch, CellStylePatch> AdditionalFontFields(StackPanel panel, StyleDefinition style)
    {
        panel.Children.Add(new TextBlock { Text = UiText.Get("字体") });
        var name = new TextBox { Name = "FormatFontName", Text = style.FontName ?? "Microsoft YaHei UI", MaxLength = 256 };
        AutomationProperties.SetName(name, UiText.Get("字体")); panel.Children.Add(name);
        panel.Children.Add(new TextBlock { Text = UiText.Get("字体颜色") });
        var color = style.FontColor;
        panel.Children.Add(ColorSelector("FormatFontColor", color, UiText.Get("自动字体色"), value => color = value));
        panel.Children.Add(new TextBlock { Text = UiText.Get("下划线") });
        string[] values = ["none", "single", "double", "singleAccounting", "doubleAccounting"];
        var underline = new ComboBox { Name = "FormatUnderline", ItemsSource = new[] { UiText.Get("无"), UiText.Get("单下划线"), UiText.Get("双下划线"), UiText.Get("会计单下划线"), UiText.Get("会计双下划线") }, SelectedIndex = Math.Max(0, Array.IndexOf(values, style.Underline ?? "none")) };
        AutomationProperties.SetName(underline, UiText.Get("下划线")); panel.Children.Add(underline);
        var strike = new CheckBox { Name = "FormatStrike", Content = UiText.Get("删除线"), IsChecked = style.Strike }; panel.Children.Add(strike);
        return patch => patch with
        {
            FontName = name.Text == (style.FontName ?? "Microsoft YaHei UI") ? null : string.IsNullOrWhiteSpace(name.Text) ? throw new ArgumentException(UiText.Get("请输入字体名称。")) : name.Text.Trim(),
            FontColor = color == style.FontColor ? null : color, ClearFontColor = color != style.FontColor && color is null,
            Underline = values[underline.SelectedIndex] == (style.Underline ?? "none") ? null : values[underline.SelectedIndex],
            Strike = (strike.IsChecked == true) == style.Strike ? null : strike.IsChecked == true
        };
    }

    private void AddFillAndProtection(TabControl tabs, StyleDefinition style)
    {
        var fill = style.FillColor;
        var fillPanel = new StackPanel { Spacing = 14, Margin = new(8, 20) };
        fillPanel.Children.Add(new TextBlock { Text = UiText.Get("填充颜色") });
        fillPanel.Children.Add(ColorSelector("FormatFillColor", fill, UiText.Get("无填充"), value => fill = value));
        tabs.Items.Add(new TabItem { Header = UiText.Get("填充"), Content = fillPanel });
        var protection = new StackPanel { Spacing = 14, Margin = new(8, 20) };
        var locked = new CheckBox { Name = "FormatLocked", Content = UiText.Get("锁定"), IsChecked = style.Locked };
        var hidden = new CheckBox { Name = "FormatFormulaHidden", Content = UiText.Get("隐藏公式"), IsChecked = style.FormulaHidden };
        protection.Children.Add(locked); protection.Children.Add(hidden);
        protection.Children.Add(new TextBlock { Text = UiText.Get("保护选项在启用工作表保护后生效。"), TextWrapping = TextWrapping.Wrap, Foreground = Brushes.Gray });
        tabs.Items.Add(new TabItem { Header = UiText.Get("保护"), Content = protection });
        _fillAndProtection = patch => patch with
        {
            FillColor = fill == style.FillColor ? null : fill, ClearFillColor = fill != style.FillColor && fill is null,
            Locked = (locked.IsChecked == true) == style.Locked ? null : locked.IsChecked == true,
            FormulaHidden = (hidden.IsChecked == true) == style.FormulaHidden ? null : hidden.IsChecked == true
        };
    }

    private static DropDownButton ColorSelector(string name, string? initial, string reset, Action<string?> change)
    {
        var menu = new MenuFlyout();
        var button = new DropDownButton { Name = name, MinWidth = 200, HorizontalAlignment = HorizontalAlignment.Left, Flyout = menu };
        void Choose(string? color)
        {
            initial = color;
            var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 };
            row.Children.Add(ColorPalette.Swatch(color)); row.Children.Add(new TextBlock { Text = color is null ? reset : "#" + color[2..] });
            button.Content = row; change(color);
        }
        var items = new List<Control>();
        void Add(string title, string? color)
        {
            var item = new MenuItem { Header = title, Icon = ColorPalette.Swatch(color), Tag = color };
            item.Click += (_, _) => Choose(color); items.Add(item);
        }
        Add(reset, null); items.Add(new Separator());
        foreach (var (title, color) in ColorPalette.Common) Add(title, color);
        var custom = new MenuItem { Header = UiText.Get("自定义颜色…") };
        custom.Click += async (_, _) =>
        {
            if (TopLevel.GetTopLevel(button) is Window owner && await new CustomColorDialog(initial, UiText.Get("自定义颜色…")).ShowAsync(owner) is { } color) Choose(color);
        };
        items.Add(new Separator()); items.Add(custom); menu.ItemsSource = items;
        AutomationProperties.SetName(button, name == "FormatFillColor" ? UiText.Get("填充颜色") : UiText.Get("字体颜色"));
        Choose(initial); return button;
    }
}
