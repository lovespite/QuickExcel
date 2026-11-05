using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class PasteOptionsPanel : Border
{
    internal Dictionary<string, Button> Options { get; } = [];
    public event Action<string>? Selected;

    public PasteOptionsPanel(string current, bool hasFormats)
    {
        Width = 190; Padding = new(8); CornerRadius = new(4);
        Background = Brushes.White; BorderBrush = new SolidColorBrush(Color.Parse("#B8B8B8")); BorderThickness = new(1);
        BoxShadow = new BoxShadows(new BoxShadow { Blur = 12, OffsetY = 3, Color = Color.FromArgb(45, 0, 0, 0) });
        var body = new StackPanel { Spacing = 5 }; Child = body;
        void Group(string title, params (string Mode, string Label, string Badge, RibbonIcon Icon)[] items)
        {
            if (body.Children.Count > 0) body.Children.Add(new Border { Height = 1, Background = new SolidColorBrush(Color.Parse("#E4E4E4")), Margin = new(0, 3) });
            body.Children.Add(new TextBlock { Text = UiText.Get(title), FontWeight = FontWeight.SemiBold, Margin = new(3, 1) });
            var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5 }; body.Children.Add(row);
            foreach (var item in items)
            {
                var symbol = new global::Avalonia.Controls.Grid { Width = 28, Height = 28 };
                var glyph = RibbonIcons.Create(item.Icon); glyph.Foreground = new SolidColorBrush(Color.Parse("#447C67"));
                symbol.Children.Add(glyph);
                if (item.Badge.Length > 0) symbol.Children.Add(new Border { Background = Brushes.White, HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Bottom, Child = new TextBlock { Text = item.Badge, FontSize = 10, FontWeight = FontWeight.SemiBold, Foreground = new SolidColorBrush(Color.Parse("#275C83")) } });
                var button = new Button { Content = symbol, Width = 42, Height = 38, Padding = new(4), MinWidth = 0, MinHeight = 0,
                    IsEnabled = item.Mode != "formats" || hasFormats, BorderThickness = new(1),
                    BorderBrush = item.Mode == current ? new SolidColorBrush(Color.Parse("#81B599")) : Brushes.Transparent,
                    Background = item.Mode == current ? new SolidColorBrush(Color.Parse("#EDF6F0")) : Brushes.Transparent };
                ToolTip.SetTip(button, UiText.Get(item.Label)); AutomationProperties.SetName(button, UiText.Get(item.Label));
                button.Click += (_, _) => Selected?.Invoke(item.Mode);
                Options.Add(item.Mode, button); row.Children.Add(button);
            }
        }
        Group("粘贴", ("all", "粘贴全部", "", RibbonIcon.Paste), ("formulas", "粘贴公式", "fx", RibbonIcon.Paste));
        Group("粘贴数值", ("values", "只贴值", "123", RibbonIcon.Paste));
        Group("其他粘贴选项", ("formats", "只贴格式", "", RibbonIcon.FormatBrush), ("transpose", "转置粘贴", "↱", RibbonIcon.Paste));
    }

    public void FocusFirst() => Options.Values.First().Focus();
}
