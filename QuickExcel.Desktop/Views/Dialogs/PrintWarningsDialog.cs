using System.Windows;
using System.Windows.Controls;

namespace QuickExcel.Desktop;

internal sealed class PrintWarningsDialog : Window
{
    internal PrintWarningsDialog(Window owner, IReadOnlyList<string> warnings)
    {
        Owner = owner; Title = "打印内容提示"; Width = 620; Height = 420;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var root = new DockPanel { Margin = new(20) };
        var proceed = new Button { Content = "继续预览", IsDefault = true, MinWidth = 100, Padding = new(12, 6, 12, 6),
            HorizontalAlignment = HorizontalAlignment.Right, Margin = new(0, 16, 0, 0) };
        proceed.Click += (_, _) => Close();
        DockPanel.SetDock(proceed, Dock.Bottom); root.Children.Add(proceed);
        var description = new TextBlock { Text = "以下内容暂时无法完整渲染，将按说明跳过或简化。其余内容可继续预览、打印或导出 PDF，工作簿原内容保持不变。",
            TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 12) };
        DockPanel.SetDock(description, Dock.Top); root.Children.Add(description);
        root.Children.Add(new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = new TextBlock { Text = string.Join("\n\n", warnings.Select(w => "• " + w)), TextWrapping = TextWrapping.Wrap } });
        Content = root;
    }
}
