using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private void BuildHome()
    {
        var actions = new StackPanel { Spacing = 14 };
        actions.Children.Add(new TextBlock { Text = "QuickExcel", FontSize = 38, FontWeight = FontWeight.SemiBold });
        actions.Children.Add(new TextBlock { Text = "工作簿与脚本", FontSize = 18, Foreground = Brushes.Gray, Margin = new(0, 0, 0, 22) });
        void Action(string text, string geometry, Func<Task> run)
        {
            var icon = new global::Avalonia.Controls.Shapes.Path { Data = Geometry.Parse(geometry), Stroke = new SolidColorBrush(Color.Parse("#237A57")), StrokeThickness = 1.7, Width = 24, Height = 24 };
            var content = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 18 };
            content.Children.Add(new Viewbox { Width = 32, Height = 32, Child = icon });
            content.Children.Add(new TextBlock { Text = text, FontSize = 18, FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center });
            var button = new Button { Name = "HomeAction" + actions.Children.Count, Content = content, HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Left,
                MinHeight = 76, Padding = new(22, 16), CornerRadius = new(8), Background = new SolidColorBrush(Color.Parse("#F1F7F4")),
                BorderBrush = new SolidColorBrush(Color.Parse("#D6E5DC")), BorderThickness = new(1) };
            AutomationProperties.SetName(button, text);
            button.Click += async (_, _) => await PerformAsync(run); actions.Children.Add(button);
        }
        Action("新建工作簿", "M14,3 H5 V21 H19 V8 Z M14,3 V8 H19 M8,14 H16 M12,10 V18", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } });
        Action("打开工作簿", "M3,8 V5 H9 L11,7 H21 V10 M3,8 H10 L12,10 H22 L19,20 H3 Z", OpenAsync);
        Action("打开示例", "M3,4 H21 V20 H3 Z M3,9 H21 M9,9 V20 M15,9 V20 M3,14 H21", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(true); ShowBook(); } });
        Action("返回当前工作簿", "M10,4 H21 V20 H10 M14,12 H2 M7,7 L2,12 L7,17", () => { ShowBook(); return Task.CompletedTask; });
        Action("设置", "M3,6 H8 M13,6 H21 M3,12 H14 M19,12 H21 M3,18 H6 M11,18 H21 M8,4 H13 V8 H8 Z M14,10 H19 V14 H14 Z M6,16 H11 V20 H6 Z", OpenSettingsAsync);
        _start.Children.Add(actions);
        var history = new StackPanel { Spacing = 18, Margin = new(0, 8, 0, 0) };
        history.Children.Add(new TextBlock { Text = "历史文件", FontSize = 24, FontWeight = FontWeight.SemiBold });
        history.Children.Add(new TextBlock { Text = "继续处理最近打开的工作簿", Foreground = Brushes.Gray, Margin = new(0, 0, 0, 12) });
        history.Children.Add(_recent);
        global::Avalonia.Controls.Grid.SetColumn(history, 2); _start.Children.Add(history);
    }
}
