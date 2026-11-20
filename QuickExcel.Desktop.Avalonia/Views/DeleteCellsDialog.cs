using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class DeleteCellsDialog : Window
{
    public DeleteCellsDialog()
    {
        Title = UiText.Get("删除单元格"); Width = 340; CanResize = false;
        SizeToContent = SizeToContent.Height; ShowInTaskbar = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var panel = new StackPanel { Spacing = 10, Margin = new Thickness(20) };
        var choices = new[] { "右侧单元格左移", "下方单元格上移", "整行", "整列" }
            .Select((label, index) => new RadioButton
            {
                Name = "DeleteCellsChoice" + index, Content = UiText.Get(label),
                GroupName = "DeleteCells", IsChecked = index == 0
            }).ToArray();
        foreach (var radio in choices)
        {
            AutomationProperties.SetName(radio, radio.Content as string);
            panel.Children.Add(radio);
        }
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, HorizontalAlignment = HorizontalAlignment.Right, Margin = new(0, 12, 0, 0) };
        var ok = new Button { Name = "DeleteCellsConfirm", Content = UiText.Get("确定"), IsDefault = true };
        var cancel = new Button { Name = "DeleteCellsCancel", Content = UiText.Get("取消"), IsCancel = true };
        ok.Click += (_, _) => Close((int?)Array.FindIndex(choices, c => c.IsChecked == true));
        cancel.Click += (_, _) => Close((int?)null);
        buttons.Children.Add(ok); buttons.Children.Add(cancel); panel.Children.Add(buttons);
        Content = panel;
        Opened += (_, _) => choices[0].Focus();
    }
}
