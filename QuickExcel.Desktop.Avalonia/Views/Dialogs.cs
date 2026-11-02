using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;

namespace QuickExcel.Desktop.Avalonia;

internal enum SaveDecision { Cancel, Save, Discard }
internal static class Dialogs
{
    public static Task<SaveDecision> Unsaved(Window owner) => Choose(owner, "保存更改", "保存当前工作簿后再继续？",
        [("保存", SaveDecision.Save), ("不保存", SaveDecision.Discard), ("取消", SaveDecision.Cancel)], SaveDecision.Cancel);
    public static Task<bool> Confirm(Window owner, string title, string message) => Choose(owner, title, message,
        [("确定", true), ("取消", false)], false);
    public static async Task<string?> Text(Window owner, string title, string label, string initial)
    {
        var input = new TextBox { Text = initial, MinWidth = 300 };
        var window = Window(owner, title); var panel = new StackPanel { Spacing = 14, Margin = new(20) };
        panel.Children.Add(new TextBlock { Text = label }); panel.Children.Add(input);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, HorizontalAlignment = HorizontalAlignment.Right };
        var ok = new Button { Content = "确定", IsDefault = true }; var cancel = new Button { Content = "取消", IsCancel = true };
        ok.Click += (_, _) => window.Close(input.Text); cancel.Click += (_, _) => window.Close(null);
        buttons.Children.Add(ok); buttons.Children.Add(cancel); panel.Children.Add(buttons); window.Content = panel;
        window.Opened += (_, _) => { input.Focus(); input.SelectAll(); };
        return await window.ShowDialog<string?>(owner);
    }
    private static Task<T> Choose<T>(Window owner, string title, string message, (string Text, T Value)[] choices, T cancelled)
    {
        var window = Window(owner, title); var selected = false;
        var panel = new StackPanel { Spacing = 20, Margin = new(20), MaxWidth = 620 };
        panel.Children.Add(new ScrollViewer { MaxHeight = 380, Content = new TextBlock { Text = message, TextWrapping = global::Avalonia.Media.TextWrapping.Wrap } });
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, HorizontalAlignment = HorizontalAlignment.Right };
        foreach (var choice in choices)
        {
            var button = new Button { Content = choice.Text, MinWidth = 80 };
            button.Click += (_, _) => { selected = true; window.Close(choice.Value); }; buttons.Children.Add(button);
        }
        panel.Children.Add(buttons); window.Content = panel;
        window.Closing += (_, e) => { if (!selected) { e.Cancel = true; selected = true; window.Close(cancelled); } };
        return window.ShowDialog<T>(owner);
    }
    private static Window Window(Window owner, string title) => new()
    { Title = title, SizeToContent = SizeToContent.WidthAndHeight, WindowStartupLocation = WindowStartupLocation.CenterOwner, CanResize = false, ShowInTaskbar = false };
}
