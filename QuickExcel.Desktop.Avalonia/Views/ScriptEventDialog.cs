using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Presentation.Web;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class ScriptEventDialog : Window
{
    internal TextBox Search { get; } = new() { PlaceholderText = UiText.Get("搜索事件名称或说明") };
    internal ListBox Events { get; } = new() { MinHeight = 160 };
    internal TextBlock Description { get; } = new() { TextWrapping = TextWrapping.Wrap };
    internal TextBlock Timing { get; } = new() { TextWrapping = TextWrapping.Wrap };
    internal Button Confirm { get; } = new() { Content = UiText.Get("确定"), IsDefault = true };

    public ScriptEventDialog()
    {
        Title = UiText.Get("添加事件处理器"); Width = 610; Height = 590; MinWidth = 470; MinHeight = 450;
        Background = Brush.Parse("#1E1E1E"); Foreground = Brush.Parse("#CCCCCC");
        Events.Background = Search.Background = Brush.Parse("#252526");
        WindowStartupLocation = WindowStartupLocation.CenterOwner; ShowInTaskbar = false;
        var body = new LayoutGrid { RowDefinitions = new("Auto,*,Auto,Auto"), Margin = new(20), RowSpacing = 12 };
        AutomationProperties.SetName(Search, UiText.Get("搜索事件名称或说明"));
        AutomationProperties.SetName(Events, UiText.Get("可用事件"));
        body.Children.Add(Search);
        Events.ItemTemplate = new FuncDataTemplate<ScriptEventTemplate>((item, _) => new TextBlock { Text = item?.Name, FontFamily = new("Cascadia Mono, Consolas"), Margin = new(4) });
        LayoutGrid.SetRow(Events, 1); body.Children.Add(Events);
        var details = new StackPanel { Spacing = 6, MinHeight = 112 };
        details.Children.Add(new TextBlock { Text = UiText.Get("事件说明"), FontWeight = FontWeight.SemiBold }); details.Children.Add(Description);
        details.Children.Add(new TextBlock { Text = UiText.Get("触发时机"), FontWeight = FontWeight.SemiBold }); details.Children.Add(Timing);
        LayoutGrid.SetRow(details, 2); body.Children.Add(details);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10, HorizontalAlignment = HorizontalAlignment.Right };
        var cancel = new Button { Content = UiText.Get("取消"), IsCancel = true };
        Confirm.Click += (_, _) => { if (Events.SelectedItem is ScriptEventTemplate selected) Close(selected); };
        cancel.Click += (_, _) => Close(null); buttons.Children.Add(Confirm); buttons.Children.Add(cancel);
        LayoutGrid.SetRow(buttons, 3); body.Children.Add(buttons); Content = body;
        Events.SelectionChanged += (_, _) =>
        {
            var item = Events.SelectedItem as ScriptEventTemplate; Confirm.IsEnabled = item is not null;
            Description.Text = item?.Description ?? UiText.Get("没有匹配的事件。"); Timing.Text = item?.Timing ?? "";
        };
        Search.TextChanged += (_, _) => Filter(); Filter();
        Opened += (_, _) => Search.Focus();
    }
    private void Filter()
    {
        var previous = Events.SelectedItem;
        var items = ScriptEventCatalog.All.Where(e => e.Matches(Search.Text ?? "")).ToArray();
        Events.ItemsSource = items;
        Events.SelectedItem = previous is ScriptEventTemplate selected && items.Contains(selected) ? selected : items.FirstOrDefault();
        Confirm.IsEnabled = Events.SelectedItem is not null;
        if (items.Length == 0) { Description.Text = UiText.Get("没有匹配的事件。"); Timing.Text = ""; }
    }
}
