using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Threading;
using Avalonia.VisualTree;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckResponsiveRibbonAsync(Action<bool, string> check, Func<Control, string, Task> capture)
    {
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static RibbonCommandContent[] Commands(RibbonPage page) => page.GetVisualDescendants().OfType<RibbonCommandContent>().ToArray();
        static bool Scrolls(RibbonPage page) => page.Extent.Width > page.Viewport.Width + .1;
        static bool BarVisible(RibbonPage page) => page.GetVisualDescendants().OfType<ScrollBar>()
            .Single(bar => bar.Name == "PART_HorizontalScrollBar" && ReferenceEquals(bar.TemplatedParent, page)).IsVisible;
        var originalWidth = Width;
        try
        {
            Width = 1800;
            foreach (var tab in _ribbon.Items.Cast<TabItem>())
            {
                _ribbon.SelectedItem = tab; await Layout();
                var page = (RibbonPage)tab.Content!;
                var commands = Commands(page);
                if (Equals(tab.Header, UiText.Get("开始")))
                {
                    check(commands.Length > 0 && commands.All(c => !c.Label.IsVisible), "Home keeps icon-only commands at wide widths");
                    continue;
                }
                check(commands.Length > 0 && commands.All(c => c.Label.IsVisible) && !Scrolls(page) && !BarVisible(page),
                    $"Wide {tab.Header} shows command labels without a scrollbar");
                foreach (var command in commands)
                {
                    var button = command.GetVisualAncestors().OfType<ContentControl>().First(c => c.Classes.Contains("ribbon-command"));
                    check(command.Label.Text == AutomationProperties.GetName(button) && command.Label.Bounds.Width + .1 >= command.Label.DesiredSize.Width,
                        $"Full command name is visible: {command.Label.Text}");
                }
                await capture(_ribbon, $"ribbon-labeled-{_ribbon.SelectedIndex}");
                var fullWidth = ((Control)page.Content!).DesiredSize.Width;
                try
                {
                    page.Width = fullWidth / 2; await Layout();
                    var compactWidth = ((Control)page.Content!).DesiredSize.Width;
                    check(commands.All(c => !c.Label.IsVisible) && compactWidth < fullWidth,
                        $"{tab.Header} hides labels before allowing overflow");
                    page.Width = (fullWidth + compactWidth) / 2; await Layout();
                    check(commands.All(c => !c.Label.IsVisible) && !Scrolls(page) && !BarVisible(page),
                        $"{tab.Header} fits with icons and no scrollbar at intermediate width");
                    if (Equals(tab.Header, UiText.Get("数据"))) await capture(_ribbon, "ribbon-data-compact");
                    // Exercise widths below the main window minimum using only the page viewport.
                    page.Width = compactWidth - 30; await Layout();
                    check(commands.All(c => !c.Label.IsVisible) && Scrolls(page) && BarVisible(page),
                        $"{tab.Header} scrolls only after its icons no longer fit");
                    page.Offset = new(page.Extent.Width - page.Viewport.Width, 0); await Layout();
                    check(page.Offset.X > 0, $"Last compact commands remain reachable on {tab.Header}");
                    if (Equals(tab.Header, UiText.Get("数据"))) await capture(_ribbon, "ribbon-data-scrolling");
                    for (var round = 0; round < 3; round++)
                    {
                        page.Width = fullWidth + 2; await Layout();
                        check(commands.All(c => c.Label.IsVisible) && !Scrolls(page) && !BarVisible(page) && page.Offset.X == 0,
                            $"{tab.Header} restores labels and resets scrolling when widened ({round})");
                        page.Width = fullWidth - 2; await Layout();
                        check(commands.All(c => !c.Label.IsVisible) && !Scrolls(page) && !BarVisible(page),
                            $"{tab.Header} stays compact near the label threshold ({round})");
                        page.InvalidateMeasure(); await Layout();
                        check(commands.All(c => !c.Label.IsVisible), $"{tab.Header} does not oscillate after another layout pass ({round})");
                    }
                }
                finally { page.ClearValue(WidthProperty); }
            }
            _ribbon.SelectedItem = _ribbon.Items.Cast<TabItem>().Single(t => Equals(t.Header, UiText.Get("数据")));
            await Layout();
            check(Commands((RibbonPage)((TabItem)_ribbon.SelectedItem!).Content!).All(c => c.Label.IsVisible),
                "Returning to a tab restores labels for the current viewport");
        }
        finally { Width = originalWidth; _ribbon.SelectedIndex = 0; await Layout(); }
    }
}
