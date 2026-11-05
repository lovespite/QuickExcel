using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private IEnumerable<Button> RecentOpenButtons() => _recent.GetVisualDescendants().OfType<Button>().Where(b => b.Name == "OpenRecentFile");

    private async Task CheckHomeHistoryAsync(string directory, Action<bool, string> check, Func<Control, string, Task> capture)
    {
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        var scroll = _start.GetVisualDescendants().OfType<ScrollViewer>().Single(s => s.Name == "HomeRecentScroll");
        var verticalBar = scroll.GetVisualDescendants().OfType<ScrollBar>().Single(b => b.Orientation == global::Avalonia.Layout.Orientation.Vertical);
        void CheckPinClearance(string state)
        {
            var contentRight = verticalBar.IsVisible ? verticalBar.TranslatePoint(default, scroll)!.Value.X : scroll.Bounds.Width;
            var pins = _recent.GetVisualDescendants().OfType<ToggleButton>().Where(b => b.Name == "PinRecentFile").ToArray();
            check(pins.Length > 0 && pins.All(pin => pin.TranslatePoint(default, scroll)!.Value.X >= 0 &&
                pin.TranslatePoint(new(pin.Bounds.Width, 0), scroll)!.Value.X <= contentRight + .1),
                "History pin buttons stay entirely clear of the scrollbar: " + state);
        }
        var manage = _start.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "ManageRecentFiles");
        var container = _start.GetVisualDescendants().OfType<Border>().Single(b => b.Name == "HomeRecentContainer");
        var actionScroll = (ScrollViewer)_start.Children[0];
        check(!_start.GetVisualAncestors().OfType<ScrollViewer>().Any() && scroll.Extent.Height > scroll.Viewport.Height && scroll.Viewport.Height > 70,
            "At minimum size history overflows only its bounded list viewport");
        CheckPinClearance("minimum window");
        var searchPosition = _recentSearch.TranslatePoint(default, this)!.Value;
        var managePosition = manage.TranslatePoint(default, this)!.Value;
        check(managePosition.Y + manage.Bounds.Height <= _start.TranslatePoint(default, this)!.Value.Y + _start.Bounds.Height &&
            manage.TranslatePoint(default, container)!.Value.X + manage.Bounds.Width <= container.Bounds.Width,
            "Manage history remains inside the bottom-right of its container at minimum size");
        using var pointer = new Pointer(207, PointerType.Mouse, true);
        var first = _recent.Children[0];
        first.RaiseEvent(new PointerWheelEventArgs(first, pointer, this, first.TranslatePoint(new(20, 20), this)!.Value, 1,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.Other), KeyModifiers.None, new Vector(0, -3)));
        await Layout();
        check(scroll.Offset.Y > 0 && actionScroll.Offset.Y == 0 && _recentSearch.TranslatePoint(default, this) == searchPosition &&
            manage.TranslatePoint(default, this) == managePosition, "Scrolling history leaves the actions, search and management button in place");
        scroll.ScrollToEnd(); await Layout();
        check(scroll.Offset.Y > 0 && _recent.Children[^1].TranslatePoint(default, scroll)!.Value.Y + _recent.Children[^1].Bounds.Height <= scroll.Bounds.Height + 1,
            "The last history entry remains reachable inside the list");
        await capture(this, "home-history-scrolled");
        CheckPinClearance("scrolled to the last entry");

        var persisted = await File.ReadAllTextAsync(Path.Combine(directory, "settings", "recent-files.json"));
        _recentSearch.Text = "  历史工作簿-03.XLSX  "; await Layout();
        check(RecentOpenButtons().Count() == 1 && scroll.Offset.Y == 0,
            "Filename search ignores case and surrounding whitespace and resets list scrolling");
        check(!verticalBar.IsVisible, "A single filtered history entry does not need a scrollbar");
        CheckPinClearance("filtered without a scrollbar");
        _recentSearch.Text = "historyfolder-3"; await Layout();
        check(RecentOpenButtons().Count() == 1, "History search also matches directory paths case-insensitively");
        _recentSearch.Text = "年度预算"; await Layout();
        check(RecentOpenButtons().Count() == 2, "Unicode filename search retains matching files from different directories");
        _recentSearch.Text = "no-matching-workbook"; await Layout();
        check(_recent.Children.Single() is TextBlock noMatch && noMatch.Text == UiText.Get("没有匹配的历史文件。") && manage.IsVisible,
            "An unmatched query shows a distinct empty result while keeping history management available");
        _recentSearch.Text = "  "; await Layout();
        check(RecentOpenButtons().Count() == 10 && await File.ReadAllTextAsync(Path.Combine(directory, "settings", "recent-files.json")) == persisted,
            "Clearing the query restores the original order and never rewrites persisted history");
        Width = 1280; Height = 820; await Layout();
        check(scroll.Extent.Height > scroll.Viewport.Height && manage.TranslatePoint(default, this)!.Value.Y + manage.Bounds.Height <= Bounds.Height,
            "A full history list stays bounded after enlarging the window");
        CheckPinClearance("enlarged window");
        await capture(this, "home-history");

        _recentSearch.Text = "historyfolder-3"; await Layout();
        _dialogDriver = async window =>
        {
            await Layout();
            var navigation = window.GetVisualDescendants().OfType<ListBox>().Single(l => l.Name == "SettingsNavigation");
            var list = window.GetVisualDescendants().OfType<ListBox>().Single(l => l.Name == "SettingsRecentFiles");
            check(Equals(navigation.SelectedItem, UiText.Get("历史管理")) && list.ItemCount == 10,
                "Manage history opens Settings directly at the complete history section");
            var pageScroll = window.GetVisualDescendants().OfType<ScrollViewer>().Single(s => s.Name == "SettingsPageScroll");
            var listScroll = list.GetVisualDescendants().OfType<ScrollViewer>().Single();
            var remove = window.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, UiText.Get("移除选中记录")));
            void CheckHistoryLayout(string state)
            {
                check(pageScroll.Extent.Height <= pageScroll.Viewport.Height + .1 && listScroll.Viewport.Height >= 80 &&
                    listScroll.Extent.Height > listScroll.Viewport.Height,
                    "Settings history scrolls only inside the file list: " + state);
                check(remove.TranslatePoint(default, window)!.Value.Y >= list.TranslatePoint(default, window)!.Value.Y + list.Bounds.Height &&
                    remove.TranslatePoint(new(0, remove.Bounds.Height), window)!.Value.Y <= window.Bounds.Height,
                    "History management actions remain visible below the list: " + state);
            }
            CheckHistoryLayout("default size");
            await capture(window, "settings-history");
            window.Width = window.MinWidth; window.Height = window.MinHeight; await Layout();
            CheckHistoryLayout("minimum size");
            await capture(window, "settings-history-minimum");
            window.Width = 820; window.Height = 600; await Layout();
            listScroll.ScrollToEnd(); await Layout();
            CheckHistoryLayout("scrolled to the last entry");
            navigation.SelectedItem = UiText.Get("保存与计算"); await Layout();
            check(pageScroll.VerticalScrollBarVisibility == ScrollBarVisibility.Auto,
                "Other settings sections retain their page scrolling after leaving history");
            navigation.SelectedItem = UiText.Get("历史管理"); await Layout();
            list = window.GetVisualDescendants().OfType<ListBox>().Single(l => l.Name == "SettingsRecentFiles");
            list.SelectedItem = _recentFiles.Paths.Single(path => path.Contains("HistoryFolder-3", StringComparison.Ordinal));
            window.GetVisualDescendants().OfType<Button>().Single(b => Equals(b.Content, UiText.Get("移除选中记录")))
                .RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            window.Close(true);
        };
        manage.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); await _operation; await Layout();
        check(_recentSearch.Text == "historyfolder-3" && _recent.Children.Single() is TextBlock && _recentFiles.Paths.Count == 9,
            "Returning from history management refreshes results and preserves the active search");
        _recentSearch.Text = ""; await Layout();
        check(RecentOpenButtons().Count() == 9, "Clearing the search reflects removals made in Settings");
        await CheckRecentFileFeaturesAsync(directory, check, capture);
    }
}
