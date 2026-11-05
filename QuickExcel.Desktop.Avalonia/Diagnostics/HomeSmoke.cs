using Avalonia;
using Avalonia.Controls;
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
        var manage = _start.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "ManageRecentFiles");
        var container = _start.GetVisualDescendants().OfType<Border>().Single(b => b.Name == "HomeRecentContainer");
        var actionScroll = (ScrollViewer)_start.Children[0];
        check(!_start.GetVisualAncestors().OfType<ScrollViewer>().Any() && scroll.Extent.Height > scroll.Viewport.Height && scroll.Viewport.Height > 70,
            "At minimum size history overflows only its bounded list viewport");
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

        var persisted = await File.ReadAllTextAsync(Path.Combine(directory, "settings", "recent-files.json"));
        _recentSearch.Text = "  历史工作簿-03.XLSX  "; await Layout();
        check(RecentOpenButtons().Count() == 1 && scroll.Offset.Y == 0,
            "Filename search ignores case and surrounding whitespace and resets list scrolling");
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
        await capture(this, "home-history");

        _recentSearch.Text = "historyfolder-3"; await Layout();
        _dialogDriver = async window =>
        {
            await Layout();
            var navigation = window.GetVisualDescendants().OfType<ListBox>().Single(l => l.Name == "SettingsNavigation");
            var list = window.GetVisualDescendants().OfType<ListBox>().Single(l => l.Name == "SettingsRecentFiles");
            check(Equals(navigation.SelectedItem, UiText.Get("历史管理")) && list.ItemCount == 10,
                "Manage history opens Settings directly at the complete history section");
            await capture(window, "settings-history");
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
