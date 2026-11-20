using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Presenters;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly TextBox _recentSearch = new() { Name = "HomeRecentSearch", PlaceholderText = UiText.Get("搜索文件名或路径") };

    private ToggleButton CreateRecentPinButton(RecentFileEntry entry)
    {
        var label = UiText.Get(entry.IsPinned ? "取消固定" : "固定");
        var pin = new ToggleButton
        {
            Name = "PinRecentFile", Tag = entry.Path, IsChecked = entry.IsPinned,
            Width = 36, Height = 36, Margin = new(8, 0, 0, 0), Padding = new(0),
            HorizontalContentAlignment = HorizontalAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Background = Brush.Parse(entry.IsPinned ? "#D9EEE3" : "#F8FAF9"),
            Content = new global::Avalonia.Controls.Shapes.Path
            {
                // Keep the complete 20x20 view box so the narrow glyph stays centered.
                Data = Geometry.Parse("M6,2 H14 M7,2 V8 L4,12 V14 H16 V12 L13,8 V2 M10,14 V18"),
                Stroke = Brush.Parse(entry.IsPinned ? "#237A57" : "#66756D"), StrokeThickness = 1.7,
                Width = 20, Height = 20, Stretch = Stretch.None, IsHitTestVisible = false
            }
        };
        pin.Styles.Add(new Style(s => s.OfType<ToggleButton>().Class(":checked").Template().OfType<ContentPresenter>())
        {
            Setters = { new Setter(ContentPresenter.BackgroundProperty, Brush.Parse("#D9EEE3")),
                new Setter(ContentPresenter.BorderBrushProperty, Brush.Parse("#94B9A4")) }
        });
        ToolTip.SetTip(pin, label); AutomationProperties.SetName(pin, label + " · " + System.IO.Path.GetFileName(entry.Path));
        pin.Click += async (_, _) =>
        {
            pin.IsChecked = entry.IsPinned;
            await PerformAsync(() =>
            {
                _recentFiles.SetPinned(entry.Path, !entry.IsPinned);
                var error = _recentFiles.Error;
                RefreshRecentFiles();
                if (error is not null) _status.Text = error;
                return Task.CompletedTask;
            });
        };
        return pin;
    }

    private async Task OpenRecentFileAsync(string path)
    {
        if (await HandleMissingRecentFileAsync(path) || !await ConfirmDiscardAsync()) return;
        var previous = Document;
        try { await OpenPathAsync(path); }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException && ReferenceEquals(Document, previous))
        {
            // The file can disappear while the user handles unsaved changes.
            if (!await HandleMissingRecentFileAsync(path)) throw;
        }
    }

    private async Task<bool> HandleMissingRecentFileAsync(string path)
    {
        try
        {
            // Unlike File.Exists, this preserves permission and other IO errors.
            await Task.Run(() => File.GetAttributes(path));
            return false;
        }
        catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException) { }

        if (await Dialogs.MissingRecentFile(this, path))
        {
            _recentFiles.Remove(path);
            var error = _recentFiles.Error;
            RefreshRecentFiles();
            _status.Text = error ?? UiText.Get("已移除历史记录。");
        }
        return true;
    }

    private void BuildHome()
    {
        var actions = new StackPanel { Spacing = 14 };
        actions.Children.Add(new TextBlock { Text = "QuickExcel", FontSize = 38, FontWeight = FontWeight.SemiBold });
        actions.Children.Add(new TextBlock { Text = UiText.Get("工作簿与脚本"), FontSize = 18, Foreground = Brushes.Gray, Margin = new(0, 0, 0, 22) });
        void Action(string text, string geometry, Func<Task> run, string? name = null)
        {
            var icon = new global::Avalonia.Controls.Shapes.Path { Data = Geometry.Parse(geometry), Stroke = new SolidColorBrush(Color.Parse("#237A57")), StrokeThickness = 1.7, Width = 24, Height = 24 };
            var content = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 18 };
            content.Children.Add(new Viewbox { Width = 32, Height = 32, Child = icon });
            content.Children.Add(new TextBlock { Text = text, FontSize = 18, FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center });
            var button = new Button
            {
                Name = name ?? "HomeAction" + actions.Children.Count,
                Content = content,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                MinHeight = 76,
                Padding = new(22, 16),
                CornerRadius = new(8),
                Background = new SolidColorBrush(Color.Parse("#F1F7F4")),
                BorderBrush = new SolidColorBrush(Color.Parse("#D6E5DC")),
                BorderThickness = new(1)
            };
            AutomationProperties.SetName(button, text);
            button.Click += async (_, _) => await PerformAsync(run); actions.Children.Add(button);
        }
        Action(UiText.Get("新建工作簿"), "M14,3 H5 V21 H19 V8 Z M14,3 V8 H19 M8,14 H16 M12,10 V18", async () => { if (await ConfirmDiscardAsync()) { _workspace.New(); ShowBook(); } });
        Action(UiText.Get("打开工作簿"), "M3,8 V5 H9 L11,7 H21 V10 M3,8 H10 L12,10 H22 L19,20 H3 Z", OpenAsync);
        Action(UiText.Get("打开示例工作簿"), "M3,4 H21 V20 H3 Z M3,9 H21 M9,9 V20 M15,9 V20 M3,14 H21", OpenSampleWorkbooksAsync, "OpenSampleWorkbooks");
        Action(UiText.Get("返回当前工作簿"), "M10,4 H21 V20 H10 M14,12 H2 M7,7 L2,12 L7,17", () => { ShowBook(); return Task.CompletedTask; });
        Action(UiText.Get("设置"), "M3,6 H8 M13,6 H21 M3,12 H14 M19,12 H21 M3,18 H6 M11,18 H21 M8,4 H13 V8 H8 Z M14,10 H19 V14 H14 Z M6,16 H11 V20 H6 Z", OpenSettingsAsync);
        _start.Children.Add(new ScrollViewer
        {
            Content = actions,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        });
        var history = new global::Avalonia.Controls.Grid { RowDefinitions = new("Auto,Auto,Auto,*,Auto") };
        history.Children.Add(new TextBlock { Text = UiText.Get("历史文件"), FontSize = 24, FontWeight = FontWeight.SemiBold, Margin = new(0, 0, 0, 8) });
        var description = new TextBlock { Text = UiText.Get("继续处理最近打开的工作簿"), Foreground = Brushes.Gray, TextWrapping = TextWrapping.Wrap, Margin = new(0, 0, 0, 16) };
        global::Avalonia.Controls.Grid.SetRow(description, 1); history.Children.Add(description);
        _recentSearch.Margin = new(0, 0, 0, 12);
        AutomationProperties.SetName(_recentSearch, UiText.Get("搜索历史文件"));
        var scroll = new ScrollViewer
        {
            Name = "HomeRecentScroll",
            Content = _recent,
            // Reserve scrollbar width so its expanded track cannot cover the pin buttons.
            AllowAutoHide = false,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };
        _recentSearch.TextChanged += (_, _) => { FilterRecentFiles(); scroll.Offset = default; };
        global::Avalonia.Controls.Grid.SetRow(_recentSearch, 2); history.Children.Add(_recentSearch);
        global::Avalonia.Controls.Grid.SetRow(scroll, 3); history.Children.Add(scroll);
        var manage = new Button
        {
            Name = "ManageRecentFiles",
            Content = UiText.Get("管理历史文件"),
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new(0, 12, 0, 0)
        };
        manage.Click += async (_, _) => await PerformAsync(() => OpenSettingsAsync(SettingsSection.History));
        global::Avalonia.Controls.Grid.SetRow(manage, 4); history.Children.Add(manage);
        var container = new Border
        {
            Name = "HomeRecentContainer",
            Child = history,
            Padding = new(18),
            CornerRadius = new(8),
            Background = Brush.Parse("#F8FAF9"),
            BorderBrush = Brush.Parse("#D6E5DC"),
            BorderThickness = new(1)
        };
        global::Avalonia.Controls.Grid.SetColumn(container, 2); _start.Children.Add(container);
    }
}
