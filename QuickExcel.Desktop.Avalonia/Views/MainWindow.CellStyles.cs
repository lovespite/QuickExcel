using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private QuickCellStyleStore _quickStyles = null!;
    private readonly List<(QuickCellStyle Style, Button Button)> _quickStyleButtons = [];
    private Flyout? _cellStyleFlyout;

    private void InstallCellStyles()
    {
        _quickStyles = new(Path.Combine(_userDataDirectory, "cell-styles.json")); _quickStyles.Load();
        var body = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Center };
        body.Children.Add(ConditionalFormatDropdown());
        var gallery = new global::Avalonia.Controls.Grid { ColumnDefinitions = new("*,24"), Height = 76 };
        var tiles = new UniformGrid { Columns = 5, Rows = 2, Width = 430 };
        foreach (var style in QuickCellStyles.BuiltIn.Take(10))
        {
            var button = CellStyleButton(style, false); _quickStyleButtons.Add((style, button)); tiles.Children.Add(button);
        }
        void FitGallery()
        {
            var columns = Math.Clamp((int)((ClientSize.Width - 1000) / 86), 2, 5);
            tiles.Columns = columns; tiles.Width = columns * 86;
            for (var i = 0; i < tiles.Children.Count; i++) tiles.Children[i].IsVisible = i < columns * 2;
        }
        PropertyChanged += (_, e) => { if (e.Property == ClientSizeProperty) FitGallery(); }; FitGallery();
        gallery.Children.Add(tiles);
        _cellStyleFlyout = new Flyout { Placement = PlacementMode.BottomEdgeAlignedRight };
        _cellStyleFlyout.FlyoutPresenterTheme = new ControlTheme(typeof(FlyoutPresenter))
        {
            BasedOn = this.FindResource(typeof(FlyoutPresenter)) as ControlTheme,
            Setters = { new Setter(MaxWidthProperty, 600d), new Setter(MaxHeightProperty, 650d) }
        };
        _cellStyleFlyout.Opening += (_, _) => { _quickStyles.Load(); _cellStyleFlyout.Content = BuildCellStyleGallery(); };
        var more = new Button { Name = "CellStylesMore", Content = "▾", Padding = new(0), HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch,
            Flyout = _cellStyleFlyout, Background = Brushes.Transparent, BorderThickness = new(1, 0, 0, 0), BorderBrush = Brushes.LightGray };
        AutomationProperties.SetName(more, UiText.Get("更多单元格样式")); ToolTip.SetTip(more, UiText.Get("更多单元格样式"));
        global::Avalonia.Controls.Grid.SetColumn(more, 1); gallery.Children.Add(more);
        var frame = new Border { Name = "CellStyleGallery", Child = gallery, BorderBrush = Brushes.LightGray, BorderThickness = new(1), CornerRadius = new(2) };
        body.Children.Add(frame); _mutationControls.Add(frame);
        RibbonContentGroup(UiText.Get("开始"), UiText.Get("样式"), body);
    }

    private Control BuildCellStyleGallery()
    {
        var root = new StackPanel { Width = 548, Spacing = 8 }; var groups = new StackPanel { Spacing = 10 };
        foreach (var group in QuickCellStyles.BuiltIn.Concat(_quickStyles.Styles).GroupBy(s => s.Group))
        {
            groups.Children.Add(new TextBlock { Text = UiText.Get(group.Key), FontWeight = FontWeight.SemiBold, Margin = new(4, 2) });
            var tiles = new UniformGrid { Columns = 4 };
            foreach (var style in group) tiles.Children.Add(CellStyleButton(style, true)); groups.Children.Add(tiles);
        }
        root.Children.Add(new ScrollViewer { Content = groups, MaxHeight = 450, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled });
        if (_quickStyles.Error is { } error) root.Children.Add(new TextBlock { Text = error, TextWrapping = TextWrapping.Wrap, Foreground = Brushes.Firebrick });
        root.Children.Add(new Separator());
        var footer = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        var create = new Button { Name = "NewCellStyle", Content = UiText.Get("新建单元格样式…") };
        create.Click += async (_, _) => { _cellStyleFlyout!.Hide(); await PerformAsync(() => EditCellStyleAsync()); };
        var manage = new Button { Name = "ManageCellStyles", Content = UiText.Get("管理自定义样式…"), IsEnabled = !_quickStyles.Styles.IsEmpty };
        manage.Click += async (_, _) => { _cellStyleFlyout!.Hide(); await PerformAsync(ManageCellStylesAsync); };
        footer.Children.Add(create); footer.Children.Add(manage); root.Children.Add(footer);
        root.Children.Add(new TextBlock { Text = UiText.Get("自定义样式保存在本机；右键样式可复制、修改或删除。"), FontSize = 12, Foreground = Brushes.Gray, TextWrapping = TextWrapping.Wrap }); return root;
    }

    private Button CellStyleButton(QuickCellStyle style, bool expanded)
    {
        var name = QuickCellStyles.DisplayName(style);
        var button = new Button { Name = "CellStyle_" + style.Id, Content = CellStyleSample(style, name), Padding = new(3), Margin = new(1),
            Height = expanded ? 38 : 36, MinHeight = 0, Background = Brushes.Transparent, BorderBrush = Brushes.Transparent, BorderThickness = new(1),
            HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Stretch, VerticalContentAlignment = VerticalAlignment.Stretch };
        AutomationProperties.SetName(button, name); ToolTip.SetTip(button, name);
        button.Click += async (_, _) => { _cellStyleFlyout?.Hide(); await PerformAsync(() => ApplyStyleAsync(style.Patch)); };
        var context = new ContextMenu();
        var copy = new MenuItem { Header = UiText.Get("复制为新样式…") };
        copy.Click += async (_, _) => { _cellStyleFlyout?.Hide(); await PerformAsync(() => EditCellStyleAsync(style, true)); }; context.Items.Add(copy);
        if (_quickStyles.Styles.Any(s => s.Id == style.Id))
        {
            var edit = new MenuItem { Header = UiText.Get("修改样式…") }; edit.Click += async (_, _) => { _cellStyleFlyout?.Hide(); await PerformAsync(() => EditCellStyleAsync(style)); };
            var delete = new MenuItem { Header = UiText.Get("删除样式") }; delete.Click += async (_, _) => { _cellStyleFlyout?.Hide(); await PerformAsync(() => DeleteCellStyleAsync(style)); };
            context.Items.Add(edit); context.Items.Add(delete);
        }
        button.ContextMenu = context; return button;
    }

    internal static Control CellStyleSample(QuickCellStyle style, string name)
    {
        var s = style.Format;
        var text = new TextBlock { Text = name, FontSize = Math.Clamp(s.FontSize ?? 12, 11, 16),
            FontFamily = new FontFamily(s.FontName ?? "Microsoft YaHei UI"),
            FontWeight = s.Bold ? FontWeight.Bold : FontWeight.Normal, FontStyle = s.Italic ? FontStyle.Italic : FontStyle.Normal,
            Foreground = Brush.Parse("#" + (s.FontColor ?? "FF1F2937")), VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis };
        if (s.Underline is not (null or "none")) text.TextDecorations = TextDecorations.Underline;
        if (s.Strike) text.TextDecorations = new TextDecorationCollection((text.TextDecorations ?? []).Concat(TextDecorations.Strikethrough));
        return new Border { Child = text, Padding = new(4, 0), Background = Brush.Parse("#" + (s.FillColor ?? "FFFFFFFF")),
            BorderThickness = new(s.LeftBorder > 0 ? 1 : 0, s.TopBorder > 0 ? 1 : 0, s.RightBorder > 0 ? 1 : 0, s.BottomBorder > 0 ? (s.BottomBorder == 6 ? 2 : 1) : 0),
            BorderBrush = Brush.Parse("#" + (s.BottomBorderColor ?? s.TopBorderColor ?? s.LeftBorderColor ?? "FF9A9A9A")) };
    }

    private void UpdateCellStyleSelection()
    {
        foreach (var (style, button) in _quickStyleButtons)
        {
            var candidate = style with { Format = CurrentStyle };
            var selected = candidate.Patch == style.Patch;
            button.BorderBrush = selected ? Brush.Parse("#217346") : Brushes.Transparent;
            AutomationProperties.SetHelpText(button, selected ? UiText.Get("当前单元格样式") : QuickCellStyles.DisplayName(style));
        }
    }

    private async Task EditCellStyleAsync(QuickCellStyle? source = null, bool copy = false)
    {
        if (!await CommitEditorsAsync()) return;
        _quickStyles.Load(); if (_quickStyles.Error is { } error) throw new InvalidOperationException(error);
        var session = Document.Session.SessionId;
        var existing = source is not null && !copy ? _quickStyles.Styles.SingleOrDefault(s => s.Id == source.Id) : null;
        if (source is not null && !copy && existing is null) throw new InvalidOperationException(UiText.Get("样式已删除，请重新打开样式库。"));
        string NextName()
        {
            for (var i = 1; ; i++) { var name = UiText.Format($"样式 {i}"); if (!_quickStyles.Styles.Any(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))) return name; }
        }
        var seed = existing ?? new QuickCellStyle(Guid.NewGuid().ToString("N"), NextName(), "自定义", (source?.Format ?? CurrentStyle) with { NamedStyleId = 0 }, source?.Parts ?? CellStyleParts.All);
        QuickCellStyle[] Updated(QuickCellStyle value) => _quickStyles.Styles.Where(s => s.Id != value.Id).Append(value).ToArray();
        var dialog = new CellStyleDialog(seed, FormatSample, Document.Snapshot.DateSystem, Document.Snapshot.Styles, value => QuickCellStyleStore.Validate(Updated(value)));
        if (await ShowOwnedAsync<QuickCellStyle?>(dialog) is not { } result) return;
        if (session != Document.Session.SessionId) throw new InvalidOperationException(UiText.Get("工作簿已切换，请重新设置格式。"));
        _quickStyles.Save(Updated(result)); _status.Text = UiText.Get("自定义样式已保存到本机。选择样式即可应用到单元格。");
    }

    private Task DeleteCellStyleAsync(QuickCellStyle style)
    {
        _quickStyles.Save(_quickStyles.Styles.Where(s => s.Id != style.Id));
        _status.Text = UiText.Get("已删除本机样式，已应用的单元格格式保持不变。"); return Task.CompletedTask;
    }

    private async Task ManageCellStylesAsync()
    {
        _quickStyles.Load();
        var window = new Window { Title = UiText.Get("管理自定义样式"), Width = 480, Height = 440, WindowStartupLocation = WindowStartupLocation.CenterOwner, ShowInTaskbar = false };
        var root = new DockPanel { Margin = new(18), LastChildFill = true };
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new(0, 12, 0, 0) }; DockPanel.SetDock(buttons, Dock.Bottom); root.Children.Add(buttons);
        var error = new TextBlock { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap }; DockPanel.SetDock(error, Dock.Bottom); root.Children.Add(error);
        var list = new ListBox { Name = "CustomCellStyles" }; root.Children.Add(list);
        void Refresh() { list.ItemsSource = _quickStyles.Styles.Select(s => s.Name).ToArray(); if (list.ItemCount > 0) list.SelectedIndex = 0; }
        var edit = new Button { Name = "EditCustomCellStyle", Content = UiText.Get("修改样式…") };
        var delete = new Button { Name = "DeleteCustomCellStyle", Content = UiText.Get("删除样式") };
        var close = new Button { Content = UiText.Get("关闭"), IsCancel = true }; close.Click += (_, _) => window.Close((string?)null);
        edit.Click += (_, _) => { if (list.SelectedIndex >= 0) window.Close(_quickStyles.Styles[list.SelectedIndex].Id); };
        delete.Click += (_, _) =>
        {
            if (list.SelectedIndex < 0) return;
            try { _quickStyles.Save(_quickStyles.Styles.RemoveAt(list.SelectedIndex)); Refresh(); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidOperationException) { error.Text = ex.Message; }
        };
        list.SelectionChanged += (_, _) => edit.IsEnabled = delete.IsEnabled = list.SelectedIndex >= 0;
        buttons.Children.Add(edit); buttons.Children.Add(delete); buttons.Children.Add(close); Refresh(); window.Content = root;
        if (await ShowOwnedAsync<string?>(window) is { } id && _quickStyles.Styles.SingleOrDefault(s => s.Id == id) is { } style) await EditCellStyleAsync(style);
    }
}
