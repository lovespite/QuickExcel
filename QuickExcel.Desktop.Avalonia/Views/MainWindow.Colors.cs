using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private Border? _fontColorSwatch, _fillColorSwatch;
    // Window-level tool choices survive selection changes and document undo/redo.
    private string? _lastFontColor;
    private string? _lastFillColor = "FFFFFF00";

    private Control[] ColorMenuItems(string resetLabel, Func<string?, Task> apply, Func<Task> custom, Func<string, Func<Task>, MenuItem>? createItem = null)
    {
        createItem ??= CommandMenu;
        var items = new List<Control>();
        var reset = createItem(resetLabel, () => apply(null)); reset.Icon = ColorPalette.Swatch(null); items.Add(reset);
        items.Add(new Separator());
        foreach (var (name, value) in ColorPalette.Common)
        {
            var item = createItem(name, () => apply(value)); item.Icon = ColorPalette.Swatch(value); items.Add(item);
        }
        items.Add(new Separator()); items.Add(createItem(UiText.Get("自定义颜色…"), custom)); return items.ToArray();
    }

    private SplitButton RibbonColorDropdown(bool fill)
    {
        var title = fill ? UiText.Get("填充颜色") : UiText.Get("字体颜色");
        var swatch = ColorPalette.Swatch(fill ? _lastFillColor : _lastFontColor ?? "FF000000");
        if (fill) _fillColorSwatch = swatch; else _fontColorSwatch = swatch;
        var content = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5 };
        content.Children.Add(RibbonIcons.Create(fill ? RibbonIcon.FillColor : RibbonIcon.FontColor)); content.Children.Add(swatch);
        var button = new SplitButton { Name = fill ? "FillColorDropdown" : "FontColorDropdown", Content = content,
            Height = 28, MinHeight = 28, Padding = new(4, 2), Margin = new(2, 1), Background = Brushes.Transparent, BorderThickness = new(0),
            HorizontalAlignment = HorizontalAlignment.Left, HorizontalContentAlignment = HorizontalAlignment.Left,
            Flyout = new MenuFlyout { ItemsSource = ColorMenuItems(fill ? UiText.Get("无填充") : UiText.Get("自动字体色"), color => ApplyCellColorAsync(fill, color), () => CustomCellColorAsync(fill)) } };
        button.Classes.Add("ribbon-command");
        button.Click += async (_, _) => await PerformAsync(() => ApplyCellColorAsync(fill, fill ? _lastFillColor : _lastFontColor));
        AutomationProperties.SetName(button, title); ToolTip.SetTip(button, title); _mutationControls.Add(button); return button;
    }

    private void UpdateColorSwatches()
    {
        void Update(Border? swatch, string? color)
        {
            if (swatch is null) return;
            var sample = ColorPalette.Swatch(color); var child = sample.Child; sample.Child = null;
            swatch.Background = sample.Background; swatch.Child = child;
        }
        Update(_fontColorSwatch, _lastFontColor ?? "FF000000"); Update(_fillColorSwatch, _lastFillColor);
    }

    private void RememberCellColor(bool fill, string? color)
    {
        if (fill) _lastFillColor = color; else _lastFontColor = color;
        UpdateColorSwatches();
    }

    private async Task ApplyCellColorAsync(bool fill, string? color)
    {
        if (!await CommitEditorsAsync()) return;
        if (await ExecuteAsync(new SetCellStyleCommand(VisibleAddresses(), fill ? new(FillColor: color, ClearFillColor: color is null) : new(FontColor: color, ClearFontColor: color is null))))
            RememberCellColor(fill, color);
    }

    private async Task CustomCellColorAsync(bool fill)
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var addresses = VisibleAddresses();
        var color = await new CustomColorDialog(fill ? CurrentStyle.FillColor : CurrentStyle.FontColor, fill ? UiText.Get("自定义填充颜色") : UiText.Get("自定义字体颜色")).ShowAsync(this);
        if (color is not null && await ExecuteAsync(new SetCellStyleCommand(addresses, fill ? new(FillColor: color) : new(FontColor: color)), revision))
            RememberCellColor(fill, color);
    }
}
