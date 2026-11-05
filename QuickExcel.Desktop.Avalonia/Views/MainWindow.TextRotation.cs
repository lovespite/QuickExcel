using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private Control TextRotationDropdown()
    {
        var menu = new MenuFlyout();
        var items = new List<Control>();
        var presets = new List<(MenuItem Item, int Rotation)>();
        foreach (var preset in TextRotationOptions.Presets)
        {
            var item = CommandMenu(UiText.Get(preset.Name), () => ApplyStyleAsync(new(TextRotation: preset.Rotation)));
            item.Name = "TextRotation" + preset.Rotation;
            item.ToggleType = MenuItemToggleType.CheckBox;
            item.Icon = TextRotationIcon(preset.Rotation);
            items.Add(item); presets.Add((item, preset.Rotation));
        }
        items.Add(new Separator());
        items.Add(CommandMenu(UiText.Get("设置单元格对齐方式…"), () => ShowNumberFormatAsync(true, alignmentTab: true)));
        menu.ItemsSource = items;
        menu.Opening += (_, _) =>
        {
            foreach (var (item, rotation) in presets) item.IsChecked = CurrentStyle.TextRotation == rotation;
        };
        var button = new DropDownButton
        {
            Name = "TextRotationDropdown", Content = TextRotationIcon(45), Flyout = menu,
            Height = 28, MinHeight = 28, Padding = new(6, 2), Margin = new(2, 1),
            Background = Brushes.Transparent, BorderThickness = new(0), HorizontalAlignment = HorizontalAlignment.Left
        };
        button.Classes.Add("ribbon-command");
        AutomationProperties.SetName(button, UiText.Get("文字方向")); ToolTip.SetTip(button, UiText.Get("文字方向"));
        _mutationControls.Add(button); return button;
    }

    private static Control TextRotationIcon(int rotation) => new Border
    {
        Width = 20, Height = 20,
        Child = new TextBlock
        {
            Text = rotation == 255 ? "a\nb" : "ab", FontSize = rotation == 255 ? 9 : 13,
            HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center,
            TextAlignment = TextAlignment.Center, RenderTransformOrigin = RelativePoint.Center,
            RenderTransform = new RotateTransform(-TextRotationOptions.ToDegrees(rotation))
        }
    };
}
