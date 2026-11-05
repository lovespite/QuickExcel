using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly TextBlock _numberFormatLabel = new() { Text = UiText.Get("常规"), Width = 151, VerticalAlignment = VerticalAlignment.Center };
    private CellValue FormatSample => Document.Snapshot.Sheets[SheetIndex].GetCell(_grid.Current.Row, _grid.Current.Column)?.Value ?? CellValue.Blank;

    private void InstallNumberFormatRibbon()
    {
        var menu = new MenuFlyout();
        var dropdown = new DropDownButton { Name = "NumberFormatDropdown", Content = _numberFormatLabel, Width = 200, Height = 28, MinHeight = 28, Padding = new(8, 2), Flyout = menu };
        AutomationProperties.SetName(dropdown, UiText.Get("数字格式")); ToolTip.SetTip(dropdown, UiText.Get("选择数字格式")); _mutationControls.Add(dropdown);
        var previews = new List<(NumberFormatPreset Preset, TextBlock Text)>();
        var items = new List<Control>();
        foreach (var preset in NumberFormatOptions.Quick)
        {
            var item = CommandMenu(preset.Name, () => ApplyStyleAsync(preset.Patch));
            item.Icon = new TextBlock { Text = preset.Icon, FontSize = 18, Width = 32, TextAlignment = TextAlignment.Center };
            var text = new StackPanel { Spacing = 1, MinWidth = 180 };
            text.Children.Add(new TextBlock { Text = preset.Name, FontWeight = FontWeight.SemiBold });
            var preview = new TextBlock { MaxWidth = 250, TextTrimming = TextTrimming.CharacterEllipsis, Foreground = Brushes.Gray, FontSize = 12 };
            text.Children.Add(preview); previews.Add((preset, preview));
            item.Header = text; AutomationProperties.SetName(item, preset.Name); items.Add(item);
        }
        items.Add(new Separator()); items.Add(CommandMenu(UiText.Get("其他数字格式…"), NumberFormatAsync)); menu.ItemsSource = items;
        menu.Opening += (_, _) =>
        {
            var sample = FormatSample.Kind == CellValueKind.Blank ? CellValue.FromNumber(1234.56) : FormatSample;
            foreach (var (preset, text) in previews)
                text.Text = CellDisplayFormatter.Format(sample, new() { NumberFormatCode = preset.Code }, Document.Snapshot.DateSystem);
        };
        Button Symbol(string glyph, string label, Func<Task> action)
        {
            var button = new Button { Content = glyph, FontSize = 19, Width = 32, Height = 28, MinHeight = 28, Padding = new(2), Background = Brushes.Transparent, BorderThickness = new(0) };
            AutomationProperties.SetName(button, label); ToolTip.SetTip(button, label); _mutationControls.Add(button);
            button.Click += async (_, _) => await PerformAsync(action); return button;
        }
        var accounting = new MenuFlyout();
        accounting.ItemsSource = NumberFormatOptions.Symbols.Select(symbol => CommandMenu(symbol == NumberFormatOptions.NoSymbol ? UiText.Get("无货币符号") : symbol, () => ApplyStyleAsync(new(NumberFormatId: 0,
            NumberFormatCode: NumberFormatOptions.Numeric(NumberFormatCategory.Accounting, 2, true, symbol, 0))))).ToArray();
        var money = new DropDownButton { Content = "¥", Flyout = accounting, FontSize = 18, Height = 28, MinHeight = 28, Padding = new(4, 0), Background = Brushes.Transparent, BorderThickness = new(0) };
        AutomationProperties.SetName(money, UiText.Get("会计专用格式")); ToolTip.SetTip(money, UiText.Get("会计专用格式 · 选择货币符号")); _mutationControls.Add(money);
        var actions = new StackPanel { Orientation = Orientation.Horizontal };
        actions.Children.Add(money);
        actions.Children.Add(Symbol("%", UiText.Get("百分比样式"), () => ApplyStyleAsync(new(NumberFormatId: 9))));
        actions.Children.Add(Symbol(",", UiText.Get("千位分隔样式"), () => ApplyStyleAsync(new(NumberFormatId: 4))));
        actions.Children.Add(RibbonButton(RibbonIcon.IncreaseDecimals, UiText.Get("增加小数位"), () => PrecisionAsync(1)));
        actions.Children.Add(RibbonButton(RibbonIcon.DecreaseDecimals, UiText.Get("减少小数位"), () => PrecisionAsync(-1)));
        var more = new Button { Content = UiText.Get("其他数字格式…"), Height = 26, MinHeight = 26, Padding = new(6, 0), Background = Brushes.Transparent, BorderThickness = new(0), HorizontalAlignment = HorizontalAlignment.Right };
        AutomationProperties.SetName(more, UiText.Get("其他数字格式")); more.Click += async (_, _) => await PerformAsync(NumberFormatAsync); _mutationControls.Add(more);
        var body = new StackPanel { Spacing = 3 }; body.Children.Add(dropdown); body.Children.Add(actions); body.Children.Add(more);
        RibbonContentGroup(UiText.Get("开始"), UiText.Get("数字"), body);
    }

    private Task NumberFormatAsync() => ShowNumberFormatAsync(false);
    private Task CellFormatAsync() => ShowNumberFormatAsync(true);
    private async Task ShowNumberFormatAsync(bool appearance, bool borderTab = false, bool alignmentTab = false)
    {
        if (!await CommitEditorsAsync()) return;
        var revision = Document.Snapshot.Revision; var addresses = VisibleAddresses(); var session = Document.Session.SessionId;
        var dialog = new NumberFormatDialog(CurrentStyle, FormatSample, Document.Snapshot.DateSystem, Document.Snapshot.Styles, appearance, appearance ? SelectionBorders(addresses) : null, borderTab, alignmentTab);
        if (await dialog.ShowAsync(this) is not { } patch || patch == new CellStylePatch() && dialog.BorderPatch is null) return;
        if (session != Document.Session.SessionId) throw new InvalidOperationException(UiText.Get("工作簿已切换，请重新设置格式。"));
        await ExecuteAsync(new SetCellFormatCommand(addresses, patch, dialog.BorderPatch), revision);
    }
}
