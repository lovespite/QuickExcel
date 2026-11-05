using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private int _borderLine = 1;
    private string? _borderColor;
    private CellBorderPatch _lastBorderPatch = CellBorders.Preset(BorderPreset.Bottom);
    private BorderPresetIcon? _borderIcon;
    private SplitButton BorderDropdown()
    {
        var items = new List<Control>();
        foreach (var (preset, name) in BorderEditor.Presets)
        {
            if (preset is BorderPreset.None or BorderPreset.DoubleBottom or BorderPreset.Inside) items.Add(new Separator());
            var item = CommandMenu(name, () => ApplyBordersAsync(CellBorders.Preset(preset, _borderLine, _borderColor)));
            item.Name = "BorderMenu" + preset;
            var patch = CellBorders.Preset(preset);
            item.Icon = new BorderPresetIcon(patch);
            items.Add(item);
        }
        items.Add(new Separator());
        var color = new MenuItem { Header = UiText.Get("线条颜色") };
        color.ItemsSource = ColorMenuItems(UiText.Get("自动（黑色）"), value => { _borderColor = value; return Task.CompletedTask; }, async () =>
        { if (await new CustomColorDialog(_borderColor, UiText.Get("边框颜色")).ShowAsync(this) is { } value) _borderColor = value; });
        items.Add(color);
        var line = new MenuItem { Header = UiText.Get("线型") };
        line.ItemsSource = BorderEditor.StyleNames.Select((name, index) =>
        {
            var item = CommandMenu(name, () => { _borderLine = index; return Task.CompletedTask; });
            item.Icon = new BorderLineSample(index) { Width = 58, Height = 18 }; return item;
        }).ToArray(); items.Add(line);
        items.Add(new Separator()); items.Add(CommandMenu(UiText.Get("其他边框…"), () => ShowNumberFormatAsync(true, true)));
        var menu = new MenuFlyout { ItemsSource = items };
        menu.Opening += (_, _) =>
        {
            color.Header = UiText.Get("线条颜色 · ") + (_borderColor is null ? UiText.Get("自动") : "#" + _borderColor[2..]);
            line.Header = UiText.Get("线型 · ") + BorderEditor.StyleNames[_borderLine];
        };
        _borderIcon = new BorderPresetIcon(_lastBorderPatch);
        var button = new SplitButton { Name = "BorderDropdown", Content = _borderIcon, Flyout = menu,
            Height = 28, MinHeight = 28, MinWidth = 0, Padding = new(4, 2), Margin = new(2, 1), Background = Brushes.Transparent, BorderThickness = new(0) };
        button.Classes.Add("ribbon-command");
        button.Click += async (_, _) => await PerformAsync(() => ApplyBordersAsync(_lastBorderPatch));
        AutomationProperties.SetName(button, UiText.Get("单元格边框")); ToolTip.SetTip(button, UiText.Get("边框 · 线型 · 颜色")); _mutationControls.Add(button); return button;
    }
    private async Task ApplyBordersAsync(CellBorderPatch patch)
    {
        if (await CommitEditorsAsync() && await ExecuteAsync(new SetCellBordersCommand(VisibleAddresses(), patch))) RememberBorders(patch);
    }
    private void RememberBorders(CellBorderPatch patch)
    {
        _lastBorderPatch = patch;
        if (_borderIcon is { } icon) icon.Patch = patch;
    }
    private BorderStroke?[] SelectionBorders(CellAddress[] addresses)
    {
        var set = addresses.ToHashSet(); var values = Enumerable.Range(0, 8).Select(_ => new HashSet<BorderStroke>()).ToArray();
        var sheet = Document.Snapshot.Sheets[SheetIndex]; var merges = sheet.MergedCells.Select(CellRectangle.Parse).ToArray();
        foreach (var a in addresses)
        {
            var s = Document.Snapshot.Styles[sheet.GetStyleId(a.Row, a.Column)];
            var merge = merges.FirstOrDefault(m => m.Contains(a.Row, a.Column));
            foreach (var (i, nr, nc, kind, color) in new[] {
                (0, a.Row, a.Column - 1, s.LeftBorder, s.LeftBorderColor), (1, a.Row, a.Column + 1, s.RightBorder, s.RightBorderColor),
                (2, a.Row - 1, a.Column, s.TopBorder, s.TopBorderColor), (3, a.Row + 1, a.Column, s.BottomBorder, s.BottomBorderColor) })
            {
                if (merge != default && merge.Contains(nr, nc)) continue;
                values[set.Contains(new(a.SheetId, nr, nc)) ? i < 2 ? 5 : 4 : i].Add(new(kind, color));
            }
            if (merge == default || a.Row == merge.Top && a.Column == merge.Left)
            { values[6].Add(new(s.DiagonalUp ? s.DiagonalBorder : 0, s.DiagonalBorderColor)); values[7].Add(new(s.DiagonalDown ? s.DiagonalBorder : 0, s.DiagonalBorderColor)); }
        }
        return values.Select(v => v.Count == 0 ? new(0) : v.Count == 1 ? v.Single() : (BorderStroke?)null).ToArray();
    }
}
