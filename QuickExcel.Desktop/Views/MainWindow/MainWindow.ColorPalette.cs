using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using QuickExcel.Core;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private sealed record PaletteColor(string Name, string Argb);
    private static readonly PaletteColor[] ColorPalette =
    [
        new("白色", "FFFFFFFF"), new("黑色", "FF000000"), new("深灰", "FF7F7F7F"), new("浅灰", "FFD9D9D9"),
        new("红色", "FFFF0000"), new("深红", "FFC00000"), new("橙色", "FFFF9900"), new("金色", "FFFFC000"),
        new("黄色", "FFFFFF00"), new("浅绿", "FF92D050"), new("绿色", "FF00B050"), new("青色", "FF00B0F0"),
        new("蓝色", "FF0070C0"), new("深蓝", "FF002060"), new("紫色", "FF7030A0"), new("粉色", "FFFF66CC"),
        new("浅黄", "FFFFF2CC"), new("浅橙", "FFFCE4D6"), new("浅红", "FFFFC7CE"), new("浅绿", "FFE2EFDA"),
        new("浅青", "FFDDEBF7"), new("浅蓝", "FFD9E1F2"), new("浅紫", "FFE4DFF1"), new("棕色", "FF843C0C")
    ];

    private void ColorMenu_Click(object sender, RoutedEventArgs e)
    {
        if (_busy || !CanCommitComposition() || _document is null || sender is not Button button) return;
        var kind = button.Tag as string;
        if (kind is not ("fill" or "font")) return;
        if (button.ContextMenu is { IsOpen: true } open) { open.IsOpen = false; return; }
        var menu = BuildColorMenu(kind);
        button.ContextMenu = menu; menu.PlacementTarget = button; menu.Placement = PlacementMode.Bottom; menu.IsOpen = true;
    }

    private ContextMenu BuildColorMenu(string kind)
    {
        var menu = new ContextMenu { Placement = PlacementMode.Bottom };
        var panel = new StackPanel { Margin = new(8), Width = 8 * 24 + 8 };
        panel.Children.Add(new TextBlock
        {
            Text = kind == "fill" ? "填充颜色" : "字体颜色", FontWeight = FontWeights.SemiBold, Margin = new(2, 0, 0, 6)
        });
        var swatches = new WrapPanel();
        foreach (var color in ColorPalette)
        {
            var brush = ColorBrush(color.Argb);
            var swatch = new Button
            {
                Width = 22, Height = 22, MinWidth = 0, MinHeight = 0, Padding = new(0), Margin = new(1),
                Background = brush, BorderBrush = Brushes.Gray, BorderThickness = new(1), ToolTip = color.Name
            };
            System.Windows.Automation.AutomationProperties.SetName(swatch, color.Name);
            var argb = color.Argb;
            swatch.Click += (_, _) => { menu.IsOpen = false; ApplyRibbonColor(kind, argb); };
            swatches.Children.Add(swatch);
        }
        panel.Children.Add(swatches);
        panel.Children.Add(new Border { Height = 1, Background = Brushes.Gainsboro, Margin = new(2, 6, 2, 6) });
        var clear = new Button { Content = kind == "fill" ? "无填充" : "自动颜色", MinHeight = 26, Padding = new Thickness(8, 3, 8, 3), Margin = new(0, 0, 0, 4) };
        clear.Click += (_, _) => { menu.IsOpen = false; ApplyRibbonColor(kind, null); };
        panel.Children.Add(clear);
        var custom = new Button { Content = "自定义颜色", MinHeight = 26, Padding = new Thickness(8, 3, 8, 3), Margin = new(0, 0, 0, 2) };
        custom.Click += (_, _) => { menu.IsOpen = false; ShowCustomColor(kind); };
        panel.Children.Add(custom);
        menu.Items.Add(panel);
        return menu;
    }

    private void ShowCustomColor(string kind)
    {
        if (_busy || !CommitEditors() || _document is null) return;
        var snapshot = _document.Snapshot; var style = snapshot.Styles[snapshot.Sheets[_sheetIndex].GetStyleId(_selectedRow, _selectedColumn)];
        var current = kind == "fill" ? style.FillColor : style.FontColor;
        var dialog = new ColorPickerDialog(this, kind == "fill" ? "自定义填充颜色" : "自定义字体颜色",
            current ?? (kind == "fill" ? "FFFFFFFF" : "FF000000"));
        if (dialog.ShowDialog() == true) ApplyRibbonColor(kind, dialog.SelectedColor);
    }

    private void ApplyRibbonColor(string kind, string? argb)
    {
        if (_busy || !CommitEditors()) return;
        try
        {
            var applied = ApplyRibbonStyle(kind == "fill"
                ? argb is null ? new CellStylePatch(ClearFillColor: true) : new CellStylePatch(FillColor: argb)
                : argb is null ? new CellStylePatch(ClearFontColor: true) : new CellStylePatch(FontColor: argb));
            if (!applied) return;
            StatusText.Text = kind == "fill"
                ? argb is null ? "已清除单元格填充色。" : "已应用填充颜色，可一次撤销。"
                : argb is null ? "已恢复自动字体颜色。" : "已应用字体颜色，可一次撤销。";
        }
        catch (Exception ex) { ReportEditFailure(ex); }
        finally { UpdateRibbonState(); }
    }

    private static Brush ColorBrush(string argb)
    {
        try { return (Brush)new BrushConverter().ConvertFromString("#" + argb)!; }
        catch (Exception ex) when (ex is FormatException or NotSupportedException) { return Brushes.White; }
    }
}
