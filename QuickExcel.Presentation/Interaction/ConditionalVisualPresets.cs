using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public static class ConditionalVisualPresets
{
    public static ImmutableArray<(string Name, string Color)> BarColors { get; } =
        [(UiText.Get("蓝色"), "FF638EC6"), (UiText.Get("绿色"), "FF63C384"), (UiText.Get("红色"), "FFFF555A"), (UiText.Get("橙色"), "FFFFB628"), (UiText.Get("浅蓝色"), "FF00B7E9"), (UiText.Get("紫色"), "FFD174D3")];
    public static ImmutableArray<(string Name, ImmutableArray<string> Colors)> Scales { get; } =
    [
        (UiText.Get("绿–黄–红"), ["FFF8696B", "FFFFEB84", "FF63BE7B"]), (UiText.Get("红–黄–绿"), ["FF63BE7B", "FFFFEB84", "FFF8696B"]),
        (UiText.Get("绿–白–红"), ["FFF8696B", "FFFFFFFF", "FF63BE7B"]), (UiText.Get("红–白–绿"), ["FF63BE7B", "FFFFFFFF", "FFF8696B"]),
        (UiText.Get("蓝–白–红"), ["FFF8696B", "FFFFFFFF", "FF5A8AC6"]), (UiText.Get("红–白–蓝"), ["FF5A8AC6", "FFFFFFFF", "FFF8696B"]),
        (UiText.Get("白–红"), ["FFF8696B", "FFFFFFFF"]), (UiText.Get("红–白"), ["FFFFFFFF", "FFF8696B"]),
        (UiText.Get("绿–白"), ["FFFFFFFF", "FF63BE7B"]), (UiText.Get("白–绿"), ["FF63BE7B", "FFFFFFFF"]),
        (UiText.Get("绿–黄"), ["FFFFEB84", "FF63BE7B"]), (UiText.Get("黄–绿"), ["FF63BE7B", "FFFFEB84"])
    ];
    public static ImmutableArray<string> IconLabels { get; } = [UiText.Get("三向箭头（彩色）"), UiText.Get("三向箭头（灰色）"), UiText.Get("三面旗帜"), UiText.Get("三色交通灯"), UiText.Get("三色交通灯（带框）"), UiText.Get("三种形状"), UiText.Get("三种符号（带圈）"), UiText.Get("三种符号"),
        UiText.Get("四向箭头（彩色）"), UiText.Get("四向箭头（灰色）"), UiText.Get("红到黑"), UiText.Get("四档评级"), UiText.Get("四色交通灯"), UiText.Get("五向箭头（彩色）"), UiText.Get("五向箭头（灰色）"), UiText.Get("五档评级"), UiText.Get("五象限图")];

    public static ConditionalFormatRule DataBar(CellRectangle range, string color, bool gradient = true) => new([range], ConditionalFormatKind.DataBar, 1, new(),
        dataBar: new(new(ConditionalThresholdKind.AutoMinimum), new(ConditionalThresholdKind.AutoMaximum), color, Gradient: gradient));
    public static ConditionalFormatRule ColorScale(CellRectangle range, IEnumerable<string> colors)
    {
        var values = colors.ToArray();
        ConditionalThreshold[] thresholds = values.Length == 2 ? [new(ConditionalThresholdKind.Minimum), new(ConditionalThresholdKind.Maximum)] :
            [new(ConditionalThresholdKind.Minimum), new(ConditionalThresholdKind.Percentile, 50), new(ConditionalThresholdKind.Maximum)];
        return new([range], ConditionalFormatKind.ColorScale, 1, new(), colorScale: new(thresholds, values));
    }
    public static ConditionalFormatRule IconSet(CellRectangle range, string name) => new([range], ConditionalFormatKind.IconSet, 1, new(),
        iconSet: new(name, Enumerable.Range(0, name[0] - '0').Select(i => new ConditionalThreshold(ConditionalThresholdKind.Percent, Math.Round(i * 100d / (name[0] - '0'))))));
}
