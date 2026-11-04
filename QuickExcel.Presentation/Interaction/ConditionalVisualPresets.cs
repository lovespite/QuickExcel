using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public static class ConditionalVisualPresets
{
    public static ImmutableArray<(string Name, string Color)> BarColors { get; } =
        [("蓝色", "FF638EC6"), ("绿色", "FF63C384"), ("红色", "FFFF555A"), ("橙色", "FFFFB628"), ("浅蓝色", "FF00B7E9"), ("紫色", "FFD174D3")];
    public static ImmutableArray<(string Name, ImmutableArray<string> Colors)> Scales { get; } =
    [
        ("绿–黄–红", ["FFF8696B", "FFFFEB84", "FF63BE7B"]), ("红–黄–绿", ["FF63BE7B", "FFFFEB84", "FFF8696B"]),
        ("绿–白–红", ["FFF8696B", "FFFFFFFF", "FF63BE7B"]), ("红–白–绿", ["FF63BE7B", "FFFFFFFF", "FFF8696B"]),
        ("蓝–白–红", ["FFF8696B", "FFFFFFFF", "FF5A8AC6"]), ("红–白–蓝", ["FF5A8AC6", "FFFFFFFF", "FFF8696B"]),
        ("白–红", ["FFF8696B", "FFFFFFFF"]), ("红–白", ["FFFFFFFF", "FFF8696B"]),
        ("绿–白", ["FFFFFFFF", "FF63BE7B"]), ("白–绿", ["FF63BE7B", "FFFFFFFF"]),
        ("绿–黄", ["FFFFEB84", "FF63BE7B"]), ("黄–绿", ["FF63BE7B", "FFFFEB84"])
    ];
    public static ImmutableArray<string> IconLabels { get; } = ["三向箭头（彩色）", "三向箭头（灰色）", "三面旗帜", "三色交通灯", "三色交通灯（带框）", "三种形状", "三种符号（带圈）", "三种符号",
        "四向箭头（彩色）", "四向箭头（灰色）", "红到黑", "四档评级", "四色交通灯", "五向箭头（彩色）", "五向箭头（灰色）", "五档评级", "五象限图"];

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
