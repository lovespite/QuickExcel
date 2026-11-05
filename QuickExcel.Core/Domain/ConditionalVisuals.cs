using QuickExcel.Core;
using System.Collections.Immutable;
using QuickExcel.Formulas;

namespace QuickExcel.Domain;

public enum ConditionalThresholdKind { Minimum, Maximum, Number, Percent, Percentile, Formula, AutoMinimum, AutoMaximum }
public enum DataBarAxis { Automatic, Middle, None }
public enum DataBarDirection { Context, LeftToRight, RightToLeft }

public sealed record ConditionalThreshold(ConditionalThresholdKind Kind, double Value = 0, string? Formula = null, bool Inclusive = true)
{
    internal void Validate()
    {
        if (!Enum.IsDefined(Kind) || !double.IsFinite(Value) || Kind is ConditionalThresholdKind.Percent or ConditionalThresholdKind.Percentile && Value is < 0 or > 100)
            throw new ArgumentException(MessageText.Get("条件格式阈值无效，百分比须为 0 到 100。"));
        if (Kind == ConditionalThresholdKind.Formula)
        {
            if (string.IsNullOrWhiteSpace(Formula)) throw new ArgumentException(MessageText.Get("条件格式阈值公式缺失。"));
            ConditionalFormatRule.ValidateFormula(Formula);
            if (FormulaSyntax.Walk(FormulaSyntax.Parse(Formula)).OfType<ReferenceNode>().Any(r => !r.ColumnAbsolute || !r.RowAbsolute || r.IsRange))
                throw new NotSupportedException(MessageText.Get("视觉规则阈值公式须使用绝对单元格引用，例如 $A$1。"));
        }
        else if (Formula is not null) throw new ArgumentException(MessageText.Get("只有公式阈值可以携带公式。"));
    }
    internal ConditionalThreshold Map(Func<string, string> transform) => Formula is null ? this : this with { Formula = transform(Formula) };
    internal static void ValidateOrder(IEnumerable<ConditionalThreshold> thresholds)
    {
        ConditionalThreshold? previous = null;
        foreach (var current in thresholds)
        {
            if (previous is not null && previous.Kind == current.Kind && current.Kind is ConditionalThresholdKind.Number or ConditionalThresholdKind.Percent or ConditionalThresholdKind.Percentile && previous.Value > current.Value)
                throw new ArgumentException(MessageText.Get("条件格式阈值须从小到大排列。"));
            previous = current;
        }
    }
}

public sealed record ConditionalColorScale
{
    public ImmutableArray<ConditionalThreshold> Thresholds { get; }
    public ImmutableArray<string> Colors { get; }
    public ConditionalColorScale(IEnumerable<ConditionalThreshold> thresholds, IEnumerable<string> colors)
    {
        Thresholds = thresholds.Take(4).ToImmutableArray(); Colors = colors.Take(4).ToImmutableArray();
        if (Thresholds.Length is < 2 or > 3 || Colors.Length != Thresholds.Length) throw new ArgumentException(MessageText.Get("色阶需要两组或三组阈值与颜色。"));
        foreach (var threshold in Thresholds)
        {
            threshold.Validate();
            if (threshold.Kind is ConditionalThresholdKind.AutoMinimum or ConditionalThresholdKind.AutoMaximum) throw new ArgumentException(MessageText.Get("色阶不支持自动数据条阈值。"));
        }
        if (Thresholds[0].Kind == ConditionalThresholdKind.Maximum || Thresholds[^1].Kind == ConditionalThresholdKind.Minimum ||
            Thresholds.Length == 3 && Thresholds[1].Kind is ConditionalThresholdKind.Minimum or ConditionalThresholdKind.Maximum)
            throw new ArgumentException(MessageText.Get("色阶阈值位置无效。"));
        foreach (var color in Colors) new DifferentialStyle(FillColor: color).Validate();
        ConditionalThreshold.ValidateOrder(Thresholds);
    }
    public bool ContentEquals(ConditionalColorScale other) => Thresholds.SequenceEqual(other.Thresholds) && Colors.SequenceEqual(other.Colors);
    internal ConditionalColorScale Map(Func<string, string> transform) => new(Thresholds.Select(t => t.Map(transform)), Colors);
}

public sealed record ConditionalDataBar(ConditionalThreshold Minimum, ConditionalThreshold Maximum, string Color,
    bool Gradient = true, bool ShowValue = true, int MinLength = 0, int MaxLength = 100,
    DataBarAxis Axis = DataBarAxis.Automatic, DataBarDirection Direction = DataBarDirection.Context,
    string NegativeColor = "FFFF0000", string AxisColor = "FF000000", string? BorderColor = null, string? NegativeBorderColor = null)
{
    internal void Validate()
    {
        Minimum.Validate(); Maximum.Validate();
        ConditionalThreshold.ValidateOrder([Minimum, Maximum]);
        if (Minimum.Kind is ConditionalThresholdKind.Maximum or ConditionalThresholdKind.AutoMaximum || Maximum.Kind is ConditionalThresholdKind.Minimum or ConditionalThresholdKind.AutoMinimum)
            throw new ArgumentException(MessageText.Get("数据条的最小/最大阈值位置无效。"));
        if (MinLength < 0 || MaxLength > 100 || MinLength > MaxLength || !Enum.IsDefined(Axis) || !Enum.IsDefined(Direction)) throw new ArgumentException(MessageText.Get("数据条长度、坐标轴或方向无效。"));
        foreach (var color in new[] { Color, NegativeColor, AxisColor, BorderColor, NegativeBorderColor }.OfType<string>()) new DifferentialStyle(FillColor: color).Validate();
        if (BorderColor is null && NegativeBorderColor is not null) throw new ArgumentException(MessageText.Get("负值边框必须同时设置正值边框。"));
    }
    internal ConditionalDataBar Map(Func<string, string> transform) => this with { Minimum = Minimum.Map(transform), Maximum = Maximum.Map(transform) };
    public bool ContentEquals(ConditionalDataBar other) => Canonical(this) == Canonical(other);
    private static ConditionalDataBar Canonical(ConditionalDataBar value) => value with
    { NegativeBorderColor = value.NegativeBorderColor ?? value.BorderColor, AxisColor = value.Axis == DataBarAxis.None ? "FF000000" : value.AxisColor };
}

public sealed record ConditionalIconSet
{
    public static ImmutableArray<string> Names { get; } = ["3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights1", "3TrafficLights2", "3Signs", "3Symbols", "3Symbols2",
        "4Arrows", "4ArrowsGray", "4RedToBlack", "4Rating", "4TrafficLights", "5Arrows", "5ArrowsGray", "5Rating", "5Quarters"];
    public string Name { get; }
    public ImmutableArray<ConditionalThreshold> Thresholds { get; }
    public bool ShowValue { get; }
    public bool Reverse { get; }
    public ConditionalIconSet(string name, IEnumerable<ConditionalThreshold> thresholds, bool showValue = true, bool reverse = false)
    {
        Name = name; Thresholds = thresholds.Take(6).ToImmutableArray(); ShowValue = showValue; Reverse = reverse;
        if (!Names.Contains(name) || Thresholds.Length != name[0] - '0') throw new ArgumentException(MessageText.Get("图标集名称或阈值数量无效。"));
        foreach (var threshold in Thresholds)
        {
            threshold.Validate();
            if (threshold.Kind is ConditionalThresholdKind.Minimum or ConditionalThresholdKind.Maximum or ConditionalThresholdKind.AutoMinimum or ConditionalThresholdKind.AutoMaximum) throw new ArgumentException(MessageText.Get("图标集需要数字、百分比、百分位或公式阈值。"));
        }
        ConditionalThreshold.ValidateOrder(Thresholds);
    }
    public bool ContentEquals(ConditionalIconSet other) => Name == other.Name && ShowValue == other.ShowValue && Reverse == other.Reverse && Thresholds.SequenceEqual(other.Thresholds);
    internal ConditionalIconSet Map(Func<string, string> transform) => new(Name, Thresholds.Select(t => t.Map(transform)), ShowValue, Reverse);
}

/// <summary>Normalized, immutable drawing data; it contains no renderer, calculator or UI types.</summary>
public sealed record ConditionalBarVisual(double Start, double End, double? Axis, string Color, bool Gradient, string? BorderColor, string AxisColor);
public sealed record ConditionalIconVisual(string Set, int Index);
public sealed record ConditionalCellVisual(ConditionalBarVisual? Bar = null, ConditionalIconVisual? Icon = null, bool HideValue = false);
