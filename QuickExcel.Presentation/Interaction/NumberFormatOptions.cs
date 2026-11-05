using System.Text.RegularExpressions;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public enum NumberFormatCategory { General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, Text, Special, Custom }
public sealed record NumberFormatSettings(NumberFormatCategory Category, int Decimals, bool Grouping, string Symbol, int Negative);

public sealed record NumberFormatPreset(string Name, string Code, NumberFormatCategory Category, string Icon)
{
    public CellStylePatch Patch => new(NumberFormatId: 0, NumberFormatCode: Code == "General" ? null : Code);
    public override string ToString() => Name;
}

/// <summary>Shared UI choices; all format changes still use the document's style command.</summary>
public static class NumberFormatOptions
{
    public static IReadOnlyList<string> Categories { get; } = Array.AsReadOnly(new[] { UiText.Get("常规"), UiText.Get("数值"), UiText.Get("货币"), UiText.Get("会计专用"), UiText.Get("日期"), UiText.Get("时间"), UiText.Get("百分比"), UiText.Get("分数"), UiText.Get("科学记数"), UiText.Get("文本"), UiText.Get("特殊"), UiText.Get("自定义") });
    public const string NoSymbol = "无";
    public static IReadOnlyList<string> Symbols { get; } = Array.AsReadOnly(new[] { "¥", "$", "€", "£", NoSymbol });
    public static IReadOnlyList<NumberFormatPreset> Quick { get; } = Array.AsReadOnly(new NumberFormatPreset[]
    {
        new(UiText.Get("常规"), "General", NumberFormatCategory.General, "123"),
        new(UiText.Get("数字"), "0.00", NumberFormatCategory.Number, "12"),
        new(UiText.Get("货币"), Numeric(NumberFormatCategory.Currency, 2, true, "¥", 0), NumberFormatCategory.Currency, "¥"),
        new(UiText.Get("会计专用"), Numeric(NumberFormatCategory.Accounting, 2, true, "¥", 0), NumberFormatCategory.Accounting, "¥ ="),
        new(UiText.Get("短日期"), "yyyy-mm-dd", NumberFormatCategory.Date, UiText.Get("日")),
        new(UiText.Get("长日期"), "yyyy\"年\"m\"月\"d\"日\"", NumberFormatCategory.Date, UiText.Get("日")),
        new(UiText.Get("时间"), "h:mm:ss", NumberFormatCategory.Time, "◷"),
        new(UiText.Get("百分比"), "0.00%", NumberFormatCategory.Percentage, "%"),
        new(UiText.Get("分数"), "# ?/?", NumberFormatCategory.Fraction, "½"),
        new(UiText.Get("科学记数"), "0.00E+00", NumberFormatCategory.Scientific, "10ⁿ"),
        new(UiText.Get("文本"), "@", NumberFormatCategory.Text, "ab")
    });

    public static string Code(StyleDefinition style) => style.NumberFormatCode ?? CellDisplayFormatter.BuiltIn(style.NumberFormatId) ?? (style.NumberFormatId == 0 ? "General" : "");
    public static string Label(StyleDefinition style)
    {
        var code = Code(style);
        return Quick.FirstOrDefault(p => p.Code == code)?.Name ?? Categories[(int)Category(code)];
    }
    public static NumberFormatCategory Category(string code)
    {
        var preset = Quick.FirstOrDefault(p => p.Code.Equals(code, StringComparison.OrdinalIgnoreCase));
        if (preset is not null) return preset.Category;
        foreach (var category in new[] { NumberFormatCategory.Date, NumberFormatCategory.Time, NumberFormatCategory.Fraction, NumberFormatCategory.Special })
            if (Types(category).Any(p => p.Code == code)) return category;
        if (ReadNumeric(code) is { } settings) return settings.Category;
        // Complex imported/custom codes stay intact until the user explicitly changes them.
        return NumberFormatCategory.Custom;
    }

    public static NumberFormatSettings? ReadNumeric(string code)
    {
        if (code.Length > 1024) return null;
        var decimals = Regex.Match(code, @"0\.(0+)", RegexOptions.CultureInvariant).Groups[1].Length;
        if (decimals > 30) return null;
        var grouping = code.Contains("#,##0", StringComparison.Ordinal);
        var symbol = Symbols.FirstOrDefault(s => s != NoSymbol && code.Contains("\"" + s + "\"", StringComparison.Ordinal)) ?? NoSymbol;
        foreach (var category in new[] { NumberFormatCategory.Number, NumberFormatCategory.Currency, NumberFormatCategory.Accounting, NumberFormatCategory.Percentage, NumberFormatCategory.Scientific })
            for (var negative = 0; negative < 4; negative++)
                if (Numeric(category, decimals, grouping, symbol, negative) == code) return new(category, decimals, grouping, symbol, negative);
        return null;
    }

    public static string Numeric(NumberFormatCategory category, int decimals, bool grouping, string symbol, int negative)
    {
        if (decimals is < 0 or > 30) throw new ArgumentOutOfRangeException(nameof(decimals));
        var number = (grouping ? "#,##0" : "0") + (decimals == 0 ? "" : "." + new string('0', decimals));
        var currency = symbol == NoSymbol ? "" : "\"" + symbol + "\"";
        if (category == NumberFormatCategory.Percentage) return number + "%";
        if (category == NumberFormatCategory.Scientific) return number + "E+00";
        if (category == NumberFormatCategory.Accounting)
            return $"_({currency}* {number}_);_({currency}* ({number});_({currency}* \"-\"{new string('?', decimals)}_);_(@_)";
        if (category == NumberFormatCategory.Currency) number = currency + number;
        return negative switch { 1 => number + ";[Red]-" + number, 2 => number + ";(" + number + ")", 3 => number + ";[Red](" + number + ")", _ => number };
    }

    public static NumberFormatPreset[] Types(NumberFormatCategory category)
    {
        (string Name, string Code)[] choices = category switch
        {
            NumberFormatCategory.General => [(UiText.Get("无特定格式"), "General")],
            NumberFormatCategory.Date => [("2001-03-14", "yyyy-mm-dd"), (UiText.Get("2001年3月14日"), "yyyy\"年\"m\"月\"d\"日\""), ("2001/3/14", "yyyy/m/d"), (UiText.Get("3月14日"), "m\"月\"d\"日\""), (UiText.Get("2001年3月"), "yyyy\"年\"m\"月\""), (UiText.Get("星期三"), "dddd"), (UiText.Get("日期和时间"), "yyyy-mm-dd hh:mm:ss")],
            NumberFormatCategory.Time => [("13:30:55", "h:mm:ss"), ("13:30", "h:mm"), ("1:30 PM", "h:mm AM/PM"), (UiText.Get("下午1时30分"), "上午/下午h\"时\"mm\"分\""), (UiText.Get("累计小时（37:30:55）"), "[h]:mm:ss"), (UiText.Get("分:秒"), "mm:ss")],
            NumberFormatCategory.Fraction => [(UiText.Get("一位分母（1/4）"), "# ?/?"), (UiText.Get("两位分母（21/25）"), "# ??/??"), (UiText.Get("三位分母（312/943）"), "# ???/???"), (UiText.Get("二分之一"), "# ?/2"), (UiText.Get("四分之一"), "# ?/4"), (UiText.Get("八分之一"), "# ?/8"), (UiText.Get("十六分之一"), "# ??/16"), (UiText.Get("十分之一"), "# ?/10"), (UiText.Get("百分之一"), "# ??/100")],
            NumberFormatCategory.Text => [(UiText.Get("文本（保留新输入的前导零）"), "@")],
            NumberFormatCategory.Special => [(UiText.Get("邮政编码（6位）"), "000000"), (UiText.Get("11位号码"), "00000000000")],
            _ => []
        };
        return choices.Select(p => new NumberFormatPreset(p.Name, p.Code, category, "")).ToArray();
    }

    public static IEnumerable<string> CustomCodes => new[] { "General", "0", "0.00", "#,##0", "#,##0.00", "#,##0;[Red](#,##0)", "0.00%", "0.00E+00", "# ?/?", "@" }
        .Concat(Quick.Select(p => p.Code)).Concat(new[] { 41, 42, 43, 44 }.Select(id => CellDisplayFormatter.BuiltIn(id)!)).Distinct();
}
