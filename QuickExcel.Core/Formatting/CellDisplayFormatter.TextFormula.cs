using System.Globalization;
using System.Text;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public enum TextFormatSupport { Supported, Invalid, Unsupported }

public static partial class CellDisplayFormatter
{
    // TEXT's format argument uses Excel's simplified-Chinese format language in this product.
    // Number-format XML itself continues to use invariant English keywords.
    public static TextFormatSupport TextSupport(string code)
    {
        if (code.Length > 1024) return TextFormatSupport.Unsupported;
        var original = Parse(code);
        if (original.Sections.Any(s => s.Color is not null || s.Tokens.Any(t => t.Kind == 'g'))) return TextFormatSupport.Invalid;
        var local = LocalTextCode(code, out var invalid);
        if (invalid) return TextFormatSupport.Invalid;
        var format = Parse(local);
        if (format.Sections.Length == 0 || format.Sections.Any(s => !s.Supported)) return TextFormatSupport.Unsupported;
        for (var i = 0; i < format.Sections.Length; i++)
        {
            var section = format.Sections[i];
            if (i == 3) { if (section.Tokens.Any(t => t.Kind is not ('l' or '@' or '*'))) return TextFormatSupport.Unsupported; }
            else if (!section.HasDate && RenderNumber(1, section, CultureInfo.GetCultureInfo("en-US"), out _) is null) return TextFormatSupport.Unsupported;
        }
        return TextFormatSupport.Supported;
    }

    public static CellValue TextFormula(CellValue value, string code, ExcelDateSystem system)
    {
        if (value.Kind == CellValueKind.Error) return value;
        if (TextSupport(code) != TextFormatSupport.Supported) return CellValue.FromError(CellErrorCode.Value);
        var normalized = LocalTextCode(code, out _); var format = Parse(normalized); var culture = CultureInfo.GetCultureInfo("en-US");
        if (value.Kind == CellValueKind.Boolean) value = CellValue.FromText(value.BooleanValue ? "TRUE" : "FALSE");
        if (value.Kind == CellValueKind.Blank) value = CellValue.FromNumber(0);
        if (value.Kind == CellValueKind.Text)
            return CellValue.FromText(FormatDisplay(value, new() { NumberFormatCode = normalized }, system, culture).Text);
        var number = value.NumberValue; var section = format.Select(number);
        if (section is null) return CellValue.FromError(CellErrorCode.Value);
        var conditions = format.Sections.Take(3).Any(s => s.ConditionOperator is not null);
        var absolute = number < 0 && (conditions ? section.ConditionOperator is not null || Array.IndexOf(format.Sections, section) == 1 : format.Sections.Length > 1);
        var input = absolute ? Math.Abs(number) : number;
        string? text;
        if (section.HasDate)
        {
            if (number < 0 && system == ExcelDateSystem.Windows1900) return CellValue.FromError(CellErrorCode.Value);
            var sign = input < 0 ? "-" : ""; input = Math.Abs(input);
            if (!section.Tokens.Any(t => t.Kind == 'f')) input = Math.Floor(input * 86400 + .5) / 86400;
            text = Render(input, system, section, culture, out _);
            if (text == "########") return CellValue.FromError(CellErrorCode.Value);
            text = sign + text;
        }
        else
        {
            text = RenderNumber(input, section, culture, out _);
            if (input < 0 && section.Tokens.All(t => t.Kind is 'l' or '*')) text = "-" + text;
        }
        return text is null ? CellValue.FromError(CellErrorCode.Value) : CellValue.FromText(text);
    }

    private static string LocalTextCode(string code, out bool invalid)
    {
        invalid = false;
        var output = new StringBuilder(); var quoted = false;
        var colors = new Dictionary<string, string> { ["黑色"] = "Black", ["蓝色"] = "Blue", ["绿色"] = "Green",
            ["洋红"] = "Magenta", ["红色"] = "Red", ["白色"] = "White", ["黄色"] = "Yellow" };
        for (var i = 0; i < code.Length; i++)
        {
            if (code[i] == '"') quoted = !quoted;
            if (!quoted && code[i] is '\\' or '_' or '*' && i + 1 < code.Length) { output.Append(code[i++]); output.Append(code[i]); continue; }
            if (!quoted && code.AsSpan(i).StartsWith("G/通用格式", StringComparison.OrdinalIgnoreCase)) { output.Append("General"); i += 5; continue; }
            if (!quoted && code[i] == '[' && code.IndexOf(']', i + 1) is var end && end > i)
            {
                var token = code[(i + 1)..end];
                if (token is "青色" or "青绿" or "青绿色" or "紫红色") invalid = true;
                if (colors.TryGetValue(token, out var color)) { output.Append('[').Append(color).Append(']'); i = end; continue; }
                if (token.StartsWith("颜色", StringComparison.Ordinal) && int.TryParse(token.AsSpan(2), NumberStyles.None, CultureInfo.InvariantCulture, out var index) && index is >= 1 and <= 56)
                { i = end; continue; }
            }
            output.Append(code[i]);
        }
        return output.ToString();
    }
}
