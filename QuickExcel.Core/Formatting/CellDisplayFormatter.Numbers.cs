using System.Globalization;
using System.Text;

namespace QuickExcel.Core;

public static partial class CellDisplayFormatter
{
    private static string? FormatColor(string color) => color switch
    {
        "black" => "FF000000", "blue" => "FF0000FF", "cyan" => "FF00FFFF", "green" => "FF00FF00",
        "magenta" => "FFFF00FF", "red" => "FFFF0000", "white" => "FFFFFFFF", "yellow" => "FFFFFF00", _ => null
    };

    // A bounded, width-independent representation. '*' repetition is supplied by the view;
    // '_' and '?' reserve spaces in copied text. Unsupported calendars are not guessed.
    private static string? RenderNumber(double value, Section section, CultureInfo culture, out int fillOffset)
    {
        fillOffset = -1;
        var tokens = section.Tokens;
        if (tokens.Any(t => t.Kind == '/')) return RenderFraction(value, section, culture, out fillOffset);
        if (tokens.Count(t => t.Kind == '%') > 1) return null;
        if (tokens.Any(t => t.Kind == 'g'))
            return tokens.All(t => t.Kind is 'g' or 'l')
                ? string.Concat(tokens.Select(t => t.Kind == 'g' ? value.ToString("G15", culture) : t.Text)) : null;
        if (tokens.Any(t => t.Kind == '@')) return tokens.All(t => t.Kind is '@' or 'l' or '*') ? value.ToString("G15", culture) : null;
        var first = tokens.FindIndex(t => t.Kind is '0' or '#' or '?');
        if (first < 0)
        {
            if (tokens.Any(t => t.Kind is not ('l' or '*'))) return null;
            var text = new StringBuilder();
            foreach (var token in tokens)
                if (token.Kind == '*') fillOffset = text.Length;
                else text.Append(token.Text);
            return text.ToString();
        }
        if (first > 0 && tokens[first - 1].Kind == '.') first--;
        var last = tokens.FindLastIndex(t => t.Kind is '0' or '#' or '?' or '.' or ',');
        var core = tokens.Skip(first).Take(last - first + 1).ToArray();
        if (core.Any(t => t.Kind is not ('0' or '#' or '?' or '.' or ','))) return null;
        if (core.Count(t => t.Kind == '.') > 1 || tokens.Count(t => t.Kind == 'e') > 1) return null;
        var exponent = tokens.FirstOrDefault(t => t.Kind == 'e');
        if (exponent is not null && tokens.FindIndex(t => t.Kind == 'e') != last + 1) return null;
        var pattern = string.Concat(core.Select(t => t.Text));
        var scale = pattern.Length - pattern.TrimEnd(',').Length;
        pattern = pattern.TrimEnd(',');
        var parts = pattern.Split('.');
        var integerPattern = parts[0];
        var decimalPattern = parts.Length > 1 ? parts[1] : "";
        if (decimalPattern.Contains(',') || decimalPattern.Length > 30) return null;
        var grouping = integerPattern.Contains(',');
        integerPattern = integerPattern.Replace(",", "");
        if (integerPattern.Length > 30) return null;
        var magnitude = Math.Abs(value);
        magnitude *= Math.Pow(100, tokens.Count(t => t.Kind == '%')) / Math.Pow(1000, scale);
        if (!double.IsFinite(magnitude)) return null;
        var exponentValue = 0;
        if (exponent is not null && magnitude != 0)
        {
            // Work from a normalized decimal string to avoid Pow(10,-324) underflow.
            var scientific = magnitude.ToString("E16", Invariant).Split('E');
            var order = int.Parse(scientific[1], Invariant);
            var period = Math.Max(1, integerPattern.Length);
            exponentValue = (int)Math.Floor((double)order / period) * period;
            magnitude = double.Parse(scientific[0], Invariant) * Math.Pow(10, order - exponentValue);
        }
        // .NET custom numeric formatting uses decimal rounding for the decimal display,
        // avoiding a separate binary Math.Round step at values such as 2.675.
        var fixedText = magnitude.ToString("0" + (decimalPattern.Length == 0 ? "" : "." + new string('0', decimalPattern.Length)), Invariant);
        if (exponent is not null && fixedText.Split('.')[0].Length > Math.Max(1, integerPattern.Length))
        {
            var period = Math.Max(1, integerPattern.Length);
            exponentValue += period;
            magnitude /= Math.Pow(10, period);
            fixedText = magnitude.ToString("0" + (decimalPattern.Length == 0 ? "" : "." + new string('0', decimalPattern.Length)), Invariant);
        }
        var digits = fixedText.Split('.');
        var integer = digits[0];
        var required = exponent is not null && magnitude == 0 ? integerPattern.Length : integerPattern.Count(c => c == '0');
        if (integer == "0" && required == 0) integer = "";
        integer = integer.PadLeft(required, '0');
        if (grouping && integer.Length > 0)
        {
            var grouped = new StringBuilder();
            for (var i = 0; i < integer.Length; i++)
            {
                if (i > 0 && (integer.Length - i) % 3 == 0) grouped.Append(culture.NumberFormat.NumberGroupSeparator);
                grouped.Append(integer[i]);
            }
            integer = grouped.ToString();
        }
        var missing = Math.Max(0, integerPattern.Length - digits[0].Length);
        if (digits[0] == "0" && required == 0) missing = integerPattern.Length;
        integer = new string(' ', integerPattern.Take(missing).Count(c => c == '?')) + integer;
        var fraction = new StringBuilder();
        if (decimalPattern.Length > 0)
        {
            var lastRequired = Math.Max(decimalPattern.LastIndexOf('0'), digits[1].TrimEnd('0').Length - 1);
            for (var i = 0; i < decimalPattern.Length; i++)
                if (i <= lastRequired) fraction.Append(digits[1][i]);
                else if (decimalPattern[i] == '?') fraction.Append(' ');
        }
        var number = integer + (parts.Length > 1 ? culture.NumberFormat.NumberDecimalSeparator + fraction : "");
        if (exponent is not null)
            number += exponent.Text[0] + (exponentValue < 0 ? "-" : exponent.Text[1] == '+' ? "+" : "") +
                Math.Abs(exponentValue).ToString(new string('0', exponent.Text.Length - 2), Invariant);
        (string? Text, int Fill) Affix(IEnumerable<Token> affix)
        {
            var result = new StringBuilder();
            var fill = -1;
            foreach (var token in affix)
                if (token.Kind == 'l') result.Append(token.Text);
                else if (token.Kind == '*') fill = result.Length;
                else if (token.Kind == '%') result.Append(culture.NumberFormat.PercentSymbol);
                else if (token.Kind != 'e') return (null, -1);
            return (result.ToString(), fill);
        }
        var prefix = Affix(tokens.Take(first));
        var suffix = Affix(tokens.Skip(last + 1));
        if (prefix.Text is null || suffix.Text is null) return null;
        // Avoid a negative sign on a value that rounds to zero at this display precision.
        var negative = value < 0 && fixedText.Any(c => c is >= '1' and <= '9');
        var sign = negative ? culture.NumberFormat.NegativeSign : "";
        fillOffset = prefix.Fill >= 0 ? sign.Length + prefix.Fill : suffix.Fill >= 0 ? sign.Length + prefix.Text.Length + number.Length + suffix.Fill : -1;
        return sign + prefix.Text + number + suffix.Text;
    }
}
