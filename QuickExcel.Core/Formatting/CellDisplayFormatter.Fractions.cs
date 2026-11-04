using System.Globalization;
using System.Text;

namespace QuickExcel.Core;

public static partial class CellDisplayFormatter
{
    private static string? RenderFraction(double value, Section section, CultureInfo culture, out int fillOffset)
    {
        fillOffset = -1;
        var tokens = section.Tokens;
        var slash = tokens.FindIndex(t => t.Kind == '/');
        if (tokens.Count(t => t.Kind == '/') != 1) return null;
        static bool Placeholder(Token t) => t.Kind is '#' or '?' or '0';
        var numeratorStart = slash;
        while (numeratorStart > 0 && Placeholder(tokens[numeratorStart - 1])) numeratorStart--;
        var denominatorEnd = slash + 1;
        while (denominatorEnd < tokens.Count && (Placeholder(tokens[denominatorEnd]) || tokens[denominatorEnd].Kind == 'v')) denominatorEnd++;
        var numeratorPattern = string.Concat(tokens.Skip(numeratorStart).Take(slash - numeratorStart).Select(t => t.Text));
        var denominatorPattern = string.Concat(tokens.Skip(slash + 1).Take(denominatorEnd - slash - 1).Select(t => t.Text));
        if (numeratorPattern.Length is < 1 or > 3 || denominatorPattern.Length is < 1 or > 4) return null;
        var fixedDenominator = denominatorPattern.All(char.IsAsciiDigit);
        int denominator;
        if (fixedDenominator)
        {
            if (!int.TryParse(denominatorPattern, out denominator) || denominator is < 1 or > 9999) return null;
        }
        else
        {
            if (denominatorPattern.Length > 3 || denominatorPattern.Any(c => c is not ('?' or '#' or '0'))) return null;
            denominator = (int)Math.Pow(10, denominatorPattern.Length) - 1;
        }
        var first = numeratorStart;
        var gapStart = first;
        while (gapStart > 0 && tokens[gapStart - 1] is { Kind: 'l' } t && t.Text.All(char.IsWhiteSpace)) gapStart--;
        var wholeStart = gapStart;
        while (wholeStart > 0 && Placeholder(tokens[wholeStart - 1])) wholeStart--;
        var mixed = gapStart < first && wholeStart < gapStart;
        var wholePattern = mixed ? string.Concat(tokens.Skip(wholeStart).Take(gapStart - wholeStart).Select(t => t.Text)) : "";
        var gap = mixed ? string.Concat(tokens.Skip(gapStart).Take(first - gapStart).Select(t => t.Text)) : "";
        if (mixed) first = wholeStart;
        var magnitude = Math.Abs(value);
        // Above this bound doubles cannot represent the fraction reliably.
        if (!double.IsFinite(magnitude) || magnitude > 1e12 || wholePattern.Length > 30) return null;
        var whole = mixed ? Math.Floor(magnitude) : 0;
        var remainder = magnitude - whole;
        long numerator;
        if (fixedDenominator) numerator = (long)Math.Round(remainder * denominator, MidpointRounding.AwayFromZero);
        else (numerator, denominator) = ClosestFraction(remainder, denominator);
        if (mixed && numerator >= denominator) { whole += numerator / denominator; numerator %= denominator; }
        static string Digits(long n, string pattern)
        {
            var digits = n.ToString(Invariant).PadLeft(pattern.Count(c => c == '0'), '0');
            return new string(' ', pattern.Take(Math.Max(0, pattern.Length - digits.Length)).Count(c => c == '?')) + digits;
        }
        var denominatorText = fixedDenominator ? denominatorPattern : denominator.ToString(Invariant).PadLeft(denominatorPattern.Count(c => c == '0'), '0');
        if (!fixedDenominator) denominatorText = denominatorText.PadRight(denominatorText.Length + denominatorPattern.Skip(denominatorText.Length).Count(c => c == '?'));
        var fraction = Digits(numerator, numeratorPattern) + "/" + denominatorText;
        if (mixed && numerator == 0 && !numeratorPattern.Contains('0')) fraction = new string(' ', fraction.Length);
        var integer = mixed ? Digits((long)whole, wholePattern) : "";
        if (mixed && whole == 0 && numerator != 0 && !wholePattern.Contains('0')) integer = new string(' ', wholePattern.Count(c => c == '?'));
        var number = integer + gap + fraction;
        var result = new StringBuilder(value < 0 && (whole != 0 || numerator != 0) ? culture.NumberFormat.NegativeSign : "");
        bool Affix(IEnumerable<Token> affix)
        {
            foreach (var t in affix)
                if (t.Kind == 'l') result.Append(t.Text);
                else if (t.Kind != '*') return false;
                else { /* Fill offset is resolved below without capturing the out parameter. */ }
            return true;
        }
        var prefix = tokens.Take(first).ToArray();
        var suffix = tokens.Skip(denominatorEnd).ToArray();
        if (!Affix(prefix)) return null;
        var prefixFill = Array.FindIndex(prefix, t => t.Kind == '*');
        if (prefixFill >= 0) fillOffset = result.Length - prefix.Skip(prefixFill + 1).Sum(t => t.Text.Length);
        result.Append(number);
        var suffixStart = result.Length;
        if (!Affix(suffix)) return null;
        var suffixFill = Array.FindIndex(suffix, t => t.Kind == '*');
        if (suffixFill >= 0) fillOffset = suffixStart + suffix.Take(suffixFill).Sum(t => t.Text.Length);
        return result.ToString();
    }

    // Bounded continued fractions avoid scanning every possible denominator for each visible cell.
    private static (long Numerator, int Denominator) ClosestFraction(double value, int maximum)
    {
        long p0 = 0, q0 = 1, p1 = 1, q1 = 0;
        var remainder = value;
        for (var i = 0; i < 64; i++)
        {
            var a = Math.Floor(remainder);
            if (q0 + a * q1 > maximum) break;
            var p2 = p0 + (long)a * p1; var q2 = q0 + (long)a * q1;
            (p0, q0, p1, q1) = (p1, q1, p2, q2);
            var fraction = remainder - a;
            if (fraction < 1e-14) return (p1, (int)q1);
            remainder = 1 / fraction;
        }
        var k = (maximum - q0) / q1;
        var p = p0 + k * p1; var q = q0 + k * q1;
        return Math.Abs((double)p1 / q1 - value) <= Math.Abs((double)p / q - value) ? (p1, (int)q1) : (p, (int)q);
    }
}
