using System.Text.RegularExpressions;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public static partial class CellDisplayFormatter
{
    /// <summary>Change only numeric decimal placeholders; preserve section colors, conditions, literals, percent/scaling and exponents.</summary>
    public static bool TryAdjustDecimalPlaces(StyleDefinition style, CellValue value, int delta, out StyleDefinition adjusted)
    {
        if (delta is not (1 or -1)) throw new ArgumentOutOfRangeException(nameof(delta));
        adjusted = style;
        var code = style.NumberFormatCode ?? BuiltIn(style.NumberFormatId);
        if (code is null && style.NumberFormatId != 0) return false;
        if (code is null || code.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            var current = value.Kind == CellValueKind.Number ? value.NumberValue.ToString("G15", Invariant) : "0";
            var exponent = current.IndexOf('E'); var mantissa = exponent < 0 ? current : current[..exponent];
            var dot = mantissa.IndexOf('.'); var digits = dot < 0 ? 0 : mantissa.Length - dot - 1;
            code = "0" + (digits == 0 ? "" : "." + new string('0', digits)) + (exponent < 0 ? "" : "E+00");
        }
        var parsed = Parse(code);
        if (parsed.Sections.Length == 0 || parsed.Sections.Any(s => !s.Supported || s.HasDate || s.Tokens.Any(t => t.Kind is 'g' or '/'))) return false;
        var mask = code.ToCharArray();
        for (var i = 0; i < mask.Length; i++)
        {
            if (mask[i] == '"')
            {
                mask[i++] = '\0';
                while (i < mask.Length && mask[i] != '"') mask[i++] = '\0';
                if (i == mask.Length) return false;
                mask[i] = '\0';
            }
            else if (mask[i] is '\\' or '_' or '*') { mask[i] = '\0'; if (++i < mask.Length) mask[i] = '\0'; else return false; }
            else if (mask[i] == '[')
            {
                while (i < mask.Length && mask[i] != ']') mask[i++] = '\0';
                if (i == mask.Length) return false;
                mask[i] = '\0';
            }
        }
        var sections = new List<string>(); var start = 0; var found = false;
        for (var end = 0; end <= code.Length; end++)
        {
            if (end < code.Length && mask[end] != ';') continue;
            var source = code[start..end]; var visible = new string(mask, start, end - start); var sectionIndex = sections.Count;
            if (sectionIndex < 3 && parsed.Sections[sectionIndex].Tokens.Any(t => t.Kind is '0' or '#'))
            {
                if (RenderNumber(1.25, parsed.Sections[sectionIndex], Invariant, out _) is null) return false;
                var match = Regex.Match(visible, @"(?:[0#?][0#?,]*(?:\.[0#?]*)?|\.[0#?]+)", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(50));
                if (!match.Success) return false;
                var tail = visible[(match.Index + match.Length)..];
                tail = Regex.Replace(tail, @"^[eE][+-][0#?]+", "", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(50));
                if (tail.Any(c => c is '0' or '#' or '?' or '.')) return false;
                var core = match.Value.TrimEnd(','); var scaling = match.Value[core.Length..];
                var dot = core.IndexOf('.'); var digits = dot < 0 ? 0 : core.Length - dot - 1;
                if (digits > 30) return false;
                var next = Math.Clamp(digits + delta, 0, 30); var integer = dot < 0 ? core : core[..dot];
                if (integer.Length == 0 && next == 0) integer = "0";
                var decimals = dot < 0 ? "" : core[(dot + 1)..];
                decimals = next <= decimals.Length ? decimals[..next] : decimals + new string(decimals.Length == 0 ? '0' : decimals[^1], next - decimals.Length);
                var replacement = integer + (next == 0 ? "" : "." + decimals) + scaling;
                source = source[..match.Index] + replacement + source[(match.Index + match.Length)..]; found = true;
            }
            sections.Add(source); start = end + 1;
        }
        if (!found) return false;
        var result = string.Join(';', sections);
        if (result.Length > 1024) return false;
        if (result == (style.NumberFormatCode ?? BuiltIn(style.NumberFormatId))) return true;
        adjusted = style with { NumberFormatId = 0, NumberFormatCode = result }; return true;
    }
}
