using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Presentation only: Excel date/time formats never change numeric inputs or calculation caches.</summary>
public static class CellDisplayFormatter
{
    private static readonly ConditionalWeakTable<StyleDefinition, ParsedFormat> Formats = new();
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

    public static bool IsDateFormat(StyleDefinition style, double serial)
    {
        var section = Formats.GetValue(style, s => Parse(s.NumberFormatCode ?? BuiltIn(s.NumberFormatId))).Select(serial);
        return section is { Supported: true } && section.Tokens.Any(t => t.Kind is 'y' or 'd');
    }

    public static string Format(CellValue value, StyleDefinition style, ExcelDateSystem dateSystem, CultureInfo? culture = null)
    {
        if (value.Kind != CellValueKind.Number) return CellInput.DisplayText(value);
        var numeric = style.NumberFormatCode ?? style.NumberFormatId switch
        { 1 => "0", 2 => "0.00", 3 => "#,##0", 4 => "#,##0.00", 9 => "0%", 10 => "0.00%", _ => null };
        if (numeric is "0" or "0.00" or "#,##0" or "#,##0.00" or "0%" or "0.00%")
            return value.NumberValue.ToString(numeric, culture ?? CultureInfo.CurrentCulture);
        var format = Formats.GetValue(style, s => Parse(s.NumberFormatCode ?? BuiltIn(s.NumberFormatId)));
        var section = format.Select(value.NumberValue);
        if (section is null || !section.Supported || !format.HasDate) return CellInput.DisplayText(value);
        if (!section.HasDate) return string.Concat(section.Tokens.Select(t => t.Text));
        return Render(value.NumberValue, dateSystem, section, culture ?? CultureInfo.CurrentCulture);
    }

    // Locale-dependent IDs use the application's simplified-Chinese convention.
    private static string? BuiltIn(int id) => id switch
    {
        14 => "mm-dd-yy", 15 => "d-mmm-yy", 16 => "d-mmm", 17 => "mmm-yy",
        18 => "h:mm AM/PM", 19 => "h:mm:ss AM/PM", 20 => "h:mm", 21 => "h:mm:ss", 22 => "m/d/yy h:mm",
        27 or 36 or 50 or 52 or 57 => "yyyy\"年\"m\"月\"",
        28 or 29 or 51 or 53 or 54 or 58 => "m\"月\"d\"日\"",
        30 => "m-d-yy", 31 => "yyyy\"年\"m\"月\"d\"日\"",
        32 => "h\"时\"mm\"分\"", 33 => "h\"时\"mm\"分\"ss\"秒\"",
        34 or 55 => "上午/下午h\"时\"mm\"分\"", 35 or 56 => "上午/下午h\"时\"mm\"分\"ss\"秒\"",
        45 => "mm:ss", 46 => "[h]:mm:ss", 47 => "mmss.0", _ => null
    };

    private sealed record Token(char Kind, string Text)
    {
        public int Length => Text.Length;
    }

    private sealed class Section
    {
        public List<Token> Tokens { get; } = [];
        public bool Supported { get; set; } = true;
        public bool HasDate => Tokens.Any(t => t.Kind is not ('l' or 'f'));
        public string? ConditionOperator { get; set; }
        public double ConditionValue { get; set; }
        public CultureInfo? Culture { get; set; }
        public bool Matches(double n) => ConditionOperator switch
        {
            "<" => n < ConditionValue, "<=" => n <= ConditionValue, ">" => n > ConditionValue,
            ">=" => n >= ConditionValue, "=" => n == ConditionValue, "<>" => n != ConditionValue, _ => true
        };
    }

    private sealed record ParsedFormat(Section[] Sections)
    {
        public bool HasDate => Sections.Any(s => s.HasDate && s.Supported);
        public Section? Select(double n)
        {
            var numeric = Sections.Take(3).ToArray();
            if (numeric.Length == 0) return null;
            if (numeric.Any(s => s.ConditionOperator is not null))
                return numeric.FirstOrDefault(s => s.ConditionOperator is not null && s.Matches(n))
                    ?? numeric.FirstOrDefault(s => s.ConditionOperator is null);
            return numeric[n < 0 && numeric.Length > 1 ? 1 : n == 0 && numeric.Length > 2 ? 2 : 0];
        }
    }

    private static ParsedFormat Parse(string? code)
    {
        if (code is null || code.Length > 1024) return new([]);
        var sections = new List<Section>();
        var start = 0;
        var quoted = false;
        for (var i = 0; i <= code.Length; i++)
        {
            if (i < code.Length)
            {
                if (code[i] is '\\' or '_' or '*' && !quoted) { i++; continue; }
                if (code[i] == '"') quoted = !quoted;
                if (code[i] != ';' || quoted) continue;
            }
            sections.Add(ParseSection(code[start..i]));
            if (sections.Count > 4) return new([]);
            start = i + 1;
        }
        return quoted ? new([]) : new(sections.ToArray());
    }

    private static Section ParseSection(string code)
    {
        var section = new Section();
        var tokens = section.Tokens;
        void Literal(string text) => tokens.Add(new('l', text));
        for (var i = 0; i < code.Length;)
        {
            var c = char.ToLowerInvariant(code[i]);
            if (c == '"')
            {
                var end = code.IndexOf('"', i + 1);
                if (end < 0) { section.Supported = false; break; }
                Literal(code[(i + 1)..end]); i = end + 1; continue;
            }
            if (c is '\\' or '_' or '*')
            {
                if (i + 1 >= code.Length) { section.Supported = false; break; }
                if (c == '\\') Literal(code[(i + 1)..(i + 2)]);
                else if (c == '_') Literal(" ");
                i += 2; continue;
            }
            if (c == '[')
            {
                var end = code.IndexOf(']', i + 1);
                if (end < 0) { section.Supported = false; break; }
                var bracket = code[(i + 1)..end];
                var lower = bracket.ToLowerInvariant();
                if (lower is "h" or "hh" or "m" or "mm" or "s" or "ss")
                    tokens.Add(new(char.ToUpperInvariant(lower[0]), lower));
                else if (lower.StartsWith("$-"))
                {
                    if (!int.TryParse(lower.AsSpan(2), NumberStyles.HexNumber, Invariant, out var lcid) || lcid > 0xffff)
                        section.Supported = false;
                    else if (lcid != 0)
                    {
                        try { section.Culture = CultureInfo.GetCultureInfo(lcid); }
                        catch (CultureNotFoundException) { section.Supported = false; }
                    }
                }
                else if (lower.Length > 0 && lower[0] is '<' or '>' or '=')
                {
                    var length = lower.Length > 1 && lower[1] is '=' or '>' ? 2 : 1;
                    var op = lower[..length];
                    if (section.ConditionOperator is not null || op is not ("<" or "<=" or ">" or ">=" or "=" or "<>") ||
                        !double.TryParse(lower.AsSpan(length), NumberStyles.Float, Invariant, out var limit) || !double.IsFinite(limit))
                        section.Supported = false;
                    else { section.ConditionOperator = op; section.ConditionValue = limit; }
                }
                else if (lower is not ("black" or "blue" or "cyan" or "green" or "magenta" or "red" or "white" or "yellow") &&
                    !(lower.StartsWith("color") && int.TryParse(lower.AsSpan(5), out var color) && color is >= 1 and <= 56))
                    section.Supported = false;
                i = end + 1; continue;
            }
            var marker = new[] { "AM/PM", "A/P", "上午/下午" }.FirstOrDefault(m => code.AsSpan(i).StartsWith(m, StringComparison.OrdinalIgnoreCase));
            if (marker is not null)
            {
                tokens.Add(new('a', code.Substring(i, marker.Length))); i += marker.Length; continue;
            }
            if (c is 'y' or 'm' or 'd' or 'h' or 's')
            {
                var end = i + 1;
                while (end < code.Length && char.ToLowerInvariant(code[end]) == c) end++;
                var length = end - i;
                if (length > (c == 'm' ? 5 : c is 'y' or 'd' ? 4 : 2)) section.Supported = false;
                tokens.Add(new(c, code[i..end])); i = end; continue;
            }
            if (c == '.' && i + 1 < code.Length && code[i + 1] == '0' && tokens.LastOrDefault()?.Kind is 's' or 'S')
            {
                var end = i + 1;
                while (end < code.Length && code[end] == '0') end++;
                if (end - i - 1 > 7) section.Supported = false;
                tokens.Add(new('f', code[(i + 1)..end])); i = end; continue;
            }
            // Reject numeric placeholders, unsupported calendar codes and general/scientific formats.
            if (char.IsAsciiLetterOrDigit(c) || c is '#' or '?' or '@' or '%' or '[' or ']')
                section.Supported = false;
            Literal(code[i..(i + 1)]); i++;
        }
        for (var i = 0; i < tokens.Count; i++)
        {
            if (tokens[i].Kind != 'm' || tokens[i].Length > 2) continue;
            var before = tokens.Take(i).LastOrDefault(t => t.Kind != 'l')?.Kind;
            var after = tokens.Skip(i + 1).FirstOrDefault(t => t.Kind != 'l')?.Kind;
            if (before is 'h' or 'H' || after is 's' or 'S') tokens[i] = tokens[i] with { Kind = 'n' };
        }
        return section;
    }

    private static string Render(double serial, ExcelDateSystem system, Section section, CultureInfo culture)
    {
        // Invalid Excel dates remain visible as overflow rather than being mistaken for ordinary numbers.
        if (serial < 0 || serial > 3652059) return "########";
        var tokens = section.Tokens;
        var hasTime = tokens.Any(t => t.Kind is 'h' or 'n' or 's' or 'H' or 'M' or 'S' or 'a');
        var digits = tokens.FirstOrDefault(t => t.Kind == 'f')?.Length ?? 0;
        var day = Math.Floor(serial);
        var ticks = (long)Math.Round((serial - day) * TimeSpan.TicksPerDay / Math.Pow(10, 7 - digits),
            MidpointRounding.AwayFromZero) * (long)Math.Pow(10, 7 - digits);
        if (hasTime && ticks == TimeSpan.TicksPerDay) { day++; ticks = 0; }
        var time = TimeSpan.FromTicks(ticks);
        var fictitious = system == ExcelDateSystem.Windows1900 && day == 60;
        var zeroDay = system == ExcelDateSystem.Windows1900 && day == 0;
        var date = new DateTime(1900, 2, 28);
        if (!fictitious && !new ExcelSerialDate(day, system).TryToDateTime(out date)) return "########";
        var year = zeroDay ? 1900 : date.Year;
        var month = zeroDay ? 1 : date.Month;
        var dateDay = fictitious ? 29 : zeroDay ? 0 : date.Day;
        var dayOfWeek = fictitious ? DayOfWeek.Wednesday : zeroDay ? DayOfWeek.Saturday : date.DayOfWeek;
        var names = (section.Culture ?? culture).DateTimeFormat;
        var twelveHour = tokens.Any(t => t.Kind == 'a');
        var builder = new StringBuilder();
        string Number(long n, int length) => n.ToString(length > 1 ? new string('0', length) : "0", Invariant);
        foreach (var token in tokens)
        {
            var length = token.Length;
            builder.Append(token.Kind switch
            {
                'l' => token.Text,
                'y' => Number(length <= 2 ? year % 100 : year, length <= 2 ? 2 : 4),
                'm' => length switch { 1 or 2 => Number(month, length), 3 => names.GetAbbreviatedMonthName(month),
                    4 => names.GetMonthName(month), _ => names.GetMonthName(month)[..1] },
                'd' => length switch { 1 or 2 => Number(dateDay, length), 3 => names.GetAbbreviatedDayName(dayOfWeek), _ => names.GetDayName(dayOfWeek) },
                'h' => Number(twelveHour ? (time.Hours + 11) % 12 + 1 : time.Hours, length),
                'n' => Number(time.Minutes, length),
                's' => Number(time.Seconds, length),
                'H' => Number((long)day * 24 + time.Hours, length),
                'M' => Number((long)day * 1440 + (long)time.TotalMinutes, length),
                'S' => Number((long)day * 86400 + (long)time.TotalSeconds, length),
                'f' => "." + (time.Ticks % TimeSpan.TicksPerSecond).ToString("D7", Invariant)[..length],
                'a' => token.Text.Split('/')[time.Hours < 12 ? 0 : 1],
                _ => ""
            });
        }
        return builder.ToString();
    }
}
