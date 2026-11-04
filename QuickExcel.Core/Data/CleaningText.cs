using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using QuickExcel.Domain;

namespace QuickExcel.Core;

internal static class CleaningText
{
    internal static string[] Parse(string text, TextSplitOptions options)
    {
        if (!options.Breaks.IsDefault)
        {
            var breaks = options.Breaks;
            if (breaks.Length > 255 || breaks.Any(n => n <= 0 || n > 65534) || !breaks.SequenceEqual(breaks.Distinct().Order()))
                throw new ArgumentException("固定宽度分界须为递增的正整数，最多 255 个。");
            var ends = new List<(int Width, int Offset)>(); var width = 0; var offset = 0;
            foreach (var rune in text.EnumerateRunes())
            {
                var c = rune.Value;
                // Chinese Excel's fixed-width import counts full-width CJK characters as two units;
                // supplementary emoji count as one character, never half a surrogate pair.
                width += c is >= 0x1100 and <= 0x115f or >= 0x2e80 and <= 0xa4cf or >= 0xac00 and <= 0xd7a3 or
                    >= 0xf900 and <= 0xfaff or >= 0xfe10 and <= 0xfe19 or >= 0xfe30 and <= 0xfe6f or
                    >= 0xff01 and <= 0xff60 or >= 0xffe0 and <= 0xffe6 or >= 0x20000 and <= 0x3fffd ? 2 : 1;
                offset += rune.Utf16SequenceLength; ends.Add((width, offset));
            }
            var result = new List<string>(); var previous = 0; var index = 0;
            foreach (var boundary in breaks)
            {
                while (index < ends.Count && ends[index].Width < boundary) index++;
                var next = index < ends.Count ? ends[index].Offset : text.Length;
                result.Add(text[previous..next]); previous = next;
            }
            result.Add(text[previous..]); return result.ToArray();
        }
        if (string.IsNullOrEmpty(options.Delimiters) || options.Delimiters.Length > 16 || options.Delimiters.Any(char.IsSurrogate) ||
            options.Qualifier is not (null or '"' or '\'') || options.Qualifier is { } q && options.Delimiters.Contains(q))
            throw new ArgumentException("请选择 1–16 个分隔字符；文本限定符只能为单引号、双引号或无，且不能同时作为分隔符。");
        var fields = new List<string>(); var field = new StringBuilder(); var quoted = false; var closed = false;
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (quoted)
            {
                if (c != options.Qualifier) field.Append(c);
                else if (i + 1 < text.Length && text[i + 1] == c) { field.Append(c); i++; }
                else { quoted = false; closed = true; }
            }
            else if (options.Delimiters.Contains(c))
            {
                fields.Add(field.ToString()); field.Clear(); closed = false;
                if (fields.Count >= 256) throw new ArgumentException("分列超过 256 列。");
                if (options.CollapseDelimiters) while (i + 1 < text.Length && options.Delimiters.Contains(text[i + 1])) i++;
            }
            else if (c == options.Qualifier && field.Length == 0 && !closed) quoted = true;
            else
            {
                if (closed) throw new ArgumentException("文本限定符后存在非分隔字符，请先修正该行或关闭文本限定符。");
                field.Append(c);
            }
        }
        if (quoted) throw new ArgumentException("存在未闭合的文本限定符，整批分列未应用。");
        fields.Add(field.ToString()); return fields.ToArray();
    }
    internal static (CellValue Value, int FormatId, string? FormatCode) Convert(string text, CleaningColumnType type,
        NumericTextOptions options, ExcelDateSystem system, ICollection<string> notices)
    {
        if (text.Length == 0) return (CellValue.Blank, 0, null);
        if (type == CleaningColumnType.Text) return (CellValue.FromText(text), 49, null);
        if (type is CleaningColumnType.DateYmd or CleaningColumnType.DateMdy or CleaningColumnType.DateDmy)
        {
            var parts = text.Trim().Split(['-', '/', '.']); var yearIndex = type == CleaningColumnType.DateYmd ? 0 : 2;
            if (parts.Length != 3 || parts[yearIndex].Length != 4 || parts.Any(p => p.Length == 0 || p.Any(c => c is < '0' or > '9')))
                throw new ArgumentException($"日期“{text}”不符合所选年月日顺序；不猜测两位年份。");
            var values = parts.Select(p => int.Parse(p, CultureInfo.InvariantCulture)).ToArray(); var year = values[yearIndex];
            var month = values[type == CleaningColumnType.DateMdy ? 0 : 1]; var day = values[type == CleaningColumnType.DateYmd ? 2 : type == CleaningColumnType.DateMdy ? 1 : 0];
            double serial;
            if (system == ExcelDateSystem.Windows1900 && year == 1900 && month == 2 && day == 29) serial = 60;
            else
            {
                if (year < 1 || year > 9999 || month is < 1 or > 12 || day < 1 || day > DateTime.DaysInMonth(year, month))
                    throw new ArgumentException($"日期“{text}”无效，整批转换未应用。");
                serial = ExcelSerialDate.FromDateTime(new(year, month, day), system).Serial;
            }
            notices.Add("日期按明确的年月日顺序与工作簿日期系统解释，输出显示为 yyyy/m/d。");
            return (CellValue.FromNumber(serial), 0, "yyyy/m/d");
        }
        if (type != CleaningColumnType.Number) throw new ArgumentException("此输出类型不可转换。");
        if (options.DecimalSeparator is not ('.' or ',') || options.GroupSeparator == options.DecimalSeparator ||
            options.GroupSeparator is { } group && (char.IsDigit(group) || group is '+' or '-' or '%' or 'e' or 'E'))
            throw new ArgumentException("小数点须为点或逗号，千分位符须不同且不能是数字、符号或指数标记。");
        var value = text.Trim(); var percent = value.EndsWith('%'); if (percent) value = value[..^1];
        if (options.TrailingMinus && value.EndsWith('-')) value = "-" + value[..^1];
        var exponentIndex = value.IndexOfAny(['e', 'E']); var mantissa = exponentIndex < 0 ? value : value[..exponentIndex];
        var decimalIndex = mantissa.IndexOf(options.DecimalSeparator); var integer = decimalIndex < 0 ? mantissa : mantissa[..decimalIndex];
        var digitsAfter = decimalIndex < 0 ? 0 : mantissa.Length - decimalIndex - 1;
        if (options.GroupSeparator is { } separator && value.Contains(separator))
        {
            var unsigned = integer.TrimStart('+', '-'); var groups = unsigned.Split(separator);
            if (groups.Length < 2 || groups[0].Length is < 1 or > 3 || groups.Skip(1).Any(g => g.Length != 3) ||
                groups.Any(g => g.Any(c => c is < '0' or > '9')) || (decimalIndex >= 0 && mantissa[(decimalIndex + 1)..].Contains(separator)) ||
                (exponentIndex >= 0 && value[exponentIndex..].Contains(separator)))
                throw new ArgumentException($"数值“{text}”的千分位分组无效。");
        }
        var digits = new string(mantissa.Where(c => c is >= '0' and <= '9').ToArray()).TrimStart('0');
        if (digits.Length > 15) throw new ArgumentException($"文本“{text}”超过 Excel 的 15 位有效输入精度；请保留文本 ID，整批转换未应用。");
        if (Regex.IsMatch(integer.TrimStart('+', '-'), "^0[0-9]", RegexOptions.CultureInvariant))
            notices.Add("已明确选择数值转换：部分文本的前导零将被移除；需要保留编号时请改选文本。");
        var normalized = options.GroupSeparator is { } g ? value.Replace(g.ToString(), "") : value;
        normalized = normalized.Replace(options.DecimalSeparator, '.');
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number))
            throw new ArgumentException($"“{text}”不是所选分隔设置下的有限数值；日期、布尔值和其他文本不自动转换。");
        if (percent) number /= 100;
        if (number == 0 && digits.Any(c => c != '0') || number != 0 && Math.Abs(number) < 2.2250738585072014e-308)
            throw new ArgumentException($"数值“{text}”超出 Excel 可保留的最小数值范围；请保留文本。");
        if (percent) return (CellValue.FromNumber(number), digitsAfter == 0 ? 9 : 0, digitsAfter == 0 ? null : "0." + new string('0', Math.Min(15, digitsAfter)) + "%");
        if (exponentIndex >= 0) return (CellValue.FromNumber(number), 11, null);
        if (options.GroupSeparator is { } grouped && value.Contains(grouped))
            return (CellValue.FromNumber(number), digitsAfter == 0 ? 3 : digitsAfter == 2 ? 4 : 0,
                digitsAfter is 0 or 2 ? null : "#,##0." + new string('0', Math.Min(15, digitsAfter)));
        if (number == Math.Truncate(number) && Math.Abs(number) >= 1e11)
        {
            notices.Add("长整数使用明确的整数格式显示全部数字，避免 General 自动切换科学计数显示。");
            return (CellValue.FromNumber(number), 1, null);
        }
        return (CellValue.FromNumber(number), 0, null);
    }
}
