using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private static string FormulaText(CellValue value) => value.Kind == CellValueKind.Number ? value.NumberValue.ToString("G15", CultureInfo.InvariantCulture) : Text(value);
    private int[] TextOffsets(string text, bool alwaysUnicode = false)
    {
        var offsets = new List<int>(text.Length + 1);
        for (var i = 0; i < text.Length; i++)
        {
            offsets.Add(i);
            if ((alwaysUnicode || snapshot.TextCompatibilityVersion >= 2) && char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) i++;
        }
        offsets.Add(text.Length); return offsets.ToArray();
    }
    private CellValue TextFunction(FunctionNode function)
    {
        var args = function.Arguments; var name = function.Name;
        CellValue Raw(int i) => LookupScalar(args[i]);
        CellValue Arg(int i) { var value = Raw(i); if (value.Kind == CellValueKind.Error) Fail(value.ErrorCode, "文本函数参数包含错误。"); return value; }
        string S(int i) => FormulaText(Arg(i));
        if (name == "IFNA")
        {
            static CellValue Result(CellValue value) => value.Kind == CellValueKind.Blank ? CellValue.FromNumber(0) : value;
            try { var value = Raw(0); return Result(value is { Kind: CellValueKind.Error, ErrorCode: CellErrorCode.NotAvailable } ? Raw(1) : value); }
            catch (FormulaValueException e) when (e.Code == CellErrorCode.NotAvailable) { return Result(Raw(1)); }
        }
        if (name is "ISBLANK" or "ISNUMBER" or "ISTEXT" or "ISERROR")
        {
            CellValue value;
            try { value = Raw(0); }
            catch (FormulaValueException e) { value = Error(e.Code); }
            catch (Exception e) when (e is FormatException or ArgumentException or OverflowException) { value = Error(CellErrorCode.Value); }
            return CellValue.FromBoolean(name switch { "ISBLANK" => value.Kind == CellValueKind.Blank, "ISNUMBER" => value.Kind == CellValueKind.Number,
                "ISTEXT" => value.Kind == CellValueKind.Text, _ => value.Kind == CellValueKind.Error });
        }
        if (name is "CONCAT" or "TEXTJOIN") return JoinText(function);
        if (name == "VALUE") return CellValue.FromNumber(ParseTextNumber(Arg(0)));
        if (name == "TEXT")
        {
            var value = Arg(0); var code = Arg(1);
            if (code.Kind == CellValueKind.Boolean) return Error(CellErrorCode.Value);
            if (value.Kind == CellValueKind.Text)
            {
                try { value = CellValue.FromNumber(ParseTextNumber(value)); }
                catch (FormulaValueException) { /* TEXT retains nonnumeric text and empty strings. */ }
            }
            return CellDisplayFormatter.TextFormula(value, FormulaText(code), snapshot.DateSystem);
        }
        if (name == "SUBSTITUTE")
        {
            var source = S(0); var old = S(1); var replacement = S(2);
            var occurrence = args.Count == 4 ? Math.Floor(Number(Arg(3))) : 0;
            if (args.Count == 4 && occurrence < 1) return Error(CellErrorCode.Value);
            if (old.Length == 0) return CellValue.FromText(source);
            var start = 0; var last = 0; var found = 0; var output = new StringBuilder();
            while ((start = source.IndexOf(old, start, StringComparison.Ordinal)) >= 0)
            {
                check(); found++;
                if (occurrence > 0 && found < occurrence) { start++; continue; }
                if ((long)output.Length + start - last + replacement.Length > WorkbookLimits.MaxCellTextLength) return Error(CellErrorCode.Value);
                output.Append(source.AsSpan(last, start - last)).Append(replacement); last = start + old.Length; start = last;
                if (occurrence > 0) break;
            }
            if ((long)output.Length + source.Length - last > WorkbookLimits.MaxCellTextLength) return Error(CellErrorCode.Value);
            return CellValue.FromText(output.Append(source.AsSpan(last)).ToString());
        }
        if (name == "REPLACE")
        {
            var source = S(0); var start = Math.Floor(Number(Arg(1))); var length = Math.Floor(Number(Arg(2))); var replacement = S(3);
            if (start < 1 || start > 32767 || length < 0 || length > 32767) return Error(CellErrorCode.Value);
            var offsets = TextOffsets(source); var from = (int)Math.Min(start - 1, offsets.Length - 1); var to = (int)Math.Min(from + length, offsets.Length - 1);
            if ((long)source.Length - (offsets[to] - offsets[from]) + replacement.Length > WorkbookLimits.MaxCellTextLength) return Error(CellErrorCode.Value);
            return CellValue.FromText(source[..offsets[from]] + replacement + source[offsets[to]..]);
        }
        if (name is "FIND" or "SEARCH")
        {
            var needle = S(0); var source = S(1); var start = args.Count == 3 ? Math.Floor(Number(Arg(2))) : 1; var offsets = TextOffsets(source);
            if (start < 1 || start > (name == "FIND" && needle.Length == 0 ? offsets.Length : offsets.Length - 1)) return Error(CellErrorCode.Value);
            if (name == "FIND")
            {
                var found = source.IndexOf(needle, offsets[(int)start - 1], StringComparison.Ordinal);
                var position = Array.BinarySearch(offsets, found);
                return found < 0 || position < 0 ? Error(CellErrorCode.Value) : CellValue.FromNumber(position + 1);
            }
            return SearchText(needle, source, offsets, (int)start - 1);
        }
        return Error(CellErrorCode.Name);
    }

    private double ParseTextNumber(CellValue value)
    {
        if (value.Kind is CellValueKind.Number or CellValueKind.Blank) return Number(value);
        if (value.Kind != CellValueKind.Text) throw new FormulaValueException(CellErrorCode.Value, "VALUE 不接受布尔值。");
        var text = new string(value.TextValue.Select(c => c is >= '０' and <= '９' ? (char)(c - '０' + '0') : c).ToArray()).Trim();
        var original = text; var sign = 1; var parenthesized = text.StartsWith('(') && text.EndsWith(')');
        if (parenthesized) { sign = -1; text = text[1..^1].Trim(); }
        var percent = text.StartsWith('%') || text.EndsWith('%');
        if (percent) text = (text.StartsWith('%') ? text[1..] : text[..^1]).Trim();
        var signed = text.StartsWith('+') || text.StartsWith('-');
        if (signed) { if (parenthesized) Fail(CellErrorCode.Value, "括号负数不能重复指定正负号。"); sign = text[0] == '-' ? -1 : 1; text = text[1..].Trim(); }
        if (text.StartsWith('$') || text.StartsWith('¥'))
        {
            if (percent) Fail(CellErrorCode.Value, "货币文本不能同时指定百分比。");
            text = text[1..].Trim();
            if (text.StartsWith('+') || text.StartsWith('-'))
            { if (parenthesized || signed) Fail(CellErrorCode.Value, "货币文本重复指定正负号。"); sign = text[0] == '-' ? -1 : 1; text = text[1..].Trim(); }
        }
        if (Regex.IsMatch(text, @"^(?:(?:[1-9]\d*(?:,\d{3,})+|\d+)(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100)) &&
            double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
            return sign * number / (percent ? 100 : 1);
        if (original == text && Regex.IsMatch(text, @"^\d{1,5}:\d{1,2}(?::\d{1,2}(?:\.\d{1,7})?)?$", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100)))
        {
            var parts = text.Split(':'); var h = double.Parse(parts[0], CultureInfo.InvariantCulture); var m = double.Parse(parts[1], CultureInfo.InvariantCulture);
            var s = parts.Length == 3 ? double.Parse(parts[2], CultureInfo.InvariantCulture) : 0;
            if (h <= 9999 && m < 60 && s < 60) return (h * 3600 + m * 60 + s) / 86400;
        }
        if (original == text && text.Length >= 8 && int.TryParse(text.AsSpan(0, 4), out var year) && year >= 1900)
        {
            try { var date = DateNumber(CellValue.FromText(text)); if (date >= 0 && date < LastDate + 1d) return date; }
            catch (FormulaValueException) { }
        }
        throw new FormulaValueException(CellErrorCode.Value, "VALUE 仅支持声明的数字、货币/百分比及年在前的日期/24 小时时间文本。");
    }

    private CellValue JoinText(FunctionNode function)
    {
        var joined = function.Name == "TEXTJOIN"; var first = joined ? 2 : 0;
        var delimiter = joined ? LookupScalar(function.Arguments[0]) : CellValue.FromText("");
        if (delimiter.Kind == CellValueKind.Error) return delimiter;
        var ignoreValue = joined ? LookupScalar(function.Arguments[1]) : CellValue.FromBoolean(true);
        if (ignoreValue.Kind == CellValueKind.Error) return ignoreValue;
        var ignore = Boolean(ignoreValue); var separator = FormulaText(delimiter); var output = new StringBuilder(); var seen = false;
        void Append(string value, long count = 1)
        {
            if (count == 0 || ignore && value.Length == 0) return;
            var delimiters = count - (seen ? 0 : 1); var size = output.Length + count * value.Length + delimiters * separator.Length;
            if (size > WorkbookLimits.MaxCellTextLength) Fail(CellErrorCode.Calculation, "连接结果超过 32767 字符。");
            if (value.Length != 0 || separator.Length != 0)
                for (long i = 0; i < count; i++) { check(); if (seen) output.Append(separator); output.Append(value); seen = true; }
            seen = true;
        }
        var values = new List<object>();
        for (var i = first; i < function.Arguments.Count; i++)
        {
            var argument = Evaluate(function.Arguments[i]);
            values.Add(argument);
            if (argument is Range cells)
            {
                foreach (var position in cells.Sheet.Cells.Keys.Where(p => cells.Reference.Contains(p.Row, p.Column)).OrderBy(p => p.Row).ThenBy(p => p.Column))
                { check(); var value = Read(cells.Sheet, position.Row, position.Column); if (value.Kind == CellValueKind.Error) return value; }
            }
            else if ((CellValue)argument is { Kind: CellValueKind.Error } error) return error;
        }
        foreach (var argument in values)
        {
            if (argument is Range range)
            {
                var shape = ConditionalRanges.Shape(range.Reference); long previous = -1;
                foreach (var position in range.Sheet.Cells.Keys.Where(p => range.Reference.Contains(p.Row, p.Column)).OrderBy(p => p.Row).ThenBy(p => p.Column))
                {
                    check(); var offset = (long)(position.Row - range.Reference.Row) * shape.Columns + position.Column - range.Reference.Column;
                    Append("", offset - previous - 1); previous = offset;
                    var value = Read(range.Sheet, position.Row, position.Column); if (value.Kind == CellValueKind.Error) return value; Append(FormulaText(value));
                }
                Append("", (long)shape.Rows * shape.Columns - previous - 1);
            }
            else { var value = (CellValue)argument; if (value.Kind == CellValueKind.Error) return value; Append(FormulaText(value)); }
        }
        return CellValue.FromText(output.ToString());
    }

    private CellValue SearchText(string pattern, string source, int[] offsets, int start)
    {
        var patternOffsets = TextOffsets(pattern); var tokens = new List<(char Kind, string Value)>();
        for (var i = 0; i < patternOffsets.Length - 1; i++)
        {
            var token = pattern[patternOffsets[i]..patternOffsets[i + 1]];
            if (token == "~") { if (++i < patternOffsets.Length - 1) tokens.Add(('c', pattern[patternOffsets[i]..patternOffsets[i + 1]])); }
            else tokens.Add((token is "*" or "?" ? token[0] : 'c', token));
        }
        for (var candidate = start; candidate < offsets.Length - 1; candidate++)
        {
            check(); var p = 0; var value = candidate; var star = -1; var consumed = candidate;
            while (p < tokens.Count)
            {
                check();
                if (tokens[p].Kind == '*') { star = p++; consumed = value; }
                else if (value < offsets.Length - 1 && (tokens[p].Kind == '?' || source.AsSpan(offsets[value], offsets[value + 1] - offsets[value]).Equals(tokens[p].Value, StringComparison.OrdinalIgnoreCase))) { value++; p++; }
                else if (star >= 0 && consumed < offsets.Length - 1) { p = star + 1; value = ++consumed; }
                else break;
            }
            if (p == tokens.Count) return CellValue.FromNumber(candidate + 1);
        }
        return Error(CellErrorCode.Value);
    }
}
