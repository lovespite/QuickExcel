using System.Text.RegularExpressions;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Presentation;

public readonly record struct FormulaTextEdit(int Start, int Length, string Text);
public sealed record FormulaCompletion(int Start, int Length, string[] Names);
public sealed record FormulaCallHint(string Name, int Argument, string Signature);
public sealed record FormulaReference(GridRange Range, string? Sheet);

/// <summary>Bounded lexical assistance for incomplete drafts; never evaluates or commits a formula.</summary>
public static class FormulaEditing
{
    private sealed record Token(int Start, int Length, string Kind, string Text);
    private static readonly Regex Reference = new(@"\G(?:(?:'(?:''|[^'])*'|[\p{L}_\\][\p{L}\p{N}_.\\]*)!)?(?:\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?:\s*:\s*\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?|\$?[A-Za-z]{1,3}\s*:\s*\$?[A-Za-z]{1,3})(?![\p{L}\p{N}_.$\\])", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
    private static bool Word(char c) => char.IsLetterOrDigit(c) || c is '_' or '.' or '$' or '\\';
    private static List<Token> Tokens(string text)
    {
        var result = new List<Token>();
        if (!text.StartsWith('=') || text.Length > 8192 || text.Contains('[') || text.Contains(']')) return result;
        for (var i = 1; i < text.Length;)
        {
            var start = i; var c = text[i];
            if (c == '"')
            {
                i++;
                while (i < text.Length) { if (text[i++] != '"') continue; if (i < text.Length && text[i] == '"') { i++; continue; } break; }
                result.Add(new(start, i - start, "text", text[start..i])); continue;
            }
            if (char.IsAsciiDigit(c) || c == '.' && i + 1 < text.Length && char.IsAsciiDigit(text[i + 1]))
            {
                i++; while (i < text.Length && (char.IsAsciiDigit(text[i]) || text[i] == '.')) i++;
                if (i < text.Length && text[i] is 'e' or 'E') { i++; if (i < text.Length && text[i] is '+' or '-') i++; while (i < text.Length && char.IsAsciiDigit(text[i])) i++; }
                result.Add(new(start, i - start, "number", text[start..i])); continue;
            }
            if (char.IsLetter(c) || c is '$' or '\'' or '_' or '\\')
            {
                var match = Reference.Match(text, i);
                if (match.Success)
                {
                    var after = i + match.Length; while (after < text.Length && char.IsWhiteSpace(text[after])) after++;
                    if (after == text.Length || text[after] is not ('(' or '!'))
                    {
                        try
                        {
                            if (FormulaSyntax.Parse("=" + match.Value) is ReferenceNode)
                            { i += match.Length; result.Add(new(start, match.Length, "reference", match.Value)); continue; }
                        }
                        catch (Exception ex) when (ex is FormatException or OverflowException or ArgumentException) { }
                    }
                }
                if (c == '\'')
                {
                    i++; while (i < text.Length) { if (text[i++] != '\'') continue; if (i < text.Length && text[i] == '\'') { i++; continue; } break; }
                    result.Add(new(start, i - start, "qualifier", text[start..i])); continue;
                }
                i++; while (i < text.Length && Word(text[i])) i++;
                result.Add(new(start, i - start, "word", text[start..i])); continue;
            }
            i++; if (!char.IsWhiteSpace(c)) result.Add(new(start, 1, c.ToString(), c.ToString()));
        }
        return result;
    }

    public static FormulaCompletion? Complete(string text, int caret)
    {
        var token = Tokens(text).LastOrDefault(t => t.Kind == "word" && t.Start < caret && t.Start + t.Length >= caret);
        if (token is null) return null;
        var prefix = text[token.Start..caret];
        var names = StandardFunctions.Signatures.Keys.Where(n => n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).Order(StringComparer.Ordinal).Take(8).ToArray();
        return names.Length == 0 ? null : new(token.Start, token.Length, names);
    }
    public static string Signature(string name)
    {
        var parameters = name.ToUpperInvariant() switch
        {
            "IF" => "logical_test, value_if_true, [value_if_false]", "IFERROR" => "value, value_if_error",
            "ROUND" => "number, num_digits", "LEFT" or "RIGHT" => "text, [num_chars]", "MID" => "text, start_num, num_chars",
            "VLOOKUP" => "lookup_value, table_array, col_index_num, [range_lookup]", "MATCH" => "lookup_value, lookup_array, [match_type]",
            "INDEX" => "array, row_num, [column_num]", "TODAY" or "NOW" => "", "ABS" => "number",
            "LEN" or "TRIM" or "UPPER" or "LOWER" => "text", "NOT" => "logical",
            "AND" or "OR" => "logical1, [logical2], …", "CONCATENATE" => "text1, [text2], …",
            "SUM" or "AVERAGE" or "MIN" or "MAX" or "COUNT" or "COUNTA" or "PRODUCT" => "number1, [number2], …",
            "SUMIF" => "range, criteria, [sum_range]",
            "AVERAGEIF" => "range, criteria, [average_range]",
            "COUNTIF" => "range, criteria",
            "COUNTIFS" => "criteria_range1, criteria1, [criteria_range2, criteria2], …",
            "SUMIFS" => "sum_range, criteria_range1, criteria1, …",
            "AVERAGEIFS" => "average_range, criteria_range1, criteria1, …",
            "COUNTBLANK" => "range",
            "HLOOKUP" => "lookup_value, table_array, row_index, [range_lookup]",
            "XLOOKUP" => "lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode: 1/-1]（模式须为常量）",
            "XMATCH" => "lookup_value, lookup_array, [match_mode], [search_mode: 1/-1]（模式须为常量）",
            "DATE" => "year, month, day", "TIME" => "hour, minute, second",
            "YEAR" or "MONTH" or "DAY" or "HOUR" or "MINUTE" or "SECOND" => "serial_number",
            "EDATE" or "EOMONTH" => "start_date, months", "WEEKDAY" => "serial_number, [return_type]",
            "WORKDAY" => "start_date, days, [holidays]", "NETWORKDAYS" => "start_date, end_date, [holidays]",
            "ROUNDUP" or "ROUNDDOWN" => "number, num_digits", "INT" => "number", "MOD" => "number, divisor",
            "TEXT" => "value, format_text", "VALUE" => "text", "SUBSTITUTE" => "text, old_text, new_text, [instance_num]",
            "REPLACE" => "old_text, start_num, num_chars, new_text", "FIND" or "SEARCH" => "find_text, within_text, [start_num]",
            "TEXTJOIN" => "delimiter, ignore_empty, text1, [text2], …", "CONCAT" => "text1, [text2], …",
            "ISBLANK" or "ISNUMBER" or "ISTEXT" or "ISERROR" => "value", "IFNA" => "value, value_if_na",
            _ => StandardFunctions.Signatures.TryGetValue(name, out var count) ? $"参数 {count.Min}–{count.Max}" : ""
        };
        return name + "(" + parameters + ")";
    }
    public static FormulaCallHint? Hint(string text, int caret)
    {
        var stack = new Stack<(string? Name, int Argument)>(); Token? previous = null;
        foreach (var token in Tokens(text).Where(t => t.Start < caret))
        {
            if (token.Kind == "(") stack.Push((previous?.Kind == "word" ? previous.Text.ToUpperInvariant() : null, 1));
            else if (token.Kind == ")" && stack.Count > 0) stack.Pop();
            else if (token.Kind == "," && stack.Count > 0) { var frame = stack.Pop(); stack.Push((frame.Name, frame.Argument + 1)); }
            previous = token;
        }
        foreach (var frame in stack)
            if (frame.Name is { } name && StandardFunctions.Signatures.ContainsKey(name)) return new(name, frame.Argument, Signature(name));
        return null;
    }
    public static FormulaTextEdit? ToggleReference(string text, int start, int length)
    {
        var tokens = Tokens(text).Where(t => t.Kind == "reference" && (length == 0 ? t.Start <= start && t.Start + t.Length >= start : t.Start < start + length && t.Start + t.Length > start)).ToArray();
        if (tokens.Length == 0) return null;
        var first = tokens[0].Start; var end = tokens[^1].Start + tokens[^1].Length; var replacement = text[first..end];
        foreach (var token in tokens.Reverse())
        {
            var reference = (ReferenceNode)FormulaSyntax.Parse("=" + token.Text);
            static (bool Column, bool Row) Cycle(bool col, bool row) => (col, row) switch { (false, false) => (true, true), (true, true) => (false, true), (false, true) => (true, false), _ => (false, false) };
            var a = Cycle(reference.ColumnAbsolute, reference.RowAbsolute); var b = Cycle(reference.EndColumnAbsolute, reference.EndRowAbsolute);
            if (reference.WholeColumns) { a = (!reference.ColumnAbsolute, false); b = (!reference.EndColumnAbsolute, false); }
            var updated = reference with { ColumnAbsolute = a.Column, RowAbsolute = a.Row, EndColumnAbsolute = b.Column, EndRowAbsolute = b.Row };
            // Preserve the exact sheet spelling and whitespace outside the reference body.
            var prefixLength = token.Text.LastIndexOf('!') + 1;
            var value = token.Text[..prefixLength] + FormulaSyntax.Print(updated with { Sheet = null });
            replacement = replacement.Remove(token.Start - first, token.Length).Insert(token.Start - first, value);
        }
        return new(first, end - first, replacement);
    }
    public static FormulaTextEdit? ReferenceSlot(string text, int start, int length)
    {
        if (!text.StartsWith('=') || start < 1 || start + length > text.Length || text.Length > 8192 || text.Contains('[') || text.Contains(']')) return null;
        var tokens = Tokens(text);
        var reference = tokens.FirstOrDefault(t => t.Kind == "reference" && t.Start <= start && t.Start + t.Length >= start + length);
        if (reference is not null) return new(reference.Start, reference.Length, "");
        if (tokens.Any(t => t.Kind is "text" or "qualifier" && t.Start < start && t.Start + t.Length >= start)) return null;
        if (length > 0) return null;
        var previous = start - 1; while (previous >= 0 && char.IsWhiteSpace(text[previous])) previous--;
        return previous >= 0 && "=+-*/^&(<>,:".Contains(text[previous]) ? new(start, 0, "") : null;
    }
    /// <summary>Resolves the reference token under/next to the caret for read-only cell highlighting.</summary>
    public static FormulaReference? ReferenceAt(string text, int caret)
    {
        if (!text.StartsWith('=') || text.Length > 8192 || text.Contains('[') || text.Contains(']')) return null;
        caret = Math.Clamp(caret, 0, text.Length);
        try
        {
            var token = Tokens(text).FirstOrDefault(t => t.Kind == "reference" && t.Start <= caret && caret <= t.Start + t.Length);
            if (token is null) return null;
            if (FormulaSyntax.Parse("=" + token.Text) is not ReferenceNode reference) return null;
            return new(new GridRange(reference.Row, reference.Column, reference.EndRow, reference.EndColumn), reference.Sheet);
        }
        catch (Exception ex) when (ex is FormatException or OverflowException or RegexMatchTimeoutException or ArgumentException) { return null; }
    }
    public static string Address(GridRange range) => FormulaSyntax.ColumnName(range.Left) + (range.Top + 1) +
        (range.Count == 1 ? "" : ":" + FormulaSyntax.ColumnName(range.Right) + (range.Bottom + 1));
}
