using System.Text;
using QuickExcel.Domain;

namespace QuickExcel.Formulas;

[Flags]
public enum TableRows { Data = 1, Headers = 2, Totals = 4, All = 7, ThisRow = 8 }
public sealed record StructuredReferenceNode(string? Table, string? FirstColumn = null, string? LastColumn = null,
    TableRows Rows = TableRows.Data, string? Sheet = null) : FormulaNode;

/// <summary>Table reference syntax and snapshot binding. Names remain syntax; edits resolve identities before rewriting the AST.</summary>
public static class StructuredReferences
{
    public static string? CanonicalFormula(string? formula)
    {
        if (formula is null || !formula.Contains('[')) return formula;
        try { return "=" + FormulaSyntax.Print(FormulaSyntax.Parse(formula)); }
        catch (Exception ex) when (ex is FormatException or OverflowException) { return formula; }
    }
    public static string Token(string input, ref int offset)
    {
        var start = offset; var depth = 0;
        do
        {
            if (offset >= input.Length) throw new FormatException("结构化引用括号未闭合。");
            var c = input[offset++];
            if (c == '\'' && offset < input.Length && input[offset] is '[' or ']' or '#' or '\'' or '@') { offset++; continue; }
            if (c == '[') depth++; else if (c == ']') depth--;
            if (depth > 3) throw new FormatException("结构化引用嵌套过深。");
        } while (depth > 0);
        return input[start..offset];
    }

    public static StructuredReferenceNode Parse(string token, string? table = null, string? sheet = null, bool fromXlsx = false)
    {
        if (token == "[]" && table is not null) return new(table, Sheet: sheet);
        if (token.Length < 3 || token[0] != '[' || token[^1] != ']') throw new FormatException("结构化引用无效。");
        var body = token[1..^1]; var row = TableRows.Data; var special = false;
        var parts = new List<string>(); var separators = new List<char>();
        if (!fromXlsx && body.TrimStart().StartsWith('@')) { body = body.TrimStart()[1..]; row = TableRows.ThisRow; special = true; }
        if (body.Length == 0 && row == TableRows.ThisRow) return new(table, Rows: row, Sheet: sheet);
        if (body.TrimStart().StartsWith('['))
        {
            var pos = 0;
            while (pos < body.Length)
            {
                while (pos < body.Length && char.IsWhiteSpace(body[pos])) pos++;
                if (pos >= body.Length || body[pos] != '[') throw new FormatException("表列限定符缺少括号。");
                var leaf = Token(body, ref pos); parts.Add(leaf[1..^1]);
                while (pos < body.Length && char.IsWhiteSpace(body[pos])) pos++;
                if (pos == body.Length) break;
                var op = body[pos++]; if (op is not (',' or ':')) throw new FormatException("不支持的表引用运算符。"); separators.Add(op);
                if (pos == body.Length) throw new FormatException("表引用缺少后续限定符。");
            }
        }
        else parts.Add(body);
        string? first = null, last = null;
        for (var i = 0; i < parts.Count; i++)
        {
            var part = parts[i];
            if (part.StartsWith('#'))
            {
                if (first is not null || i > 0 && separators[i - 1] != ',') throw new FormatException("表行限定符必须在列名前。");
                var selected = part.ToUpperInvariant() switch { "#ALL" => TableRows.All, "#DATA" => TableRows.Data, "#HEADERS" => TableRows.Headers,
                    "#TOTALS" => TableRows.Totals, "#THIS ROW" => TableRows.ThisRow, _ => throw new FormatException("未知表行限定符。") };
                if (special && (row.HasFlag(selected) || row == TableRows.ThisRow || selected == TableRows.ThisRow)) throw new FormatException("重复或不兼容的表行限定符。");
                row = special ? row | selected : selected; special = true; continue;
            }
            var name = Decode(part, fromXlsx);
            if (name.Length is 0 or > 255) throw new FormatException("表列名称为空或过长。");
            if (first is null) { if (i > 0 && separators[i - 1] != ',') throw new FormatException("表行与列之间须用逗号。"); first = name; }
            else if (last is null && separators[i - 1] == ':') last = name;
            else throw new FormatException("不连续的表列引用尚未支持。");
        }
        if (row == (TableRows.Headers | TableRows.Totals)) throw new FormatException("不连续的表行引用尚未支持。");
        return new(table, first, last, row, sheet);
    }

    private static string Decode(string value, bool fromXlsx)
    {
        var result = new StringBuilder();
        for (var i = 0; i < value.Length; i++)
        {
            var c = value[i];
            if (c == '\'' && i + 1 < value.Length && value[i + 1] is '[' or ']' or '#' or '\'' or '@') result.Append(value[++i]);
            else if (c is '[' or ']' or '#' or '\'' || c == '@' && !fromXlsx) throw new FormatException("表列名称中的特殊字符须转义。");
            else result.Append(c);
        }
        return result.ToString();
    }
    public static string Print(StructuredReferenceNode reference, bool toXlsx = false)
    {
        string Column(string name) => "[" + string.Concat(name.Select(c => c is '[' or ']' or '#' or '\'' || c == '@' && !toXlsx ? "'" + c : c.ToString())) + "]";
        var parts = new List<string>();
        if (reference.Rows == TableRows.ThisRow) parts.Add("[#This Row]");
        else if (reference.Rows == TableRows.All) parts.Add("[#All]");
        else if (reference.Rows != TableRows.Data)
        {
            if (reference.Rows.HasFlag(TableRows.Headers)) parts.Add("[#Headers]");
            if (reference.Rows.HasFlag(TableRows.Data)) parts.Add("[#Data]");
            if (reference.Rows.HasFlag(TableRows.Totals)) parts.Add("[#Totals]");
        }
        if (reference.FirstColumn is { } first) parts.Add(Column(first) + (reference.LastColumn is { } last ? ":" + Column(last) : ""));
        if (parts.Count == 0) parts.Add("[#Data]");
        return FormulaSyntax.SheetPrefix(reference.Sheet) + reference.Table + "[" + string.Join(",", parts) + "]";
    }

    public static (WorksheetSnapshot Sheet, WorksheetTable Table)? Resolve(WorkbookSnapshot snapshot, SheetId context,
        CellPosition? position, StructuredReferenceNode reference)
    {
        var owner = reference.Sheet is null ? null : snapshot.Sheets.FirstOrDefault(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
        if (reference.Sheet is not null && owner is null) return null;
        foreach (var sheet in snapshot.Sheets)
        {
            if (owner is not null && owner.Id != sheet.Id) continue;
            foreach (var table in sheet.Data.Tables)
                if (reference.Table is { } name ? table.Name.Equals(name, StringComparison.OrdinalIgnoreCase) :
                    sheet.Id == context && position is { } p && table.Range.Contains(p.Row, p.Column)) return (sheet, table);
        }
        return null;
    }

    public static FormulaNode Bind(StructuredReferenceNode reference, WorkbookSnapshot snapshot, SheetId context, CellPosition? position)
    {
        FormulaNode Error(CellErrorCode code) => new ValueNode(CellValue.FromError(code));
        if (Resolve(snapshot, context, position, reference) is not { } resolved) throw new FormatException("结构化引用的表未建模或缺少当前单元格上下文。");
        var (sheet, table) = resolved; var range = table.Range;
        int Column(string name)
        { for (var i = 0; i < table.Columns.Length; i++) if (table.Columns[i].Name.Equals(name, StringComparison.OrdinalIgnoreCase)) return i; return -1; }
        var left = reference.FirstColumn is null ? 0 : Column(reference.FirstColumn);
        var right = reference.LastColumn is null ? reference.FirstColumn is null ? table.Columns.Length - 1 : left : Column(reference.LastColumn);
        if (left < 0 || right < left) return Error(CellErrorCode.Reference);
        var firstData = range.Top + (table.HeaderRow ? 1 : 0); var lastData = range.Bottom - (table.TotalsRow ? 1 : 0);
        int top, bottom;
        if (reference.Rows == TableRows.ThisRow)
        {
            if (position is not { } p || p.Row < firstData || p.Row > lastData) return Error(CellErrorCode.Value);
            top = bottom = p.Row;
        }
        else
        {
            var rows = reference.Rows;
            if (!table.HeaderRow) rows &= ~TableRows.Headers; if (!table.TotalsRow) rows &= ~TableRows.Totals;
            if (rows == 0) return Error(CellErrorCode.Reference);
            top = rows.HasFlag(TableRows.Headers) ? range.Top : rows.HasFlag(TableRows.Data) ? firstData : range.Bottom;
            bottom = rows.HasFlag(TableRows.Totals) ? range.Bottom : rows.HasFlag(TableRows.Data) ? lastData : range.Top;
            if (bottom < top) return Error(CellErrorCode.Reference);
        }
        return new ReferenceNode(sheet.Name, top, range.Left + left, bottom, range.Left + right, true, true, true, true,
            IsRange: top != bottom || left != right);
    }
}
