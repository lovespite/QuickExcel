using System.Text;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private Range LookupRange(FormulaNode node)
    {
        if (node is ReferenceNode reference) return new(snapshot.Sheets.First(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase)), reference);
        if (node is ValueNode { Value.Kind: CellValueKind.Error } error) Fail(error.Value.ErrorCode, "查找引用无效。");
        throw new FormulaValueException(CellErrorCode.Value, "查找需要矩形引用。");
    }
    private CellValue LookupScalar(FormulaNode node) => node is ReferenceNode reference && ConditionalRanges.Shape(reference) == (1, 1)
        ? Read(LookupRange(reference).Sheet, reference.Row, reference.Column) : Scalar(node);
    private CellValue LookupValue(FunctionNode function)
    {
        var args = function.Arguments; var name = function.Name; var modern = name is "XLOOKUP" or "XMATCH";
        var needle = args[0] is MissingNode ? CellValue.Blank : LookupScalar(args[0]);
        if (needle.Kind == CellValueKind.Error) return needle;
        if (!modern && needle.Kind == CellValueKind.Text && needle.TextValue.Length > 255) return Error(CellErrorCode.Value);
        if (!modern && needle.Kind == CellValueKind.Blank) needle = CellValue.FromNumber(0);
        var lookup = LookupRange(args[1]); var shape = ConditionalRanges.Shape(lookup.Reference);
        if (name is "MATCH" or "XMATCH" or "XLOOKUP" && shape.Rows > 1 && shape.Columns > 1)
            return Error(name == "MATCH" ? CellErrorCode.NotAvailable : CellErrorCode.Value);
        var horizontal = name == "HLOOKUP" || name is "MATCH" or "XMATCH" or "XLOOKUP" && shape.Rows == 1;
        var length = horizontal ? shape.Columns : shape.Rows; var outputIndex = 0; Range? output = null;
        if (name is "VLOOKUP" or "HLOOKUP")
        {
            var indexValue = LookupScalar(args[2]); if (indexValue.Kind == CellValueKind.Error) return indexValue;
            var number = Math.Truncate(Number(indexValue));
            if (number < 1) return Error(CellErrorCode.Value);
            if (number > (horizontal ? shape.Rows : shape.Columns)) return Error(CellErrorCode.Reference);
            outputIndex = (int)number - 1;
        }
        else if (name == "XLOOKUP")
        {
            output = LookupRange(args[2]);
            if (ConditionalRanges.Shape(output.Reference) != shape) return Error(CellErrorCode.Value);
        }
        double Mode(int index, double fallback)
        {
            if (args.Count <= index || args[index] is MissingNode) return fallback;
            var value = LookupScalar(args[index]); if (value.Kind == CellValueKind.Error) Fail(value.ErrorCode, "查找模式包含错误。");
            return Number(value);
        }
        double LegacyMode()
        {
            if (args.Count < 4) return 1;
            var value = LookupScalar(args[3]); if (value.Kind == CellValueKind.Error) Fail(value.ErrorCode, "查找模式包含错误。");
            return Boolean(value) ? 1 : 0;
        }
        var mode = modern ? Math.Floor(Mode(name == "XLOOKUP" ? 4 : 2, 0)) : name == "MATCH" ? Math.Sign(Mode(2, 1)) : LegacyMode();
        var direction = modern ? Math.Floor(Mode(name == "XLOOKUP" ? 5 : 3, 1)) : 1;
        if (modern && (mode is not (-1 or 0 or 1 or 2) || direction is not (-1 or 1))) return Error(CellErrorCode.Value);
        var wildcard = needle.Kind == CellValueKind.Text && (modern ? mode == 2 : mode == 0) ? LookupPattern(needle.TextValue) : null;
        var best = -1; CellValue bestValue = default;
        foreach (var index in LookupPositions(lookup, horizontal, length, modern && needle.Kind == CellValueKind.Blank, (int)direction))
        {
            check(); var value = Read(lookup.Sheet, lookup.Reference.Row + (horizontal ? 0 : index), lookup.Reference.Column + (horizontal ? index : 0));
            if (value.Kind == CellValueKind.Error || value.Kind == CellValueKind.Blank && needle.Kind != CellValueKind.Blank) continue;
            var sameType = value.Kind == needle.Kind;
            var comparison = LookupCompare(value, needle);
            var exact = sameType && (wildcard is not null ? wildcard(value.TextValue) : comparison == 0);
            if (exact && (modern || mode == 0)) { best = index; break; }
            if (modern)
            {
                if (mode is not (-1 or 1) || (mode == -1 ? comparison > 0 : comparison < 0)) continue;
                var better = best < 0 ? 0 : LookupCompare(value, bestValue);
                if (best < 0 || (mode == -1 ? better > 0 : better < 0)) { best = index; bestValue = value; }
            }
            else
            {
                if (mode == 0 || !sameType || (mode > 0 ? comparison > 0 : comparison < 0)) continue;
                var better = best < 0 ? 0 : LookupCompare(value, bestValue);
                // Ascending legacy searches pick the last equal key; descending MATCH picks the first.
                if (best < 0 || (mode > 0 ? better >= 0 : better < 0)) { best = index; bestValue = value; }
            }
        }
        if (best < 0)
        {
            var fallback = name == "XLOOKUP" && args.Count > 3 && args[3] is not MissingNode ? LookupScalar(args[3]) : Error(CellErrorCode.NotAvailable);
            return fallback.Kind == CellValueKind.Blank ? CellValue.FromNumber(0) : fallback;
        }
        if (name is "MATCH" or "XMATCH") return CellValue.FromNumber(best + 1);
        var result = output is not null ? Read(output.Sheet, output.Reference.Row + (horizontal ? 0 : best), output.Reference.Column + (horizontal ? best : 0))
            : Read(lookup.Sheet, lookup.Reference.Row + (horizontal ? outputIndex : best), lookup.Reference.Column + (horizontal ? best : outputIndex));
        return result.Kind == CellValueKind.Blank ? CellValue.FromNumber(0) : result;
    }

    private IEnumerable<int> LookupPositions(Range range, bool horizontal, int length, bool includeBlank, int direction)
    {
        if (!_positions.TryGetValue(range.Sheet.Id, out var positions))
            _positions[range.Sheet.Id] = positions = range.Sheet.Cells.Keys.OrderBy(p => p.Row).ThenBy(p => p.Column).ToArray();
        var indices = new List<int>(); var r = range.Reference;
        foreach (var p in positions)
        {
            check(); if (horizontal ? p.Row != r.Row : p.Column != r.Column) continue;
            var index = horizontal ? p.Column - r.Column : p.Row - r.Row;
            if (index >= 0 && index < length) indices.Add(index);
        }
        indices.Sort();
        if (includeBlank)
        {
            var gap = direction == 1 ? 0 : length - 1;
            foreach (var index in direction == 1 ? indices : indices.AsEnumerable().Reverse())
            {
                check(); if (index != gap) break; gap += direction;
            }
            if (gap >= 0 && gap < length) { indices.Add(gap); indices.Sort(); }
        }
        return direction == 1 ? indices : indices.AsEnumerable().Reverse();
    }
    private static int LookupCompare(CellValue a, CellValue b) => a.Kind == CellValueKind.Blank || b.Kind == CellValueKind.Blank
        ? a.Kind == b.Kind ? 0 : a.Kind == CellValueKind.Blank ? 1 : -1
        : a.Kind == CellValueKind.Text && b.Kind == CellValueKind.Text
            ? string.Compare(a.TextValue, b.TextValue, StringComparison.InvariantCultureIgnoreCase) : Compare(a, b);
    private Func<string, bool> LookupPattern(string text)
    {
        var tokens = new List<(char Kind, char Value)>(); var literal = new StringBuilder(); var wildcard = false;
        for (var i = 0; i < text.Length; i++)
        {
            // Lookup's escape language consumes a trailing tilde, unlike conditional criteria.
            if (text[i] == '~') { if (i + 1 >= text.Length) continue; var escaped = text[++i]; tokens.Add(('c', escaped)); literal.Append(escaped); }
            else if (text[i] is '*' or '?') { tokens.Add((text[i], default)); wildcard = true; }
            else { tokens.Add(('c', text[i])); literal.Append(text[i]); }
        }
        if (!wildcard) { var exact = literal.ToString(); return value => string.Equals(exact, value, StringComparison.InvariantCultureIgnoreCase); }
        return value =>
        {
            var p = 0; var v = 0; var star = -1; var consumed = 0; var operations = 0;
            while (v < value.Length)
            {
                if ((operations++ & 255) == 0) check();
                if (p < tokens.Count && (tokens[p].Kind == '?' || tokens[p].Kind == 'c' && char.ToUpperInvariant(tokens[p].Value) == char.ToUpperInvariant(value[v]))) { p++; v++; }
                else if (p < tokens.Count && tokens[p].Kind == '*') { star = p++; consumed = v; }
                else if (star >= 0) { p = star + 1; v = ++consumed; }
                else return false;
            }
            while (p < tokens.Count && tokens[p].Kind == '*') p++;
            return p == tokens.Count;
        };
    }
}
