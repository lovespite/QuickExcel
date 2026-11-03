using System.Text;

namespace QuickExcel.Formulas;

public static partial class ExcelFormulaNames
{
    // _xlpm. belongs to lexical LET parameters, not sheet names, table columns, quoted text or workbook names.
    // Retain original whitespace/parentheses/casing so a native roundtrip does not rewrite unrelated formula text.
    private readonly record struct NameToken(int Start, int Length, string Text, bool Identifier);
    private static string RewriteLocalParameters(string formula, bool encode)
    {
        if (!formula.Contains("LET", StringComparison.OrdinalIgnoreCase) || !encode && !formula.Contains("_xlpm.", StringComparison.OrdinalIgnoreCase)) return formula;
        try { _ = FormulaSyntax.Parse(formula); } catch (Exception e) when (e is FormatException or OverflowException) { return formula; }
        var tokens = new List<NameToken>();
        for (var i = 0; i < formula.Length;)
        {
            if (char.IsWhiteSpace(formula[i])) { i++; continue; }
            var start = i; var ch = formula[i++]; var identifier = false;
            if (ch is '\'' or '"')
            {
                while (i < formula.Length)
                {
                    if (formula[i++] != ch) continue;
                    if (i < formula.Length && formula[i] == ch) { i++; continue; }
                    break;
                }
            }
            else if (ch == '[')
            {
                i = start; _ = StructuredReferences.Token(formula, ref i);
            }
            else if (char.IsLetter(ch) || ch is '_' or '$' or '\\')
            {
                identifier = true;
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '_' or '$' or '.' or '\\')) i++;
            }
            tokens.Add(new(start, i - start, formula[start..i], identifier));
        }
        var closing = new Dictionary<int, int>(); var nesting = new Stack<int>();
        for (var i = 0; i < tokens.Count; i++)
        {
            if (tokens[i].Text == "(") nesting.Push(i);
            else if (tokens[i].Text == ")") closing[nesting.Pop()] = i;
        }
        var changes = new Dictionary<int, string>();
        void Segment(int start, int end, Dictionary<string, string> locals)
        {
            for (var i = start; i < end; i++)
            {
                var token = tokens[i];
                if (token.Identifier && i + 1 < end && tokens[i + 1].Text == "(")
                {
                    var last = closing[i + 1]; var ranges = new List<(int Start, int End)>(); var argument = i + 2;
                    for (var j = argument; j < last; j++)
                    {
                        if (tokens[j].Text == "(") j = closing[j];
                        else if (tokens[j].Text == ",") { ranges.Add((argument, j)); argument = j + 1; }
                    }
                    ranges.Add((argument, last));
                    if (CanonicalName(token.Text) == "LET" && StandardFunctions.AcceptsArguments("LET", ranges.Count))
                    {
                        var scope = new Dictionary<string, string>(locals, StringComparer.OrdinalIgnoreCase);
                        for (var j = 0; j < ranges.Count - 1; j += 2)
                        {
                            var declaration = ranges[j];
                            if (declaration.End != declaration.Start + 1 || !tokens[declaration.Start].Identifier) return;
                            Segment(ranges[j + 1].Start, ranges[j + 1].End, scope);
                            var name = tokens[declaration.Start].Text;
                            var renamed = encode ? name.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase) ? name : "_xlpm." + name
                                : name.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase) ? name[6..] : name;
                            scope[name] = renamed; changes[declaration.Start] = renamed;
                        }
                        Segment(ranges[^1].Start, ranges[^1].End, scope);
                    }
                    else foreach (var range in ranges) Segment(range.Start, range.End, locals);
                    i = last;
                }
                else if (token.Identifier && locals.TryGetValue(token.Text, out var name) &&
                    (i == 0 || tokens[i - 1].Text != "!") && (i + 1 == tokens.Count || tokens[i + 1].Text != "!" && !tokens[i + 1].Text.StartsWith('[')))
                    changes[i] = name;
            }
        }
        Segment(0, tokens.Count, new(StringComparer.OrdinalIgnoreCase));
        var result = new StringBuilder(formula);
        foreach (var (index, name) in changes.OrderByDescending(p => p.Key))
            result.Remove(tokens[index].Start, tokens[index].Length).Insert(tokens[index].Start, name);
        return result.ToString();
    }
}
