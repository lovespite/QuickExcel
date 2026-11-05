using System.Text;

namespace QuickExcel.Formulas;

/// <summary>Only implemented extension functions have aliases; all other version prefixes remain opaque.</summary>
public static partial class ExcelFormulaNames
{
    private const string Prefix = "_xlfn.";
    private static readonly string[] Extensions = ["TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER", "TAKE", "DROP", "SORTBY", "XLOOKUP", "XMATCH", "IFNA", "IFS", "SWITCH", "CONCAT", "TEXTJOIN", "SEQUENCE", "FILTER", "SORT", "UNIQUE", "LET", "VAR.S", "VAR.P", "STDEV.S", "STDEV.P", "RANK.EQ"];
    private static bool Extension(string name) => Extensions.Contains(name, StringComparer.OrdinalIgnoreCase);
    public static string CanonicalName(string name)
    {
        if (name.Equals("_xlfn.ANCHORARRAY", StringComparison.OrdinalIgnoreCase)) return "ANCHORARRAY";
        if (name.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase))
        {
            var suffix = name[Prefix.Length..];
            if (suffix.StartsWith("_xlws.", StringComparison.OrdinalIgnoreCase) && suffix[6..].ToUpperInvariant() is "FILTER" or "SORT") suffix = suffix[6..];
            if (Extension(suffix)) return suffix.ToUpperInvariant();
        }
        return name.ToUpperInvariant();
    }
    public static string Decode(string formula)
    {
        var decoded = Rewrite(RewriteLocalParameters(formula, false), false);
        if (!decoded.Contains('#') && !decoded.Contains("ANCHORARRAY", StringComparison.OrdinalIgnoreCase)) return decoded;
        try
        {
            var parsed = FormulaSyntax.Parse(decoded);
            return FormulaSyntax.Walk(parsed).Any(n => n is UnaryNode { Operator: "#" }) ? "=" + FormulaSyntax.Print(parsed) : decoded;
        }
        catch (Exception error) when (error is FormatException or OverflowException) { return decoded; }
    }
    public static string Encode(string formula)
    {
        if (formula.Contains('#'))
        {
            try
            {
                var parsed = FormulaSyntax.Parse(formula);
                if (FormulaSyntax.Walk(parsed).Any(n => n is UnaryNode { Operator: "#" }))
                    formula = "=" + FormulaSyntax.Print(FormulaSyntax.Map(parsed, n => n is UnaryNode { Operator: "#" } u
                        ? new FunctionNode("_xlfn.ANCHORARRAY", [u.Operand]) : n));
            }
            catch (Exception error) when (error is FormatException or OverflowException) { }
        }
        return Rewrite(RewriteLocalParameters(formula, true), true);
    }
    private static string Rewrite(string formula, bool encode)
    {
        if (!encode && !formula.Contains(Prefix, StringComparison.OrdinalIgnoreCase) || encode && !Extensions.Any(n => formula.Contains(n, StringComparison.OrdinalIgnoreCase))) return formula;
        var output = new StringBuilder(formula.Length + 16); var index = 0;
        while (index < formula.Length)
        {
            var ch = formula[index];
            if (ch is '\'' or '"')
            {
                output.Append(ch); index++;
                while (index < formula.Length)
                {
                    var next = formula[index++]; output.Append(next);
                    if (next != ch) continue;
                    if (index < formula.Length && formula[index] == ch) { output.Append(formula[index++]); continue; }
                    break;
                }
            }
            else if (ch == '[')
            {
                var start = index;
                try { output.Append(StructuredReferences.Token(formula, ref index)); }
                catch (FormatException) { output.Append(formula[start..]); break; }
            }
            else if (char.IsLetterOrDigit(ch) || ch == '_')
            {
                var start = index++;
                while (index < formula.Length && (char.IsLetterOrDigit(formula[index]) || formula[index] is '_' or '.')) index++;
                var token = formula[start..index]; var next = index;
                while (next < formula.Length && char.IsWhiteSpace(formula[next])) next++;
                if (next < formula.Length && formula[next] == '(')
                {
                    if (encode && Extension(token)) token = Prefix + (token.ToUpperInvariant() is "FILTER" or "SORT" ? "_xlws." : "") + token;
                    else if (!encode && token.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase) && Extension(CanonicalName(token)))
                    {
                        token = token[Prefix.Length..];
                        if (token.StartsWith("_xlws.",StringComparison.OrdinalIgnoreCase)) token = token[6..];
                    }
                }
                output.Append(token);
            }
            else { output.Append(ch); index++; }
        }
        return output.ToString();
    }
}
