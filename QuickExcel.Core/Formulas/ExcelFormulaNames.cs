using System.Text;

namespace QuickExcel.Formulas;

/// <summary>Only implemented extension functions have aliases; all other version prefixes remain opaque.</summary>
public static class ExcelFormulaNames
{
    private const string Prefix = "_xlfn.";
    private static readonly string[] Extensions = ["XLOOKUP", "XMATCH", "IFNA", "CONCAT", "TEXTJOIN"];
    private static bool Extension(string name) => Extensions.Contains(name, StringComparer.OrdinalIgnoreCase);
    public static string CanonicalName(string name) => name.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase) && Extension(name[Prefix.Length..])
        ? name[Prefix.Length..].ToUpperInvariant() : name.ToUpperInvariant();
    public static string Decode(string formula) => Rewrite(formula, false);
    public static string Encode(string formula) => Rewrite(formula, true);
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
            else if (char.IsLetterOrDigit(ch) || ch == '_')
            {
                var start = index++;
                while (index < formula.Length && (char.IsLetterOrDigit(formula[index]) || formula[index] is '_' or '.')) index++;
                var token = formula[start..index]; var next = index;
                while (next < formula.Length && char.IsWhiteSpace(formula[next])) next++;
                if (next < formula.Length && formula[next] == '(')
                {
                    if (encode && Extension(token)) token = Prefix + token;
                    else if (!encode && token.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase) && Extension(token[Prefix.Length..])) token = token[Prefix.Length..];
                }
                output.Append(token);
            }
            else { output.Append(ch); index++; }
        }
        return output.ToString();
    }
}
