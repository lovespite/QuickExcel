using System.Text;
using QuickExcel.Formulas;

namespace QuickExcel.Storage;

/// <summary>OOXML uses #This Row; @ is a literal column character in the persisted table grammar.</summary>
internal static class XlsxStructuredFormula
{
    public static string Read(string formula) => Rewrite(formula, true);
    public static string Write(string formula) => ExcelFormulaNames.Encode(Rewrite(formula, false));
    private static string Rewrite(string formula, bool reading)
    {
        if (!formula.Contains('[')) return formula;
        var result = new StringBuilder(formula.Length); var pos = 0;
        try
        {
            while (pos < formula.Length)
            {
                var c = formula[pos];
                if (c is '\'' or '"')
                {
                    result.Append(c); pos++;
                    while (pos < formula.Length)
                    {
                        var next = formula[pos++]; result.Append(next); if (next != c) continue;
                        if (pos < formula.Length && formula[pos] == c) { result.Append(formula[pos++]); continue; } break;
                    }
                }
                else if (c == '[')
                {
                    var token = StructuredReferences.Token(formula, ref pos);
                    var reference = StructuredReferences.Parse(token, "_", fromXlsx: reading);
                    result.Append(StructuredReferences.Print(reference, toXlsx: !reading)[1..]);
                }
                else { result.Append(c); pos++; }
            }
            return result.ToString();
        }
        catch (FormatException) { return formula; } // Unsupported expressions retain their exact source and existing preservation policy.
    }
}
