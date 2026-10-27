using QuickExcel.Domain;

namespace QuickExcel.Formulas;

public sealed record DefinedName(string Name, string Formula, SheetId? Scope = null);

/// <summary>Binds sheet names and defined names without reading cached cell values.</summary>
public static class FormulaBinding
{
    public static FormulaNode Bind(FormulaNode expression, WorkbookSnapshot snapshot, SheetId currentSheet)
        => Bind(expression, snapshot, currentSheet, new HashSet<(SheetId?, string)>(), new BindingBudget(), 0);

    private sealed class BindingBudget { public int Nodes; }
    private static FormulaNode Bind(FormulaNode expression, WorkbookSnapshot snapshot, SheetId currentSheet,
        HashSet<(SheetId?, string)> visiting, BindingBudget budget, int depth)
    {
        if (depth > 16) throw new FormatException("名称引用嵌套过深。");
        return FormulaSyntax.Map(expression, node =>
        {
            if (++budget.Nodes > 16384) throw new FormatException("名称展开超过 16384 个语法节点。");
            if (node is ReferenceNode reference)
            {
                var sheet = reference.Sheet is null ? snapshot.GetSheet(currentSheet)
                    : snapshot.Sheets.FirstOrDefault(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
                return sheet is null ? new ValueNode(CellValue.FromError(CellErrorCode.Reference)) : reference with { Sheet = sheet.Name };
            }
            if (node is not NameNode name) return node;
            var context = name.Sheet is null ? snapshot.GetSheet(currentSheet)
                : snapshot.Sheets.FirstOrDefault(s => s.Name.Equals(name.Sheet, StringComparison.OrdinalIgnoreCase));
            if (context is null) return new ValueNode(CellValue.FromError(CellErrorCode.Reference));
            var definition = snapshot.DefinedNames.FirstOrDefault(n => n.Scope == context.Id && n.Name.Equals(name.Name, StringComparison.OrdinalIgnoreCase))
                ?? snapshot.DefinedNames.FirstOrDefault(n => n.Scope is null && n.Name.Equals(name.Name, StringComparison.OrdinalIgnoreCase));
            if (definition is null)
            {
                // Preserve the legacy bare-column spelling, after defined-name resolution.
                if (depth == 0 && name.Name.Length is >= 1 and <= 3 && name.Name.All(char.IsAsciiLetter))
                {
                    var column = 0;
                    foreach (var c in name.Name.ToUpperInvariant()) column = column * 26 + c - 'A' + 1;
                    if (column <= WorkbookLimits.MaxColumns)
                        return new ReferenceNode(context.Name, 0, column - 1, WorkbookLimits.MaxRows - 1, column - 1, IsRange: true, WholeColumns: true);
                }
                return new ValueNode(CellValue.FromError(CellErrorCode.Name));
            }
            var key = (definition.Scope, definition.Name.ToUpperInvariant());
            if (!visiting.Add(key)) throw new FormatException("名称之间存在循环引用。");
            try { return Bind(FormulaSyntax.Parse(definition.Formula), snapshot, context.Id, visiting, budget, depth + 1); }
            finally { visiting.Remove(key); }
        });
    }
}

public static class StandardFunctions
{
    public static IReadOnlyDictionary<string, (int Min, int Max)> Signatures { get; } =
        new Dictionary<string, (int, int)>(StringComparer.OrdinalIgnoreCase)
        {
            ["SUM"] = (1, 255), ["AVERAGE"] = (1, 255), ["MIN"] = (1, 255), ["MAX"] = (1, 255),
            ["COUNT"] = (1, 255), ["COUNTA"] = (1, 255), ["PRODUCT"] = (1, 255),
            ["ROUND"] = (2, 2), ["ABS"] = (1, 1), ["IF"] = (2, 3), ["IFERROR"] = (2, 2),
            ["AND"] = (1, 255), ["OR"] = (1, 255), ["NOT"] = (1, 1), ["CONCATENATE"] = (1, 255),
            ["LEFT"] = (1, 2), ["RIGHT"] = (1, 2), ["MID"] = (3, 3), ["LEN"] = (1, 1),
            ["TRIM"] = (1, 1), ["UPPER"] = (1, 1), ["LOWER"] = (1, 1),
            ["VLOOKUP"] = (3, 4), ["MATCH"] = (2, 3), ["INDEX"] = (2, 3),
            ["TODAY"] = (0, 0), ["NOW"] = (0, 0)
        };
}
