using QuickExcel.Domain;

namespace QuickExcel.Formulas;

public sealed record DefinedName(string Name, string Formula, SheetId? Scope = null);

/// <summary>Binds sheet names and defined names without reading cached cell values.</summary>
public static class FormulaBinding
{
    public static FormulaNode Bind(FormulaNode expression, WorkbookSnapshot snapshot, SheetId currentSheet, CellPosition? position = null)
    {
        if (snapshot.TextCompatibilityVersion > 2) throw new FormatException("未知文本计算兼容版本；不能确认公式语义。");
        var bound = Bind(expression, snapshot, currentSheet, new HashSet<(SheetId?, string)>(), new BindingBudget(), 0, false, position);
        ConditionalRanges.ValidateModes(bound);
        LookupModes.Validate(bound);
        DateMathFunctions.Validate(bound);
        TextFunctions.Validate(bound, snapshot);
        ScalarFormulaModes.Validate(bound);
        return bound;
    }

    /// <summary>Unqualified names in a name definition bind globally; cell expressions prefer their local scope.</summary>
    internal static DefinedName? ResolveName(WorkbookSnapshot snapshot, SheetId context, NameNode name, bool inDefinition = false)
    {
        var sheet = name.Sheet is null ? snapshot.GetSheet(context) : snapshot.Sheets.FirstOrDefault(s => s.Name.Equals(name.Sheet, StringComparison.OrdinalIgnoreCase));
        if (sheet is null) return null;
        return (inDefinition && name.Sheet is null ? null : snapshot.DefinedNames.FirstOrDefault(n => n.Scope == sheet.Id && n.Name.Equals(name.Name, StringComparison.OrdinalIgnoreCase)))
            ?? snapshot.DefinedNames.FirstOrDefault(n => n.Scope is null && n.Name.Equals(name.Name, StringComparison.OrdinalIgnoreCase));
    }

    private sealed class BindingBudget { public int Nodes; }
    private static FormulaNode Bind(FormulaNode expression, WorkbookSnapshot snapshot, SheetId currentSheet,
        HashSet<(SheetId?, string)> visiting, BindingBudget budget, int depth, bool inDefinition, CellPosition? position)
    {
        if (depth > 16) throw new FormatException("名称引用嵌套过深。");
        return FormulaSyntax.Map(expression, node =>
        {
            if (++budget.Nodes > 16384) throw new FormatException("名称展开超过 16384 个语法节点。");
            if (node is StructuredReferenceNode structured) return StructuredReferences.Bind(structured, snapshot, currentSheet, position);
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
            var definition = ResolveName(snapshot, context.Id, name, inDefinition);
            if (definition is null)
            {
                var table = new StructuredReferenceNode(name.Name, Sheet: name.Sheet);
                if (StructuredReferences.Resolve(snapshot, currentSheet, position, table) is not null)
                    return StructuredReferences.Bind(table, snapshot, currentSheet, position);
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
            try { return Bind(FormulaSyntax.Parse(definition.Formula), snapshot, definition.Scope ?? context.Id, visiting, budget, depth + 1, true, position); }
            finally { visiting.Remove(key); }
        });
    }
}

public static class StandardFunctions
{
    public static bool AcceptsArguments(string name, int count)
    {
        name = name.ToUpperInvariant();
        return Signatures.TryGetValue(name, out var bounds) && count >= bounds.Min && count <= bounds.Max &&
            (name is "SUMIFS" or "AVERAGEIFS" ? count % 2 == 1 : name != "COUNTIFS" || count % 2 == 0);
    }
    public static IReadOnlyDictionary<string, (int Min, int Max)> Signatures { get; } =
        new Dictionary<string, (int, int)>(StringComparer.OrdinalIgnoreCase)
        {
            ["SUM"] = (1, 255), ["AVERAGE"] = (1, 255), ["MIN"] = (1, 255), ["MAX"] = (1, 255),
            ["COUNT"] = (1, 255), ["COUNTA"] = (1, 255), ["PRODUCT"] = (1, 255),
            ["SUMIF"] = (2, 3), ["SUMIFS"] = (3, 255), ["COUNTIF"] = (2, 2), ["COUNTIFS"] = (2, 254),
            ["AVERAGEIF"] = (2, 3), ["AVERAGEIFS"] = (3, 255), ["COUNTBLANK"] = (1, 1),
            ["ROUND"] = (2, 2), ["ABS"] = (1, 1), ["IF"] = (2, 3), ["IFERROR"] = (2, 2),
            ["AND"] = (1, 255), ["OR"] = (1, 255), ["NOT"] = (1, 1), ["CONCATENATE"] = (1, 255),
            ["LEFT"] = (1, 2), ["RIGHT"] = (1, 2), ["MID"] = (3, 3), ["LEN"] = (1, 1),
            ["TRIM"] = (1, 1), ["UPPER"] = (1, 1), ["LOWER"] = (1, 1),
            ["VLOOKUP"] = (3, 4), ["MATCH"] = (2, 3), ["INDEX"] = (2, 3),
            ["HLOOKUP"] = (3, 4), ["XLOOKUP"] = (3, 6), ["XMATCH"] = (2, 4),
            ["DATE"] = (3, 3), ["YEAR"] = (1, 1), ["MONTH"] = (1, 1), ["DAY"] = (1, 1),
            ["TIME"] = (3, 3), ["HOUR"] = (1, 1), ["MINUTE"] = (1, 1), ["SECOND"] = (1, 1),
            ["EDATE"] = (2, 2), ["EOMONTH"] = (2, 2), ["WEEKDAY"] = (1, 2),
            ["WORKDAY"] = (2, 3), ["NETWORKDAYS"] = (2, 3),
            ["ROUNDUP"] = (2, 2), ["ROUNDDOWN"] = (2, 2), ["INT"] = (1, 1), ["MOD"] = (2, 2),
            ["TEXT"] = (2, 2), ["VALUE"] = (1, 1), ["SUBSTITUTE"] = (3, 4), ["REPLACE"] = (4, 4),
            ["FIND"] = (2, 3), ["SEARCH"] = (2, 3), ["TEXTJOIN"] = (3, 254), ["CONCAT"] = (1, 254),
            ["ISBLANK"] = (1, 1), ["ISNUMBER"] = (1, 1), ["ISTEXT"] = (1, 1), ["ISERROR"] = (1, 1), ["IFNA"] = (2, 2),
            ["TODAY"] = (0, 0), ["NOW"] = (0, 0)
        };
}
