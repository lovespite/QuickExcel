using QuickExcel.Domain;

namespace QuickExcel.Formulas;

public sealed record DefinedName(string Name, string Formula, SheetId? Scope = null);

/// <summary>Binds sheet names and defined names without reading cached cell values.</summary>
public static class FormulaBinding
{
    public static FormulaNode Bind(FormulaNode expression, WorkbookSnapshot snapshot, SheetId currentSheet, CellPosition? position = null, bool arrayContext = false)
    {
        if (snapshot.TextCompatibilityVersion > 2) throw new FormatException("未知文本计算兼容版本；不能确认公式语义。");
        var bound = Bind(expression, snapshot, currentSheet, new HashSet<(SheetId?, string)>(), new BindingBudget(), 0, false, position);
        if (BoundDepth(bound) > 128) throw new FormatException("局部名称展开后的求值深度超过 128 层。");
        arrayContext |= DynamicArrayFunctions.UsesArrays(bound);
        ConditionalRanges.ValidateModes(bound);
        LookupModes.Validate(bound, arrayContext);
        DateMathFunctions.Validate(bound);
        OrderStatisticFunctions.Validate(bound);
        TextFunctions.Validate(bound, snapshot, arrayContext);
        CommonArrayModes.Validate(bound);
        if (!arrayContext && !DynamicArrayFunctions.UsesArrays(bound)) ScalarFormulaModes.Validate(bound);
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

    private sealed class BindingBudget { public int Nodes; public int LocalId; }
    private static int BoundDepth(FormulaNode node) => node switch
    {
        LocalValueNode local => local.EvaluationDepth,
        UnaryNode unary => 1 + BoundDepth(unary.Operand),
        BinaryNode binary => 1 + Math.Max(BoundDepth(binary.Left), BoundDepth(binary.Right)),
        FunctionNode function => 1 + function.Arguments.Select(BoundDepth).DefaultIfEmpty().Max(),
        LetNode let => 1 + Math.Max(BoundDepth(let.Body), let.Bindings.Select(b => BoundDepth(b.Value)).DefaultIfEmpty().Max()),
        _ => 1
    };
    private static FormulaNode Bind(FormulaNode expression, WorkbookSnapshot snapshot, SheetId currentSheet,
        HashSet<(SheetId?, string)> visiting, BindingBudget budget, int depth, bool inDefinition, CellPosition? position,
        IReadOnlyDictionary<string, LocalValueNode>? locals = null, int syntaxDepth = 0)
    {
        if (depth > 16) throw new FormatException("名称引用嵌套过深。");
        if (syntaxDepth > 128) throw new FormatException("绑定后公式嵌套超过 128 层。");
        FormulaNode Child(FormulaNode node, IReadOnlyDictionary<string, LocalValueNode>? scope = null) =>
            Bind(node, snapshot, currentSheet, visiting, budget, depth, inDefinition, position, scope ?? locals, syntaxDepth + 1);
        if (++budget.Nodes > 16384) throw new FormatException("名称展开超过 16384 个语法节点。");
        if (expression is FunctionNode { Arguments.Count: 0 } location && ReferenceFunctions.IsPosition(location.Name) && position is { } origin)
            return location with { Arguments = [new ReferenceNode(snapshot.GetSheet(currentSheet).Name, origin.Row, origin.Column, origin.Row, origin.Column)] };
        if (expression is FunctionNode { Name: "LET" } let)
        {
            if (!StandardFunctions.AcceptsArguments("LET", let.Arguments.Count)) throw new FormatException("LET 需要名称/值对和最后一个计算表达式。");
            var scope = locals is null ? new Dictionary<string, LocalValueNode>(StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, LocalValueNode>(locals, StringComparer.OrdinalIgnoreCase);
            var bindings = new List<LetBinding>(); var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < let.Arguments.Count - 1; i += 2)
            {
                if (let.Arguments[i] is not NameNode { Sheet: null } name || !ComplexFormulaFunctions.IsLocalName(name.Name) || !names.Add(name.Name))
                    throw new FormatException("LET 局部名称无效或重复。");
                var value = Child(let.Arguments[i + 1], scope); var id = ++budget.LocalId;
                bindings.Add(new(id, name.Name, value));
                var valueDepth = BoundDepth(value) + 1;
                if (valueDepth > 128) throw new FormatException("LET 局部名称引用链超过 128 层。");
                scope[name.Name] = new(id, name.Name, value, DateMathFunctions.ArrayResult(value), valueDepth);
            }
            return new LetNode(bindings, Child(let.Arguments[^1], scope));
        }
        var node = expression switch
        {
            UnaryNode unary => unary with { Operand = Child(unary.Operand) },
            BinaryNode binary => binary with { Left = Child(binary.Left), Right = Child(binary.Right) },
            FunctionNode function => function with { Arguments = function.Arguments.Select(a => Child(a)).ToArray() },
            _ => expression
        };
        {
            if (node is FunctionNode { Name: "OFFSET" or "INDIRECT" } indirect)
                return ComplexFormulaFunctions.BindReference(indirect, snapshot, currentSheet, position);
            if (node is StructuredReferenceNode structured) return StructuredReferences.Bind(structured, snapshot, currentSheet, position);
            if (node is ReferenceNode reference)
            {
                var sheet = reference.Sheet is null ? snapshot.GetSheet(currentSheet)
                    : snapshot.Sheets.FirstOrDefault(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
                return sheet is null ? new ValueNode(CellValue.FromError(CellErrorCode.Reference)) : reference with { Sheet = sheet.Name };
            }
            if (node is not NameNode name) return node;
            if (name.Sheet is null && locals?.TryGetValue(name.Name, out var local) == true)
                return local.Definition is ReferenceNode localReference ? localReference with { SourceExpression = new NameNode(local.Name) } : local;
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
        }
    }
}

public static class StandardFunctions
{
    public static bool AcceptsArguments(string name, int count)
    {
        name = name.ToUpperInvariant();
        return Signatures.TryGetValue(name, out var bounds) && count >= bounds.Min && count <= bounds.Max &&
            (name is "SUMIFS" or "AVERAGEIFS" or "LET" ? count % 2 == 1 : name is not ("COUNTIFS" or "IFS") || count % 2 == 0);
    }
    public static IReadOnlyDictionary<string, (int Min, int Max)> Signatures { get; } =
        new Dictionary<string, (int, int)>(StringComparer.OrdinalIgnoreCase)
        {
            ["SUM"] = (1, 255), ["AVERAGE"] = (1, 255), ["MIN"] = (1, 255), ["MAX"] = (1, 255),
            ["TEXTSPLIT"] = (2, 6), ["TAKE"] = (2, 3), ["DROP"] = (2, 3), ["SORTBY"] = (2, 255),
            ["SEQUENCE"] = (1, 4), ["FILTER"] = (2, 3), ["SORT"] = (1, 4), ["UNIQUE"] = (1, 3),
            ["LET"] = (3, 253), ["OFFSET"] = (3, 5), ["INDIRECT"] = (1, 2),
            ["ROW"] = (0, 1), ["COLUMN"] = (0, 1), ["ROWS"] = (1, 1), ["COLUMNS"] = (1, 1),
            ["IFS"] = (2, 254), ["SWITCH"] = (3, 254), ["SUMPRODUCT"] = (1, 255),
            ["VAR.S"] = (1, 254), ["VAR.P"] = (1, 254), ["STDEV.S"] = (1, 254), ["STDEV.P"] = (1, 254), ["MEDIAN"] = (1, 255),
            ["LARGE"] = (2, 2), ["SMALL"] = (2, 2), ["RANK.EQ"] = (2, 3),
            ["COUNT"] = (1, 255), ["COUNTA"] = (1, 255), ["PRODUCT"] = (1, 255),
            ["SUBTOTAL"] = (2, 255),
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
            ["TEXTBEFORE"] = (2, 6), ["TEXTAFTER"] = (2, 6), ["TEXT"] = (2, 2), ["VALUE"] = (1, 1), ["SUBSTITUTE"] = (3, 4), ["REPLACE"] = (4, 4),
            ["FIND"] = (2, 3), ["SEARCH"] = (2, 3), ["TEXTJOIN"] = (3, 254), ["CONCAT"] = (1, 254),
            ["ISBLANK"] = (1, 1), ["ISNUMBER"] = (1, 1), ["ISTEXT"] = (1, 1), ["ISERROR"] = (1, 1), ["IFNA"] = (2, 2),
            ["TODAY"] = (0, 0), ["NOW"] = (0, 0)
        };
}
