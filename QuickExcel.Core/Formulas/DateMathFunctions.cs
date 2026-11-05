namespace QuickExcel.Formulas;

public static class DateMathFunctions
{
    public static bool IsFunction(string name) => name is "DATE" or "YEAR" or "MONTH" or "DAY" or "TIME" or "HOUR" or "MINUTE" or "SECOND"
        or "EDATE" or "EOMONTH" or "WEEKDAY" or "WORKDAY" or "NETWORKDAYS" or "ROUNDUP" or "ROUNDDOWN" or "INT" or "MOD";
    public static void Validate(FormulaNode expression)
    {
        foreach (var function in FormulaSyntax.Walk(expression).OfType<FunctionNode>().Where(f => IsFunction(f.Name)))
        for (var i = 0; i < function.Arguments.Count; i++)
        {
            var argument = function.Arguments[i];
            if (i == 2 && function.Name is "WORKDAY" or "NETWORKDAYS")
            {
                if (argument is not ReferenceNode && ArrayResult(argument))
                    throw new FormatException("节假日区域须为直接或名称绑定的引用，计算得到的区域尚未实现。");
            }
            else if (ArrayResult(argument))
                throw new FormatException("日期和数学公式目前仅支持标量参数；多格数组模式尚未实现。");
        }
    }
    public static bool ArrayResult(FormulaNode expression) => expression switch
    {
        ReferenceNode reference => ConditionalRanges.Shape(reference) != (1, 1),
        LocalValueNode local => local.ArrayResult,
        LetNode let => ArrayResult(let.Body),
        FunctionNode function when DynamicArrayFunctions.IsFunction(function.Name) => true,
        FunctionNode function when ReferenceFunctions.ReturnsArray(function) => true,
        UnaryNode unary => ArrayResult(unary.Operand),
        BinaryNode binary => ArrayResult(binary.Left) || ArrayResult(binary.Right),
        FunctionNode { Name: "IF" or "IFS" or "SWITCH" } function => function.Arguments.Any(ArrayResult),
        FunctionNode { Name: "IFERROR" or "IFNA" } function => function.Arguments.Any(ArrayResult),
        _ => false
    };
}
