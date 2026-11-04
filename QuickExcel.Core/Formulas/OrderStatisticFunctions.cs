namespace QuickExcel.Formulas;

public static class OrderStatisticFunctions
{
    public static bool IsFunction(string name) => name is "LARGE" or "SMALL" or "RANK.EQ";

    public static void Validate(FormulaNode expression)
    {
        var locals=new Dictionary<int,bool>();
        foreach (var function in FormulaSyntax.Walk(expression).OfType<FunctionNode>().Where(f=>IsFunction(f.Name)))
        {
            if (!StandardFunctions.AcceptsArguments(function.Name,function.Arguments.Count)) continue;
            var dataIndex=function.Name=="RANK.EQ" ? 1 : 0;
            for (var index=0;index<function.Arguments.Count;index++)
                if (index!=dataIndex && Multiple(function.Arguments[index],locals))
                    throw new FormatException("排名/顺序统计的名次、待排名数值和排序方向目前只支持标量；保留来源数组公式和缓存。");
        }
    }

    private static bool Multiple(FormulaNode node,Dictionary<int,bool> locals) => node switch
    {
        LocalValueNode local=>locals.TryGetValue(local.Id,out var multiple) ? multiple : locals[local.Id]=Multiple(local.Definition,locals),
        LetNode let=>Multiple(let.Body,locals),
        UnaryNode { Operator: "#" }=>true, UnaryNode unary=>Multiple(unary.Operand,locals),
        BinaryNode binary=>Multiple(binary.Left,locals) || Multiple(binary.Right,locals),
        FunctionNode { Name: "IF" or "IFERROR" or "IFNA" } function=>function.Arguments.Any(a=>Multiple(a,locals)),
        FunctionNode { Name: "XLOOKUP" } function=>DynamicArrayFunctions.UsesArrays(function),
        _=>DateMathFunctions.ArrayResult(node)
    };
}
