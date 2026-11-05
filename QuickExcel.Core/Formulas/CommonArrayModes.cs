using QuickExcel.Core;

namespace QuickExcel.Formulas;

public static class CommonArrayModes
{
    public static void Validate(FormulaNode expression)
    {
        foreach (var function in FormulaSyntax.Walk(expression).OfType<FunctionNode>())
        for (var i = 0; i < function.Arguments.Count; i++)
        {
            var scalar = function.Name == "TEXTSPLIT" && (i == 0 || i >= 3) ||
                function.Name is "TAKE" or "DROP" && i > 0 || function.Name == "SORTBY" && i > 0 && i % 2 == 0;
            if (scalar && DateMathFunctions.ArrayResult(function.Arguments[i]))
                throw new FormatException(MessageText.Get("此函数的文本或控制参数暂限标量，数组形式保留来源公式。"));
        }
    }
}
