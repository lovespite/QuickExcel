using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private Dictionary<int, Lazy<(object Value, bool Reference)>> _locals = [];
    private T WithLocals<T>(LetNode expression, Func<FormulaNode, T> evaluate)
    {
        var previous = _locals;
        var scope = new Dictionary<int, Lazy<(object Value, bool Reference)>>(previous);
        foreach (var binding in expression.Bindings)
            scope[binding.Id] = new Lazy<(object Value, bool Reference)>(() => AggregateArgument(binding.Value));
        _locals = scope;
        try { return evaluate(expression.Body); }
        finally { _locals = previous; }
    }

    private CellValue Statistic(FunctionNode function)
    {
        var numbers = new List<double>();
        foreach (var argument in function.Arguments)
        {
            var (input, reference) = AggregateArgument(argument);
            foreach (var value in Values(input))
            {
                check();
                if (value.Kind == CellValueKind.Error) return value;
                if (value.Kind == CellValueKind.Blank || reference && value.Kind != CellValueKind.Number) continue;
                if (numbers.Count >= FormulaArray.MaxElements) throw new FormulaValueException(CellErrorCode.Number, "统计函数最多支持 100000 个数值。");
                numbers.Add(Number(value));
            }
        }
        if (function.Name == "MEDIAN")
        {
            if (numbers.Count == 0) return Error(CellErrorCode.Number);
            numbers.Sort(); check(); var middle = numbers.Count / 2;
            return CellValue.FromNumber(numbers.Count % 2 == 1 ? numbers[middle] : numbers[middle - 1] / 2 + numbers[middle] / 2);
        }
        var sample = function.Name.EndsWith(".S", StringComparison.Ordinal);
        if (numbers.Count < (sample ? 2 : 1)) return Error(CellErrorCode.DivisionByZero);
        // Shift by the first observation before averaging; avoids catastrophic cancellation for large offsets.
        var origin = numbers[0]; double total = 0, compensation = 0;
        foreach (var number in numbers)
        {
            check(); var y = number - origin - compensation; var sum = total + y;
            compensation = (sum - total) - y; total = sum;
        }
        var mean = total / numbers.Count; double squares = 0; compensation = 0;
        foreach (var number in numbers)
        {
            check(); var deviation = number - origin - mean; var y = deviation * deviation - compensation; var sum = squares + y;
            compensation = (sum - squares) - y; squares = sum;
        }
        var variance = squares / (numbers.Count - (sample ? 1 : 0));
        return CellValue.FromNumber(function.Name.StartsWith("STDEV", StringComparison.Ordinal) ? Math.Sqrt(variance) : variance);
    }
}
