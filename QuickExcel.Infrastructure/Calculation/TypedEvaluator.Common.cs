using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private object SafeCommonValue(FormulaNode node)
    {
        try { return node is MissingNode ? CellValue.FromNumber(0) : Evaluate(node); }
        catch (FormulaValueException error) { return Error(error.Code); }
    }

    private object ChoiceFunction(FunctionNode function)
    {
        if (!StandardFunctions.AcceptsArguments(function.Name, function.Arguments.Count)) return Error(CellErrorCode.Value);
        var args = function.Arguments;
        var input = function.Name == "SWITCH" ? SafeCommonValue(args[0]) : null;
        object Visit(int index)
        {
            if (index >= args.Count) return Error(CellErrorCode.NotAvailable);
            if (function.Name == "SWITCH" && index == args.Count - 1) return SafeCommonValue(args[index]);
            var condition = SafeCommonValue(args[index]);
            if (function.Name == "SWITCH")
            {
                var left = AsArray(input!); var right = AsArray(condition);
                var rows = Math.Max(left.Rows, right.Rows); var columns = Math.Max(left.Columns, right.Columns);
                CheckArrayShape(rows, columns);
                CellValue Match(int row, int column)
                {
                    var a = CommonAt(left, row, column); var b = CommonAt(right, row, column);
                    if (a.Kind == CellValueKind.Error) return a;
                    if (b.Kind == CellValueKind.Error) return b;
                    return CellValue.FromBoolean((a.Kind == b.Kind || a.Kind == CellValueKind.Blank || b.Kind == CellValueKind.Blank) && LookupCompare(a, b) == 0);
                }
                condition = rows == 1 && columns == 1 ? Match(0, 0) : new FormulaArray(rows, columns, Enumerable.Range(0, rows * columns).Select(i => Match(i / columns, i % columns)));
            }
            return SelectCommon(condition, () => SafeCommonValue(args[index + 1]), () => Visit(index + 2));
        }
        return Visit(function.Name == "IFS" ? 0 : 1);
    }

    private static CellValue CommonAt(FormulaArray array, int row, int column) =>
        (array.Rows == 1 || row < array.Rows) && (array.Columns == 1 || column < array.Columns)
        ? array[array.Rows == 1 ? 0 : row, array.Columns == 1 ? 0 : column] : Error(CellErrorCode.NotAvailable);

    private object SelectCommon(object condition, Func<object> yes, Func<object> no)
    {
        if (condition is CellValue scalar)
            return scalar.Kind == CellValueKind.Error ? scalar : Boolean(scalar) ? yes() : no();
        var tests = AsArray(condition);
        var logical = new FormulaArray(tests.Rows, tests.Columns, tests.Values.Select(v => Element(() => v.Kind == CellValueKind.Error ? v : CellValue.FromBoolean(Boolean(v)))));
        var a = AsArray(logical.Values.Any(v => v.Kind == CellValueKind.Boolean && v.BooleanValue) ? yes() : CellValue.Blank);
        var b = AsArray(logical.Values.Any(v => v.Kind == CellValueKind.Boolean && !v.BooleanValue) ? no() : CellValue.Blank);
        var rows = Math.Max(tests.Rows, Math.Max(a.Rows, b.Rows)); var columns = Math.Max(tests.Columns, Math.Max(a.Columns, b.Columns));
        CheckArrayShape(rows, columns);
        return new FormulaArray(rows, columns, Enumerable.Range(0, rows * columns).Select(i =>
        {
            check(); var r = i / columns; var c = i % columns; var test = CommonAt(logical, r, c);
            return test.Kind == CellValueKind.Error ? test : CommonAt(test.BooleanValue ? a : b, r, c);
        }));
    }

    private CellValue SumProduct(FunctionNode function)
    {
        var arrays = function.Arguments.Select(a => AsArray(SafeCommonValue(a))).ToArray();
        var first = arrays[0];
        if (arrays.Any(a => a.Rows != first.Rows || a.Columns != first.Columns)) return Error(CellErrorCode.Value);
        double total = 0;
        for (var i = 0; i < first.Values.Length; i++)
        {
            check(); double product = 1;
            foreach (var array in arrays)
            {
                var value = array.Values[i];
                if (value.Kind == CellValueKind.Error) return value;
                product *= value.Kind == CellValueKind.Number ? value.NumberValue : 0;
            }
            total += product;
        }
        return CellValue.FromNumber(total);
    }
}
