using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private object ReferencePosition(FunctionNode function)
    {
        if (function.Arguments.Count != 1) return Error(CellErrorCode.Value);
        if (function.Arguments[0] is ValueNode { Value.Kind: CellValueKind.Error } error) return error.Value;
        if (function.Arguments[0] is not ReferenceNode reference) return Error(CellErrorCode.Value);
        var row = function.Name == "ROW"; var shape = ConditionalRanges.Shape(reference);
        var size = row ? shape.Rows : shape.Columns; var start = row ? reference.Row + 1 : reference.Column + 1;
        if (size == 1) return CellValue.FromNumber(start);
        CheckArrayShape(row ? size : 1, row ? 1 : size);
        return new FormulaArray(row ? size : 1, row ? 1 : size, Enumerable.Range(start, size).Select(i => { check(); return CellValue.FromNumber(i); }));
    }

    private CellValue ReferenceDimension(FunctionNode function)
    {
        if (function.Arguments[0] is ReferenceNode reference)
        {
            var shape = ConditionalRanges.Shape(reference);
            return CellValue.FromNumber(function.Name == "ROWS" ? shape.Rows : shape.Columns);
        }
        var value = Evaluate(function.Arguments[0]);
        if (value is CellValue { Kind: CellValueKind.Error } error) return error;
        var array = AsArray(value);
        return CellValue.FromNumber(function.Name == "ROWS" ? array.Rows : array.Columns);
    }
}
