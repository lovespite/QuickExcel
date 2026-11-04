using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private CellValue OrderStatistic(FunctionNode function)
    {
        var rank=function.Name=="RANK.EQ";
        CellValue Option(int index) => function.Arguments[index] is MissingNode ? CellValue.Blank : Scalar(function.Arguments[index]);
        // Excel coerces the scalar options before inspecting errors in the data range.
        var scalar=Option(rank ? 0 : 1);
        if (scalar.Kind==CellValueKind.Error) return scalar;
        var requested=Number(scalar); var ascending=false;
        if (rank && function.Arguments.Count==3)
        {
            var order=Option(2);
            if (order.Kind==CellValueKind.Error) return order;
            if (order.Kind==CellValueKind.Text) return Error(CellErrorCode.Value);
            ascending=Number(order)!=0;
        }
        var (input,reference)=AggregateArgument(function.Arguments[rank ? 1 : 0]);
        if (rank && (input is FormulaArray || !reference))
            return input is CellValue { Kind: CellValueKind.Error } error ? error : Error(CellErrorCode.Value);
        var numbers=new List<double>();
        foreach (var value in Values(input))
        {
            check();
            if (value.Kind==CellValueKind.Error) return value;
            if (value.Kind==CellValueKind.Blank || reference && value.Kind!=CellValueKind.Number) continue;
            if (numbers.Count>=FormulaArray.MaxElements) throw new FormulaValueException(CellErrorCode.Number,"顺序统计最多支持 100000 个数值。");
            numbers.Add(Number(value));
        }
        if (rank)
        {
            var found=false; var position=1;
            foreach (var number in numbers)
            {
                check(); found|=number==requested;
                if (ascending ? number<requested : number>requested) position++;
            }
            return found ? CellValue.FromNumber(position) : Error(CellErrorCode.NotAvailable);
        }
        if (requested<1 || requested>numbers.Count) return Error(CellErrorCode.Number);
        // Independent Excel evidence: LARGE rounds fractional k up, SMALL rounds it down.
        var k=(int)(function.Name=="LARGE" ? Math.Ceiling(requested) : Math.Floor(requested));
        numbers.Sort(); check();
        return CellValue.FromNumber(numbers[function.Name=="LARGE" ? numbers.Count-k : k-1]);
    }
}
