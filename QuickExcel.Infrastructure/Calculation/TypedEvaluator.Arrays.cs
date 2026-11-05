using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private object ArrayBranch(FunctionNode function)
    {
        if (!StandardFunctions.AcceptsArguments(function.Name,function.Arguments.Count)) return Error(CellErrorCode.Value);
        var args = function.Arguments;
        object TryValue(FormulaNode node)
        {
            try { return Evaluate(node); }
            catch (FormulaValueException error) { return Error(error.Code); }
        }
        var first = TryValue(args[0]);
        if (function.Name == "IF")
        {
            if (first is CellValue scalar)
                return scalar.Kind == CellValueKind.Error ? scalar : Boolean(scalar) ? TryValue(args[1]) : args.Count == 3 ? TryValue(args[2]) : CellValue.FromBoolean(false);
            var conditions = AsArray(first);
            var logical = conditions.Values.Select(v => Element(() => v.Kind == CellValueKind.Error ? v : CellValue.FromBoolean(Boolean(v)))).ToArray();
            var yes = logical.Any(v => v.Kind == CellValueKind.Boolean && v.BooleanValue) ? AsArray(TryValue(args[1])) : AsArray(CellValue.Blank);
            var no = logical.Any(v => v.Kind == CellValueKind.Boolean && !v.BooleanValue) && args.Count == 3 ? AsArray(TryValue(args[2])) : AsArray(CellValue.FromBoolean(false));
            var rows = Math.Max(conditions.Rows,Math.Max(yes.Rows,no.Rows)); var columns = Math.Max(conditions.Columns,Math.Max(yes.Columns,no.Columns));
            CheckArrayShape(rows,columns);
            CellValue At(FormulaArray a,int row,int column) => row < a.Rows || a.Rows == 1
                ? column < a.Columns || a.Columns == 1 ? a[a.Rows == 1 ? 0 : row,a.Columns == 1 ? 0 : column] : Error(CellErrorCode.NotAvailable)
                : Error(CellErrorCode.NotAvailable);
            return new FormulaArray(rows,columns,Enumerable.Range(0,rows * columns).Select(i => Element(() =>
            {
                var row = i / columns; var column = i % columns;
                var condition = At(conditions,row,column);
                if (condition.Kind == CellValueKind.Error) return condition;
                return At(Boolean(condition) ? yes : no,row,column);
            })));
        }
        bool Matches(CellValue value) => value.Kind == CellValueKind.Error && (function.Name == "IFERROR" || value.ErrorCode == CellErrorCode.NotAvailable);
        if (first is CellValue single) return Matches(single) ? TryValue(args[1]) : single;
        var input = AsArray(first);
        var fallback = AsArray(TryValue(args[1]));
        var outRows = Math.Max(input.Rows,fallback.Rows); var outColumns = Math.Max(input.Columns,fallback.Columns);
        CheckArrayShape(outRows,outColumns);
        return new FormulaArray(outRows,outColumns,Enumerable.Range(0,outRows * outColumns).Select(i => Element(() =>
        {
            var row = i / outColumns; var column = i % outColumns;
            var value = (input.Rows == 1 || row < input.Rows) && (input.Columns == 1 || column < input.Columns)
                ? input[input.Rows == 1 ? 0 : row,input.Columns == 1 ? 0 : column] : Error(CellErrorCode.NotAvailable);
            if (!Matches(value)) return value;
            return (fallback.Rows == 1 || row < fallback.Rows) && (fallback.Columns == 1 || column < fallback.Columns)
                ? fallback[fallback.Rows == 1 ? 0 : row,fallback.Columns == 1 ? 0 : column] : Error(CellErrorCode.NotAvailable);
        })));
    }

    private FormulaArray Materialize(Range range)
    {
        var shape = ConditionalRanges.Shape(range.Reference);
        CheckArrayShape(shape.Rows, shape.Columns);
        return new(shape.Rows, shape.Columns, Enumerable.Range(0, shape.Rows * shape.Columns).Select(i =>
            Read(range.Sheet, range.Reference.Row + i / shape.Columns, range.Reference.Column + i % shape.Columns)));
    }

    private static void CheckArrayShape(int rows, int columns)
    {
        try { FormulaArray.ValidateShape(rows, columns); }
        catch (ArgumentOutOfRangeException) { Fail(CellErrorCode.Calculation, "公式数组超过 100,000 元素或工作表边界。"); }
    }

    private FormulaArray AsArray(object value) => value switch
    {
        FormulaArray array => array,
        Range range => Materialize(range),
        CellValue scalar => new(1, 1, [scalar]),
        _ => throw new InvalidOperationException("未知公式值。")
    };

    private CellValue Element(Func<CellValue> evaluate)
    {
        check();
        try { return evaluate(); }
        catch (FormulaValueException error) { return Error(error.Code); }
        catch (Exception error) when (error is FormatException or OverflowException) { return Error(CellErrorCode.Value); }
    }

    private object ArrayUnary(UnaryNode unary)
    {
        CellValue Apply(CellValue value) => value.Kind == CellValueKind.Error ? value :
            CellValue.FromNumber(unary.Operator switch { "-" => -Number(value), "%" => Number(value) / 100, _ => Number(value) });
        var operand = Evaluate(unary.Operand);
        if (operand is CellValue scalar) return Element(() => Apply(scalar));
        var array = AsArray(operand);
        return new FormulaArray(array.Rows, array.Columns, array.Values.Select(v => Element(() => Apply(v))));
    }

    private object ArrayBinary(BinaryNode binary)
    {
        var left = Evaluate(binary.Left); var right = Evaluate(binary.Right);
        if (left is CellValue a && right is CellValue b) return Element(() => BinaryValues(binary.Operator, a, b));
        var x = AsArray(left); var y = AsArray(right);
        var rows = Math.Max(x.Rows, y.Rows); var columns = Math.Max(x.Columns, y.Columns);
        CheckArrayShape(rows, columns);
        CellValue At(FormulaArray array, int r, int c) => (array.Rows == 1 || r < array.Rows) && (array.Columns == 1 || c < array.Columns)
            ? array[array.Rows == 1 ? 0 : r, array.Columns == 1 ? 0 : c] : Error(CellErrorCode.NotAvailable);
        return new FormulaArray(rows, columns, Enumerable.Range(0, rows * columns).Select(i =>
            Element(() => BinaryValues(binary.Operator, At(x, i / columns, i % columns), At(y, i / columns, i % columns)))));
    }

    private object ArrayFunction(FunctionNode function)
    {
        var bounds = DynamicArrayFunctions.Signatures[function.Name];
        var args = function.Arguments;
        if (args.Count < bounds.Min || args.Count > bounds.Max) return Error(CellErrorCode.Value);
        if (function.Name == "TEXTSPLIT") return SplitTextFunction(function);
        if (function.Name is "TAKE" or "DROP") return SliceArray(function);
        if (function.Name == "SORTBY") return SortByFunction(function);
        CellValue Option(int index, CellValue fallback) => index >= args.Count || args[index] is MissingNode ? fallback : Scalar(args[index]);
        double NumericOption(int index, double fallback)
        {
            var value = Option(index, CellValue.FromNumber(fallback));
            if (value.Kind == CellValueKind.Error) Fail(value.ErrorCode, "数组参数包含错误。");
            return Number(value);
        }
        bool Flag(int index)
        {
            var value = Option(index, CellValue.FromBoolean(false));
            if (value.Kind == CellValueKind.Error) Fail(value.ErrorCode, "数组参数包含错误。");
            return Boolean(value);
        }
        if (function.Name == "SEQUENCE")
        {
            var height = Math.Truncate(NumericOption(0, 1)); var width = Math.Truncate(NumericOption(1, 1));
            if (height < 0 || width < 0) return Error(CellErrorCode.Value);
            if (height == 0 || width == 0) return Error(CellErrorCode.Calculation);
            if (height > WorkbookLimits.MaxRows || width > WorkbookLimits.MaxColumns || height * width > FormulaArray.MaxElements)
                return Error(CellErrorCode.Calculation);
            var start = NumericOption(2, 1); var step = NumericOption(3, 1);
            return new FormulaArray((int)height, (int)width, Enumerable.Range(0, (int)(height * width)).Select(i =>
            { check(); return CellValue.FromNumber(start + i * step); }));
        }
        var source = Evaluate(args[0]);
        if (source is CellValue { Kind: CellValueKind.Error }) return source;
        var input = AsArray(source);
        if (function.Name == "FILTER")
        {
            var included = AsArray(Evaluate(args[1]));
            var byRows = included.Columns == 1 && included.Rows == input.Rows;
            var byColumns = included.Rows == 1 && included.Columns == input.Columns;
            if (!byRows && !byColumns) return Error(CellErrorCode.Value);
            var selected = new List<int>();
            for (var i = 0; i < included.Values.Length; i++)
            {
                check(); var value = included.Values[i];
                if (value.Kind == CellValueKind.Error) return value;
                if (Boolean(value)) selected.Add(i);
            }
            if (selected.Count == 0) return args.Count == 3 ? FormulaResult(args[2]) : Error(CellErrorCode.Calculation);
            var rows = byRows ? selected.Count : input.Rows; var columns = byRows ? input.Columns : selected.Count;
            return new FormulaArray(rows, columns, Enumerable.Range(0, rows * columns).Select(i =>
            { check(); return input[byRows ? selected[i / columns] : i / columns, byRows ? i % columns : selected[i % columns]]; }));
        }
        var horizontal = Flag(function.Name == "SORT" ? 3 : 1);
        var count = horizontal ? input.Columns : input.Rows;
        var length = horizontal ? input.Rows : input.Columns;
        CellValue Item(int item, int offset) => input[horizontal ? offset : item, horizontal ? item : offset];
        IEnumerable<int> order;
        if (function.Name == "SORT")
        {
            var key = Math.Truncate(NumericOption(1, 1)); var direction = NumericOption(2, 1);
            if (key < 1 || key > length || direction is not (1 or -1)) return Error(CellErrorCode.Value);
            var comparer = Comparer<int>.Create((a, b) =>
            {
                check(); var x = Item(a, (int)key - 1); var y = Item(b, (int)key - 1);
                if (x.Kind == CellValueKind.Blank && y.Kind != CellValueKind.Blank) return 1;
                if (y.Kind == CellValueKind.Blank && x.Kind != CellValueKind.Blank) return -1;
                var difference = CompareSortValues(x, y);
                return difference == 0 ? a.CompareTo(b) : (int)direction * Math.Sign(difference);
            });
            order = StableArrayOrder(count, comparer);
        }
        else
        {
            var exactlyOnce = Flag(2);
            var comparer = new ArrayItemComparer(length, Item, check);
            var groups = new Dictionary<int, (int First, int Count)>(comparer);
            for (var i = 0; i < count; i++)
            {
                check();
                if (groups.TryGetValue(i, out var group)) groups[i] = (group.First, group.Count + 1);
                else groups.Add(i, (i, 1));
            }
            order = groups.Values.Where(g => !exactlyOnce || g.Count == 1).Select(g => g.First).Order().ToArray();
        }
        var indices = order.ToArray();
        if (indices.Length == 0) return Error(CellErrorCode.Calculation);
        var outRows = horizontal ? input.Rows : indices.Length; var outColumns = horizontal ? indices.Length : input.Columns;
        return new FormulaArray(outRows, outColumns, Enumerable.Range(0, outRows * outColumns).Select(i =>
        { check(); return input[horizontal ? i / outColumns : indices[i / outColumns], horizontal ? indices[i % outColumns] : i % outColumns]; }));
    }

    private static int CompareSortValues(CellValue a, CellValue b)
    {
        // Excel sorts blanks last. Empty text is a text value, not a blank cell.
        int Rank(CellValue value) => value.Kind switch { CellValueKind.Number => 0, CellValueKind.Text => 1, CellValueKind.Boolean => 2, CellValueKind.Error => 3, _ => 4 };
        var rank = Rank(a).CompareTo(Rank(b));
        return rank != 0 ? rank : a.Kind == CellValueKind.Error ? 0 : Compare(a, b);
    }

    private int[] StableArrayOrder(int count, IComparer<int> comparer)
    {
        // Do not use Array.Sort/LINQ Order here: they wrap cancellation/deadline exceptions
        // thrown by the comparator, changing a cancelled calculation into a cell failure.
        var source = Enumerable.Range(0, count).ToArray(); var target = new int[count];
        for (var width = 1; width < count; width *= 2)
        {
            for (var start = 0; start < count; start += 2 * width)
            {
                var middle = Math.Min(start + width, count); var end = Math.Min(start + 2 * width, count);
                var left = start; var right = middle;
                for (var i = start; i < end; i++)
                {
                    check();
                    target[i] = right == end || left < middle && comparer.Compare(source[left], source[right]) <= 0
                        ? source[left++] : source[right++];
                }
            }
            (source, target) = (target, source);
        }
        return source;
    }

    private sealed class ArrayItemComparer(int length, Func<int, int, CellValue> item, Action check) : IEqualityComparer<int>
    {
        public bool Equals(int x, int y)
        {
            for (var i = 0; i < length; i++)
            {
                check(); var a = item(x, i); var b = item(y, i);
                if (a.Kind != b.Kind || (a.Kind == CellValueKind.Text
                    ? !StringComparer.InvariantCultureIgnoreCase.Equals(a.TextValue, b.TextValue) : a != b)) return false;
            }
            return true;
        }
        public int GetHashCode(int key)
        {
            var hash = new HashCode();
            for (var i = 0; i < length; i++)
            {
                check(); var value = item(key, i); hash.Add(value.Kind);
                if (value.Kind == CellValueKind.Text) hash.Add(value.TextValue, StringComparer.InvariantCultureIgnoreCase); else hash.Add(value);
            }
            return hash.ToHashCode();
        }
    }
}
