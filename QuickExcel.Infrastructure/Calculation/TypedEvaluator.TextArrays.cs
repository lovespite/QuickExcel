using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed partial class TypedEvaluator
{
    private CellValue CommonArgument(FunctionNode function, int index, CellValue fallback)
    {
        if (index >= function.Arguments.Count || function.Arguments[index] is MissingNode) return fallback;
        var value = Scalar(function.Arguments[index]);
        if (value.Kind == CellValueKind.Error) Fail(value.ErrorCode, MessageText.Get("函数参数包含错误。"));
        return value;
    }
    private double CommonNumber(FunctionNode function, int index, double fallback) => Number(CommonArgument(function, index, CellValue.FromNumber(fallback)));

    private CellValue TextBoundary(FunctionNode function)
    {
        var text = FormulaText(CommonArgument(function, 0, CellValue.Blank));
        var delimiter = FormulaText(CommonArgument(function, 1, CellValue.Blank));
        var instance = Math.Truncate(CommonNumber(function, 2, 1));
        var mode = CommonNumber(function, 3, 0); var matchEnd = CommonNumber(function, 4, 0);
        if (instance == 0 || text.Length > 0 && Math.Abs(instance) > TextOffsets(text, true).Length - 1 || mode is not (0 or 1) || matchEnd is not (0 or 1))
            return Error(CellErrorCode.Value);
        var before = function.Name == "TEXTBEFORE";
        if (delimiter.Length == 0) return CellValue.FromText(before == (instance > 0) ? "" : text);
        var comparison = mode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var positions = new List<int>();
        if (instance > 0)
        for (var start = 0; start <= text.Length - delimiter.Length;)
        {
            check(); var found = text.IndexOf(delimiter, start, comparison); if (found < 0) break;
            positions.Add(found); start = found + delimiter.Length;
        }
        else
        for (var end = text.Length; end >= delimiter.Length;)
        {
            check(); var found = text.LastIndexOf(delimiter, end - 1, end, comparison); if (found < 0) break;
            positions.Add(found); end = found;
        }
        var wanted = Math.Abs(instance);
        if (wanted > positions.Count)
        {
            if (matchEnd == 1 && wanted == positions.Count + 1)
                return CellValue.FromText(before == (instance > 0) ? text : "");
            return function.Arguments.Count > 5 ? Scalar(function.Arguments[5]) : Error(CellErrorCode.NotAvailable);
        }
        var offset = positions[(int)wanted - 1];
        return CellValue.FromText(before ? text[..offset] : text[(offset + delimiter.Length)..]);
    }

    private object SplitTextFunction(FunctionNode function)
    {
        var text = FormulaText(CommonArgument(function, 0, CellValue.Blank));
        if (text.Length == 0) return Error(CellErrorCode.Value);
        var ignoreEmpty = Boolean(CommonArgument(function, 3, CellValue.FromBoolean(false)));
        var mode = CommonNumber(function, 4, 0);
        if (mode is not (0 or 1)) return Error(CellErrorCode.Value);
        var comparison = mode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        string[] Delimiters(int index)
        {
            if (index >= function.Arguments.Count || function.Arguments[index] is MissingNode) return [];
            var values = AsArray(SafeCommonValue(function.Arguments[index]));
            return values.Values.Select(v =>
            {
                if (v.Kind == CellValueKind.Error) Fail(v.ErrorCode, MessageText.Get("分隔符包含错误。"));
                var value = FormulaText(v); if (value.Length == 0) Fail(CellErrorCode.Value, MessageText.Get("分隔符不能为空。")); return value;
            }).Distinct().ToArray();
        }
        var colDelimiters = Delimiters(1); var rowDelimiters = Delimiters(2);
        if (colDelimiters.Length == 0 && rowDelimiters.Length == 0) return Error(CellErrorCode.Value);
        string[] Split(string input, string[] delimiters)
        {
            if (delimiters.Length == 0) return [input];
            var fields = new List<string>(); var start = 0;
            while (start <= input.Length)
            {
                check(); var offset = input.Length; var length = 0;
                foreach (var delimiter in delimiters)
                {
                    check();
                    var next = input.IndexOf(delimiter, start, comparison);
                    if (next >= 0 && (length == 0 || next < offset)) { offset = next; length = delimiter.Length; }
                }
                var field = input[start..offset]; if (!ignoreEmpty || field.Length > 0) fields.Add(field);
                if (fields.Count > FormulaArray.MaxElements) Fail(CellErrorCode.Calculation, MessageText.Get("文本分割超过数组预算。"));
                if (length == 0) break; start = offset + length;
            }
            return fields.ToArray();
        }
        var lines = Split(text, rowDelimiters).Select(line => Split(line, colDelimiters)).ToArray();
        if (lines.Length == 0 || lines.All(l => l.Length == 0)) return Error(CellErrorCode.Calculation);
        var columns = lines.Max(l => l.Length); CheckArrayShape(lines.Length, columns);
        var pad = function.Arguments.Count > 5 ? Scalar(function.Arguments[5]) : Error(CellErrorCode.NotAvailable);
        return new FormulaArray(lines.Length, columns, lines.SelectMany(line => Enumerable.Range(0, columns).Select(c => c < line.Length ? CellValue.FromText(line[c]) : pad)));
    }

    private object SliceArray(FunctionNode function)
    {
        var source = SafeCommonValue(function.Arguments[0]);
        if (source is CellValue { Kind: CellValueKind.Error }) return source;
        var input = AsArray(source); var take = function.Name == "TAKE";
        var rowCount = Math.Truncate(CommonNumber(function, 1, take ? input.Rows : 0));
        var columnCount = Math.Truncate(CommonNumber(function, 2, take ? input.Columns : 0));
        (int Start, int Count) Slice(double count, int length)
        {
            var amount = (int)Math.Min(length, Math.Abs(count));
            return take ? (count < 0 ? length - amount : 0, amount) : (count < 0 ? 0 : amount, length - amount);
        }
        var rows = Slice(rowCount, input.Rows); var columns = Slice(columnCount, input.Columns);
        if (rows.Count == 0 || columns.Count == 0) return Error(CellErrorCode.Calculation);
        return new FormulaArray(rows.Count, columns.Count, Enumerable.Range(0, rows.Count * columns.Count).Select(i =>
        { check(); return input[rows.Start + i / columns.Count, columns.Start + i % columns.Count]; }));
    }

    private object SortByFunction(FunctionNode function)
    {
        var source = SafeCommonValue(function.Arguments[0]); if (source is CellValue { Kind: CellValueKind.Error }) return source;
        var input = AsArray(source); var keys = new List<(FormulaArray Values, int Direction)>(); bool? horizontal = null;
        for (var i = 1; i < function.Arguments.Count; i += 2)
        {
            var value = SafeCommonValue(function.Arguments[i]); if (value is CellValue { Kind: CellValueKind.Error }) return value;
            var key = AsArray(value); var direction = CommonNumber(function, i + 1, 1);
            var byRows = key.Columns == 1 && key.Rows == input.Rows; var byColumns = key.Rows == 1 && key.Columns == input.Columns;
            if (direction is not (1 or -1) || !byRows && !byColumns) return Error(CellErrorCode.Value);
            horizontal ??= !byRows;
            if (horizontal.Value ? !byColumns : !byRows) return Error(CellErrorCode.Value);
            keys.Add((key, (int)direction));
        }
        var across = horizontal == true;
        var order = StableArrayOrder(across ? input.Columns : input.Rows, Comparer<int>.Create((a, b) =>
        {
            check();
            foreach (var (key, direction) in keys)
            {
                var x = key.Values[a]; var y = key.Values[b];
                if (x.Kind == CellValueKind.Blank && y.Kind != CellValueKind.Blank) return 1;
                if (y.Kind == CellValueKind.Blank && x.Kind != CellValueKind.Blank) return -1;
                var difference = CompareSortValues(x, y); if (difference != 0) return direction * Math.Sign(difference);
            }
            return a.CompareTo(b);
        }));
        return new FormulaArray(input.Rows, input.Columns, Enumerable.Range(0, input.Rows * input.Columns).Select(i =>
        { check(); return input[across ? i / input.Columns : order[i / input.Columns], across ? order[i % input.Columns] : i % input.Columns]; }));
    }
}
