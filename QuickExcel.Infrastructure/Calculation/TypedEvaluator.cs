using System.Globalization;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

internal sealed class FormulaValueException(CellErrorCode code, string message) : Exception(message)
{
    public CellErrorCode Code { get; } = code;
}

/// <summary>Conversions are explicit. Reference/range aggregates ignore text and booleans, including numeric-looking text.</summary>
internal sealed partial class TypedEvaluator(WorkbookSnapshot snapshot, IReadOnlyDictionary<CellAddress, CellCalculation> results,
    IScriptRuntime? runtime, IReadOnlyDictionary<string, ScriptFunctionDescriptor> scripts, Action check)
{
    public SheetId CurrentSheet { get; set; }
    private readonly Dictionary<SheetId, CellPosition[]> _positions = [];
    private readonly DateTime _now = DateTime.Now;
    private sealed record Range(WorksheetSnapshot Sheet, ReferenceNode Reference);
    private static CellValue Error(CellErrorCode error) => CellValue.FromError(error);
    private static void Fail(CellErrorCode code, string message) => throw new FormulaValueException(code, message);

    public CellValue Scalar(FormulaNode node)
    {
        var value = Evaluate(node);
        return value is CellValue scalar ? scalar : Error(CellErrorCode.Value);
    }

    private object Evaluate(FormulaNode node)
    {
        check();
        switch (node)
        {
            case ValueNode value: return value.Value;
            case ReferenceNode reference:
                var sheet = snapshot.Sheets.First(s => s.Name.Equals(reference.Sheet, StringComparison.OrdinalIgnoreCase));
                return reference.IsRange || reference.WholeColumns ? new Range(sheet, reference) : Read(sheet, reference.Row, reference.Column);
            case UnaryNode unary:
                var operand = Scalar(unary.Operand);
                if (operand.Kind == CellValueKind.Error) return operand;
                var number = Number(operand);
                return CellValue.FromNumber(unary.Operator switch { "-" => -number, "%" => number / 100, _ => number });
            case BinaryNode binary: return Binary(binary);
            case FunctionNode function: return Function(function);
            default: return Error(CellErrorCode.Name);
        }
    }

    private CellValue Read(WorksheetSnapshot sheet, int row, int column)
    {
        check();
        if (results.TryGetValue(new(sheet.Id, row, column), out var calculated)) return calculated.Value ?? CellValue.Blank;
        var cell = sheet.Cells.GetValueOrDefault(new(row, column));
        if (cell?.Content.Kind == CellContentKind.Formula) throw new InvalidOperationException("依赖计划尝试读取尚未计算的公式。");
        return cell?.Value ?? CellValue.Blank;
    }

    private CellValue Binary(BinaryNode binary)
    {
        var a = Scalar(binary.Left); var b = Scalar(binary.Right);
        if (a.Kind == CellValueKind.Error) return a;
        if (b.Kind == CellValueKind.Error) return b;
        if (binary.Operator == "&") return CellValue.FromText(Text(a) + Text(b));
        if (binary.Operator is "=" or "<>" or "<" or "<=" or ">" or ">=")
        {
            var c = Compare(a, b);
            return CellValue.FromBoolean(binary.Operator switch { "=" => c == 0, "<>" => c != 0, "<" => c < 0, "<=" => c <= 0, ">" => c > 0, _ => c >= 0 });
        }
        var left = Number(a); var right = Number(b);
        if (binary.Operator == "/" && right == 0) return Error(CellErrorCode.DivisionByZero);
        return CellValue.FromNumber(binary.Operator switch { "+" => left + right, "-" => left - right, "*" => left * right, "/" => left / right, "^" => Math.Pow(left, right), _ => double.NaN });
    }

    private CellValue Function(FunctionNode function)
    {
        if (scripts.TryGetValue(function.Name, out var descriptor)) return Script(function, descriptor);
        if (!StandardFunctions.Signatures.TryGetValue(function.Name, out var signature)) return Error(CellErrorCode.Name);
        if (function.Arguments.Count < signature.Min || function.Arguments.Count > signature.Max) return Error(CellErrorCode.Value);
        if (function.Name == "IF")
        {
            var condition = Scalar(function.Arguments[0]);
            if (condition.Kind == CellValueKind.Error) return condition;
            return Boolean(condition) ? Scalar(function.Arguments[1])
                : function.Arguments.Count == 3 ? Scalar(function.Arguments[2]) : CellValue.FromBoolean(false);
        }
        if (function.Name == "IFERROR")
        {
            try
            {
                var value = Scalar(function.Arguments[0]);
                return value.Kind == CellValueKind.Error ? Scalar(function.Arguments[1]) : value;
            }
            catch (Exception e) when (e is FormulaValueException or FormatException or OverflowException or ArgumentException)
            { return Scalar(function.Arguments[1]); }
        }
        if (ConditionalRanges.IsFunction(function.Name)) return Conditional(function);
        if (LookupModes.IsFunction(function.Name)) return LookupValue(function);
        if (DateMathFunctions.IsFunction(function.Name)) return DateMath(function);
        if (TextFunctions.IsFunction(function.Name)) return TextFunction(function);
        var args = function.Arguments.Select(Evaluate).ToArray();
        if (function.Name is "SUM" or "AVERAGE" or "MIN" or "MAX" or "COUNT" or "COUNTA" or "PRODUCT") return Aggregate(function, args);
        foreach (var arg in args) if (arg is CellValue { Kind: CellValueKind.Error } error) return error;
        if (function.Name == "INDEX") return Index(args);
        if (function.Name is "AND" or "OR")
        {
            var logicalCount = 0; var all = true; var any = false;
            for (var i = 0; i < args.Length; i++)
            foreach (var value in Values(args[i]))
            {
                if (value.Kind == CellValueKind.Error) return value;
                if (function.Arguments[i] is ReferenceNode && value.Kind is CellValueKind.Text or CellValueKind.Blank) continue;
                var logical = Boolean(value); logicalCount++; all &= logical; any |= logical;
            }
            return logicalCount == 0 ? Error(CellErrorCode.Value) : CellValue.FromBoolean(function.Name == "AND" ? all : any);
        }
        if (args.Any(a => a is Range)) return Error(CellErrorCode.Value);
        var values = args.Cast<CellValue>().ToArray();
        CellValue A(int i) => values[i];
        int Integer(int i) => checked((int)Math.Truncate(Number(A(i))));
        string S(int i) => Text(A(i));
        switch (function.Name)
        {
            case "ABS": return CellValue.FromNumber(Math.Abs(Number(A(0))));
            case "NOT": return CellValue.FromBoolean(!Boolean(A(0)));
            case "ROUND":
                var digits = Integer(1); var factor = Math.Pow(10, Math.Abs((double)digits));
                var n = Number(A(0));
                if (digits >= 0)
                    return CellValue.FromNumber(!double.IsFinite(n * factor) ? n : Math.Round(n * factor, MidpointRounding.AwayFromZero) / factor);
                return CellValue.FromNumber(!double.IsFinite(factor) ? 0 : Math.Round(n / factor, MidpointRounding.AwayFromZero) * factor);
            case "CONCATENATE": return CellValue.FromText(string.Concat(values.Select(Text)));
            case "LEN": return CellValue.FromNumber(TextOffsets(S(0)).Length - 1);
            case "UPPER": return CellValue.FromText(S(0).ToUpperInvariant());
            case "LOWER": return CellValue.FromText(S(0).ToLowerInvariant());
            case "TRIM": return CellValue.FromText(string.Join(" ", S(0).Split(' ', StringSplitOptions.RemoveEmptyEntries)));
            case "LEFT":
            case "RIGHT":
                var count = args.Length == 1 ? 1 : Integer(1);
                if (count < 0) return Error(CellErrorCode.Value);
                var text = S(0); var textOffsets = TextOffsets(text, alwaysUnicode: true); count = Math.Min(count, textOffsets.Length - 1);
                return CellValue.FromText(function.Name == "LEFT" ? text[..textOffsets[count]] : text[textOffsets[textOffsets.Length - count - 1]..]);
            case "MID":
                var start = Integer(1); var length = Integer(2); var source = S(0);
                if (start < 1 || length < 0) return Error(CellErrorCode.Value);
                var sourceOffsets = TextOffsets(source);
                if (start >= sourceOffsets.Length) return CellValue.FromText("");
                return CellValue.FromText(source[sourceOffsets[start - 1]..sourceOffsets[(int)Math.Min((long)start - 1 + length, sourceOffsets.Length - 1)]]);
            case "TODAY": return Serial(_now.Date);
            case "NOW": return Serial(_now);
            default: return Error(CellErrorCode.Name);
        }
    }

    private CellValue Serial(DateTime value)
    {
        var origin = snapshot.DateSystem == ExcelDateSystem.Mac1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 31);
        var serial = (value - origin).TotalDays;
        if (snapshot.DateSystem == ExcelDateSystem.Windows1900 && value >= new DateTime(1900, 3, 1)) serial++;
        return CellValue.FromNumber(serial);
    }

    private CellValue Aggregate(FunctionNode function, object[] args)
    {
        var numberCount = 0; var count = 0; double sum = 0, product = 1, minimum = double.PositiveInfinity, maximum = double.NegativeInfinity;
        void Add(double number) { numberCount++; sum += number; product *= number; minimum = Math.Min(minimum, number); maximum = Math.Max(maximum, number); }
        for (var i = 0; i < args.Length; i++)
        foreach (var value in Values(args[i]))
        {
            check();
            if (function.Name == "COUNTA") { if (value.Kind != CellValueKind.Blank) count++; continue; }
            if (value.Kind == CellValueKind.Error) { if (function.Name == "COUNT") continue; return value; }
            if (value.Kind == CellValueKind.Number) { Add(value.NumberValue); continue; }
            if (function.Arguments[i] is ReferenceNode || value.Kind == CellValueKind.Blank) continue;
            if (function.Name == "COUNT")
            {
                if (value.Kind == CellValueKind.Boolean || value.Kind == CellValueKind.Text && double.TryParse(value.TextValue, NumberStyles.Float, CultureInfo.InvariantCulture, out _)) Add(0);
            }
            else Add(Number(value));
        }
        return function.Name switch
        {
            "COUNTA" => CellValue.FromNumber(count), "COUNT" => CellValue.FromNumber(numberCount),
            "SUM" => CellValue.FromNumber(sum),
            "AVERAGE" => numberCount == 0 ? Error(CellErrorCode.DivisionByZero) : CellValue.FromNumber(sum / numberCount),
            "MIN" => CellValue.FromNumber(numberCount == 0 ? 0 : minimum),
            "MAX" => CellValue.FromNumber(numberCount == 0 ? 0 : maximum),
            "PRODUCT" => CellValue.FromNumber(numberCount == 0 ? 0 : product),
            _ => Error(CellErrorCode.Name)
        };
    }

    private IEnumerable<CellValue> Values(object arg) => arg is Range range ? Stored(range) : [(CellValue)arg];
    private IEnumerable<CellValue> Stored(Range range)
    {
        if (!_positions.TryGetValue(range.Sheet.Id, out var positions))
            _positions[range.Sheet.Id] = positions = range.Sheet.Cells.Keys.OrderBy(p => p.Row).ThenBy(p => p.Column).ToArray();
        foreach (var p in positions)
        {
            check();
            if (p.Row > range.Reference.EndRow) break;
            if (range.Reference.Contains(p.Row, p.Column)) yield return Read(range.Sheet, p.Row, p.Column);
        }
    }

    private static (int Rows, int Columns) Shape(Range range)
    {
        var r = range.Reference;
        return (r.WholeColumns ? range.Sheet.UsedRowCount : r.EndRow - r.Row + 1, r.EndColumn - r.Column + 1);
    }

    private CellValue Index(object[] args)
    {
        if (args.Skip(1).Any(a => a is not CellValue) || args[0] is not Range range) return Error(CellErrorCode.Value);
        var (rows, columns) = Shape(range);
        var r = range.Reference;
        var row = checked((int)Math.Truncate(Number((CellValue)args[1]))) - 1;
        var column = args.Length == 3 ? checked((int)Math.Truncate(Number((CellValue)args[2]))) - 1 : 0;
        return row < 0 || row >= rows || column < 0 || column >= columns ? Error(CellErrorCode.Reference) : Read(range.Sheet, r.Row + row, r.Column + column);
    }

    private CellValue Script(FunctionNode function, ScriptFunctionDescriptor descriptor)
    {
        var prepared = PrepareScript(function, descriptor);
        if (prepared is CellValue error) return error;
        var call = (PreparedScriptCall)prepared;
        try
        {
            check(); var value = runtime!.Invoke(function.Name, call.Arguments); check();
            return CompleteScript(call, new(value, null));
        }
        catch (Exception e) when (e is not OperationCanceledException and not TimeoutException and not ScriptRuntimeFailureException and not FormulaValueException and not FormatException)
        { throw new FormulaValueException(CellErrorCode.Script, $"{function.Name}：{e.Message}"); }
    }

    internal sealed record PreparedScriptCall(ScriptFunctionDescriptor Descriptor, IReadOnlyList<object> Arguments);

    // Only a whole-cell call with no nested UDF can be prepared without executing script code.
    // IF/IFERROR, binary expressions and nested UDFs retain the existing lazy, serial evaluator.
    public bool CanBatch(FormulaNode node) => node is FunctionNode function && scripts.ContainsKey(function.Name)
        && !function.Arguments.SelectMany(FormulaSyntax.Walk).OfType<FunctionNode>().Any(f => scripts.ContainsKey(f.Name));

    public object PrepareScript(FormulaNode node)
    {
        var function = (FunctionNode)node;
        return PrepareScript(function, scripts[function.Name]);
    }

    private object PrepareScript(FunctionNode function, ScriptFunctionDescriptor descriptor)
    {
        if (function.Arguments.Count < descriptor.MinArguments || function.Arguments.Count > descriptor.MaxArguments)
            return Error(CellErrorCode.Value);
        var bridge = new List<object>();
        foreach (var expression in function.Arguments)
        {
            var arg = Evaluate(expression);
            if (arg is CellValue scalar)
            {
                if (scalar.Kind == CellValueKind.Error) return scalar;
                bridge.Add(ObjectValue(scalar, descriptor.ApiVersion)!);
            }
            else if (arg is Range range)
            {
                var (rows, columns) = Shape(range);
                if ((long)rows * columns > 1_000_000) Fail(CellErrorCode.Value, "JavaScript 范围参数最多支持 1,000,000 个位置。");
                var values = new List<object?>();
                for (var row = 0; row < rows; row++)
                for (var column = 0; column < columns; column++)
                {
                    var item = Read(range.Sheet, range.Reference.Row + row, range.Reference.Column + column);
                    if (item.Kind == CellValueKind.Error) return item;
                    values.Add(ObjectValue(item, descriptor.ApiVersion));
                }
                bridge.Add(descriptor.ApiVersion == 1 ? values : new { kind = "range", rows, columns, values });
            }
        }
        return new PreparedScriptCall(descriptor, bridge);
    }

    public CellValue CompleteScript(PreparedScriptCall call, ScriptCallResult reply)
    {
        if (reply.ValueTooLong) throw new FormatException("计算结果超过单元格文本长度限制。");
        try
        {
            if (reply.Error is not null) throw new InvalidOperationException(reply.Error);
            var result = reply.Value switch
            {
                null => CellValue.Blank, string s => CellValue.FromText(s), bool b => CellValue.FromBoolean(b),
                double n => CellValue.FromNumber(n), int n => CellValue.FromNumber(n), _ => Error(CellErrorCode.Script)
            };
            if (call.Descriptor.ReturnType != "any" && call.Descriptor.ReturnType != result.Kind.ToString().ToLowerInvariant())
                Fail(CellErrorCode.Script, "脚本返回类型不符合描述符。");
            return result;
        }
        catch (Exception e) when (e is not OperationCanceledException and not TimeoutException and not ScriptRuntimeFailureException)
        { throw new FormulaValueException(CellErrorCode.Script, $"{call.Descriptor.Name}：{e.Message}"); }
    }

    private static object? ObjectValue(CellValue value, int apiVersion) => value.Kind switch
    {
        CellValueKind.Blank => apiVersion == 1 ? "" : null, CellValueKind.Text => value.TextValue,
        CellValueKind.Number => value.NumberValue, CellValueKind.Boolean => value.BooleanValue,
        _ => throw new InvalidOperationException()
    };

    private static double Number(CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => 0, CellValueKind.Number => value.NumberValue,
        CellValueKind.Boolean => value.BooleanValue ? 1 : 0,
        CellValueKind.Text when double.TryParse(value.TextValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) && double.IsFinite(n) => n,
        _ => throw new FormulaValueException(CellErrorCode.Value, "值不能转换为有限数字。")
    };
    private static bool Boolean(CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => false, CellValueKind.Boolean => value.BooleanValue, CellValueKind.Number => value.NumberValue != 0,
        CellValueKind.Text when bool.TryParse(value.TextValue, out var b) => b,
        _ => throw new FormulaValueException(CellErrorCode.Value, "值不能转换为逻辑值。")
    };
    private static string Text(CellValue value) => CellInput.DisplayText(value);
    private static int Compare(CellValue a, CellValue b)
    {
        if (a.Kind == CellValueKind.Blank) a = b.Kind switch { CellValueKind.Text => CellValue.FromText(""), CellValueKind.Boolean => CellValue.FromBoolean(false), _ => CellValue.FromNumber(0) };
        if (b.Kind == CellValueKind.Blank) b = a.Kind switch { CellValueKind.Text => CellValue.FromText(""), CellValueKind.Boolean => CellValue.FromBoolean(false), _ => CellValue.FromNumber(0) };
        static int Rank(CellValueKind kind) => kind switch { CellValueKind.Number => 0, CellValueKind.Text => 1, CellValueKind.Boolean => 2, _ => 3 };
        if (a.Kind != b.Kind) return Rank(a.Kind).CompareTo(Rank(b.Kind));
        return a.Kind switch { CellValueKind.Number => a.NumberValue.CompareTo(b.NumberValue),
            CellValueKind.Boolean => a.BooleanValue.CompareTo(b.BooleanValue),
            _ => string.Compare(Text(a), Text(b), StringComparison.OrdinalIgnoreCase) };
    }
}
