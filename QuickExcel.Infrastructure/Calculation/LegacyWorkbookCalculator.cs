using System.Diagnostics;
using System.Globalization;
using ExcelIO;
using ExcelIO.Formula;
using QuickExcel.Domain;

using QuickExcel.Core;

namespace QuickExcel.Calculation;

/// <summary>
/// The 0.2 compatibility calculator: the existing parser and scalar builtins, evaluated against
/// immutable sparse cells. Cross-sheet formulas and their dependents keep blocked caches.
/// Scalar arithmetic and aggregate coercion retain the prototype's numeric-text rules; this is
/// not a complete Excel formula implementation. JavaScript ranges are flat row-major arrays
/// including blanks, and whole-column references end at the sheet's last stored row.
/// Cancellation is checked between operations; native JavaScript calls have their own one-second
/// deadline and cannot be interrupted by this token. A cancelled or timed-out round returns nothing.
/// </summary>
public sealed class LegacyWorkbookCalculator(IScriptRuntimeFactory runtimeFactory) : IWorkbookCalculator
{
    public static TimeSpan CalculationTimeLimit { get; } = TimeSpan.FromSeconds(5);
    /// <summary>Maximum number of positions, including blanks, in one JavaScript range argument.</summary>
    public const int MaxScriptRangeCells = 1_000_000;
    private static readonly IReadOnlyDictionary<string, ExcelFunction> Builtins = CreateBuiltins();
    private readonly IScriptRuntimeFactory _runtimeFactory = runtimeFactory ?? throw new ArgumentNullException(nameof(runtimeFactory));

    public CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Snapshot);
        if (request.ScriptsEnabled && !request.Snapshot.ScriptModules.CanExecuteLegacy)
            throw new NotSupportedException("多模块运行需要 M4 模块加载器。");
        var round = new Round(request, cancellationToken);
        round.Check();
        var previousCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            // Compilation exceptions escape; an incomplete round must never be submitted to the session.
            using var runtime = request.ScriptsEnabled ? _runtimeFactory.Create(request.Snapshot.ScriptSource) : null;
            round.Check();
            round.Runtime = runtime;
            round.ScriptFunctions = runtime?.FunctionNames.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];
            foreach (var sheet in request.Snapshot.Sheets) round.CalculateSheet(sheet);
            if (runtime is not null)
                foreach (var line in runtime.GetLogs()) round.AddDiagnostic($"JavaScript：{line}");
            round.Check();
            return new CalculationResult(request.Stamp, round.Results, round.Diagnostics);
        }
        finally { CultureInfo.CurrentCulture = previousCulture; }
    }

    private sealed class Round(CalculationRequest request, CancellationToken cancellationToken)
    {
        private readonly Stopwatch _timer = Stopwatch.StartNew();
        public IScriptRuntime? Runtime { get; set; }
        public HashSet<string> ScriptFunctions { get; set; } = [];
        public Dictionary<CellAddress, CellCalculation> Results { get; } = [];
        public List<string> Diagnostics { get; } = [];

        public void Check()
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (_timer.Elapsed > CalculationTimeLimit)
                throw new TimeoutException("重算达到 5 秒限制；本轮结果未提交，保留原有缓存。");
        }

        public void AddDiagnostic(string message)
        {
            if (Diagnostics.Count < 100 && !Diagnostics.Contains(message)) Diagnostics.Add(message);
        }

        public void CalculateSheet(WorksheetSnapshot sheet)
        {
            var nodes = new Dictionary<CellPosition, Node>();
            foreach (var (position, cell) in sheet.Cells)
            {
                Check();
                if (cell.Content.Kind != CellContentKind.Formula) continue;
                var node = new Node(cell);
                nodes.Add(position, node);
                if (HasCrossSheetReference(cell.Content.FormulaText))
                {
                    node.Blocked = true;
                    node.Diagnostics.Add("0.2 暂不重算跨工作表引用；保留缓存，结果可能已过期。");
                    continue;
                }
                try
                {
                    var formula = cell.Content.FormulaText;
                    if (formula.Length > 8_192 || formula.Count(c => c is '(' or '+' or '-' or '^' or '&') > 128)
                        throw new FormatException("公式过长或嵌套过深。");
                    node.Expression = FormulaParser.Parse(formula);
                    Visit(node.Expression, expression =>
                    {
                        if (expression is ErrorExpr) node.Diagnostics.Add("公式包含无法解析的表达式。");
                        if (expression is not FunctionExpr function || Builtins.ContainsKey(function.Name) || ScriptFunctions.Contains(function.Name)) return;
                        node.Blocked |= !request.ScriptsEnabled;
                        node.Diagnostics.Add(request.ScriptsEnabled ? $"未找到函数 {function.Name}。" : $"函数 {function.Name} 尚未启用；保留缓存，结果可能已过期。");
                    });
                }
                catch (Exception exception) when (exception is not OperationCanceledException and not TimeoutException)
                {
                    node.Expression = new ErrorExpr("#VALUE!");
                    node.Diagnostics.Add(exception.Message);
                }
            }

            // Index actual formula cells by row. Whole-column ranges do not allocate empty cells.
            var byRow = nodes.Keys.GroupBy(position => position.Row).ToDictionary(group => group.Key, group => group.ToArray());
            foreach (var (position, node) in nodes)
            {
                Check();
                var references = new HashSet<CellPosition>();
                if (node.Expression is not null) Visit(node.Expression, expression =>
                {
                    Check();
                    if (expression is CellRefExpr reference && nodes.ContainsKey(new(reference.Row, reference.Col)))
                        references.Add(new(reference.Row, reference.Col));
                    if (expression is not RangeRefExpr range) return;
                    foreach (var (row, candidates) in byRow)
                    {
                        Check();
                        if (row < range.StartRow || row > range.EndRow) continue;
                        foreach (var candidate in candidates)
                            if (candidate.Column >= range.StartCol && candidate.Column <= range.EndCol) references.Add(candidate);
                    }
                });
                node.Pending = references.Count;
                foreach (var reference in references) nodes[reference].Dependents.Add(position);
            }

            // Stable ordering also makes stateful script calls deterministic within a round.
            var ready = new PriorityQueue<CellPosition, (int Row, int Column)>();
            foreach (var (position, node) in nodes)
                if (node.Pending == 0) ready.Enqueue(position, (position.Row, position.Column));
            var evaluator = new SparseEvaluator(this, sheet);
            while (ready.TryDequeue(out var position, out _))
            {
                Check();
                var node = nodes[position];
                node.Processed = true;
                CellValue? value = node.Cell.Calculation?.Value;
                if (!node.Blocked)
                {
                    evaluator.Diagnostics.Clear();
                    try { value = evaluator.Scalar(node.Expression!); }
                    catch (Exception exception) when (exception is not OperationCanceledException and not TimeoutException)
                    {
                        value = CellValue.FromError(CellErrorCode.Value);
                        node.Diagnostics.Add(exception.Message);
                    }
                    node.Diagnostics.AddRange(evaluator.Diagnostics);
                    if (value is { Kind: CellValueKind.Text } text && text.TextValue.Length > WorkbookLimits.MaxCellTextLength)
                    {
                        value = CellValue.FromError(CellErrorCode.Value);
                        node.Diagnostics.Add("计算结果超过单元格文本长度限制。");
                    }
                }
                Store(sheet, position, node, value, node.Blocked ? CalculationState.Blocked : value?.Kind == CellValueKind.Error ? CalculationState.Error : CalculationState.Fresh);
                foreach (var dependent in node.Dependents)
                {
                    var next = nodes[dependent];
                    if (node.Blocked && !next.Blocked)
                    {
                        next.Blocked = true;
                        next.Diagnostics.Add("依赖的公式无法重算；保留缓存，结果可能已过期。");
                    }
                    if (--next.Pending == 0) ready.Enqueue(dependent, (dependent.Row, dependent.Column));
                }
            }
            foreach (var (position, node) in nodes)
            {
                Check();
                if (node.Processed) continue;
                node.Diagnostics.Add("存在循环引用或依赖于循环引用。");
                Store(sheet, position, node, CellValue.FromError(CellErrorCode.Cycle), CalculationState.Error);
            }
        }

        private void Store(WorksheetSnapshot sheet, CellPosition position, Node node, CellValue? value, CalculationState state)
        {
            foreach (var message in node.Diagnostics) AddDiagnostic($"{sheet.Name}!{Address(position)}：{message}");
            var stamp = request.Stamp;
            Results[new(sheet.Id, position.Row, position.Column)] = new CellCalculation(value, state,
                stamp.InputRevision, stamp.CalculationGeneration, stamp.ScriptGeneration, node.Diagnostics);
        }
    }

    private sealed class SparseEvaluator(Round round, WorksheetSnapshot sheet)
    {
        private CellPosition[]? _orderedPositions;
        public List<string> Diagnostics { get; } = [];

        public CellValue Scalar(Expr expression, int depth = 0) => Evaluate(expression, depth) is CellValue value
            ? value : CellValue.FromError(CellErrorCode.Value);

        private object Evaluate(Expr expression, int depth)
        {
            round.Check();
            if (depth > 128) throw new FormatException("公式嵌套过深。");
            return expression switch
            {
                NumberExpr number => CellValue.FromNumber(number.Value),
                TextExpr text => CellValue.FromText(text.Value),
                BoolExpr boolean => CellValue.FromBoolean(boolean.Value),
                ErrorExpr error => CellValue.FromError(ParseError(error.Error)),
                CellRefExpr reference => Read(reference.Row, reference.Col),
                RangeRefExpr range => new RangeValue(range),
                UnaryExpr unary => Unary(unary, depth + 1),
                BinaryExpr binary => Binary(binary, depth + 1),
                FunctionExpr function => Function(function, depth + 1),
                _ => CellValue.FromError(CellErrorCode.Value)
            };
        }

        private CellValue Read(int row, int column)
        {
            if (row < 0 || column < 0 || row >= WorkbookLimits.MaxRows || column >= WorkbookLimits.MaxColumns)
                return CellValue.FromError(CellErrorCode.Reference);
            if (round.Results.TryGetValue(new(sheet.Id, row, column), out var result)) return result.Value ?? CellValue.Blank;
            return sheet.Cells.GetValueOrDefault(new(row, column))?.Value ?? CellValue.Blank;
        }

        private CellValue Unary(UnaryExpr unary, int depth)
        {
            var value = Scalar(unary.Operand, depth);
            if (value.Kind == CellValueKind.Error) return value;
            var number = Number(value);
            return CellValue.FromNumber(unary.Op switch { UnaryOp.Neg => -number, UnaryOp.Percent => number / 100, _ => number });
        }

        private CellValue Binary(BinaryExpr binary, int depth)
        {
            var left = Scalar(binary.Left, depth);
            var right = Scalar(binary.Right, depth);
            if (left.Kind == CellValueKind.Error) return left;
            if (right.Kind == CellValueKind.Error) return right;
            if (binary.Op == BinaryOp.Concat) return CellValue.FromText(Text(left) + Text(right));
            if (binary.Op is BinaryOp.Eq or BinaryOp.Ne or BinaryOp.Lt or BinaryOp.Le or BinaryOp.Gt or BinaryOp.Ge)
            {
                var comparison = Compare(left, right);
                return CellValue.FromBoolean(binary.Op switch
                {
                    BinaryOp.Eq => comparison == 0, BinaryOp.Ne => comparison != 0,
                    BinaryOp.Lt => comparison < 0, BinaryOp.Le => comparison <= 0,
                    BinaryOp.Gt => comparison > 0, _ => comparison >= 0
                });
            }
            var first = Number(left);
            var second = Number(right);
            if (binary.Op == BinaryOp.Div && second == 0) return CellValue.FromError(CellErrorCode.DivisionByZero);
            return CellValue.FromNumber(binary.Op switch
            {
                BinaryOp.Add => first + second, BinaryOp.Sub => first - second,
                BinaryOp.Mul => first * second, BinaryOp.Div => first / second,
                BinaryOp.Pow => Math.Pow(first, second), _ => double.NaN
            });
        }

        private CellValue Function(FunctionExpr function, int depth)
        {
            var isScript = round.ScriptFunctions.Contains(function.Name);
            var builtin = Builtins.GetValueOrDefault(function.Name);
            if (!isScript && builtin is null) return CellValue.FromError(CellErrorCode.Name);
            if (!isScript && (function.Arguments.Count < builtin!.MinArgs || function.Arguments.Count > builtin.MaxArgs))
                return CellValue.FromError(CellErrorCode.Value);
            if (!isScript && function.Name == "IF")
            {
                var condition = Scalar(function.Arguments[0], depth);
                return condition.Kind == CellValueKind.Error ? condition : Scalar(function.Arguments[Boolean(condition) ? 1 : 2], depth);
            }
            if (!isScript && function.Name == "IFERROR")
            {
                var value = Scalar(function.Arguments[0], depth);
                return value.Kind == CellValueKind.Error ? Scalar(function.Arguments[1], depth) : value;
            }
            var arguments = function.Arguments.Select(argument => Evaluate(argument, depth)).ToArray();
            foreach (var argument in arguments)
                if (argument is CellValue { Kind: CellValueKind.Error } error) return error;
            if (!isScript && function.Name is "VLOOKUP" or "MATCH" or "INDEX") return Lookup(function.Name, arguments);
            var bridge = new List<object>();
            foreach (var argument in arguments)
            {
                if (argument is CellValue value) bridge.Add(ToObject(value));
                else if (argument is RangeValue range)
                {
                    var values = new List<object>();
                    foreach (var item in isScript ? RectangularValues(range.Expression) : StoredValues(range.Expression))
                    {
                        if (item.Kind == CellValueKind.Error) return item;
                        values.Add(ToObject(item));
                    }
                    bridge.Add(values);
                }
            }
            if (isScript)
            {
                try
                {
                    round.Check();
                    var result = round.Runtime!.Invoke(function.Name, bridge);
                    round.Check();
                    return FromObject(result);
                }
                catch (Exception exception) when (exception is not OperationCanceledException and not TimeoutException)
                {
                    Diagnostics.Add($"{function.Name}：{exception.Message}");
                    return CellValue.FromError(CellErrorCode.Script);
                }
            }
            return FromObject(builtin!.Delegate(bridge.AsReadOnly(), ScalarBuiltinContext.Instance));
        }

        private IEnumerable<CellValue> StoredValues(RangeRefExpr range)
        {
            _orderedPositions ??= sheet.Cells.Keys.OrderBy(position => position.Row).ThenBy(position => position.Column).ToArray();
            foreach (var position in _orderedPositions)
            {
                round.Check();
                if (position.Row > range.EndRow) break;
                if (position.Row >= range.StartRow && position.Row <= range.EndRow && position.Column >= range.StartCol && position.Column <= range.EndCol)
                    yield return Read(position.Row, position.Column);
            }
        }

        private IEnumerable<CellValue> RectangularValues(RangeRefExpr range)
        {
            var firstRow = range.StartRow == int.MinValue ? 0 : range.StartRow;
            var lastRow = range.EndRow == int.MaxValue ? sheet.UsedRowCount - 1 : range.EndRow;
            if (firstRow < 0 || lastRow >= WorkbookLimits.MaxRows || range.StartCol < 0 || range.EndCol >= WorkbookLimits.MaxColumns)
                throw new FormatException("JavaScript 范围引用超出工作表支持的坐标。");
            var rows = Math.Max(0L, (long)lastRow - firstRow + 1);
            var columns = Math.Max(0L, (long)range.EndCol - range.StartCol + 1);
            if (rows * columns > MaxScriptRangeCells)
                throw new FormatException($"JavaScript 范围参数最多支持 {MaxScriptRangeCells:N0} 个位置（包括空单元格）。");
            for (var row = firstRow; row <= lastRow; row++)
            for (var column = range.StartCol; column <= range.EndCol; column++)
            {
                round.Check();
                yield return Read(row, column);
            }
        }

        private CellValue Lookup(string name, object[] arguments)
        {
            if (name == "INDEX")
            {
                if (arguments[0] is not RangeValue table) return CellValue.FromError(CellErrorCode.Value);
                var row = checked((int)Number((CellValue)arguments[1])) - 1;
                var column = arguments.Length > 2 ? checked((int)Number((CellValue)arguments[2])) - 1 : 0;
                var startRow = Math.Max(0, table.Expression.StartRow);
                return row < 0 || column < 0 || row > (long)table.Expression.EndRow - startRow || column > (long)table.Expression.EndCol - table.Expression.StartCol
                    ? CellValue.FromError(CellErrorCode.Reference) : Read(startRow + row, table.Expression.StartCol + column);
            }
            if (arguments[1] is not RangeValue range) return CellValue.FromError(CellErrorCode.Value);
            var sought = (CellValue)arguments[0];
            var firstRow = Math.Max(0, range.Expression.StartRow);
            var lastRow = Math.Min(sheet.UsedRowCount - 1, range.Expression.EndRow);
            var firstColumn = Math.Max(0, range.Expression.StartCol);
            var lastColumn = Math.Min(WorkbookLimits.MaxColumns - 1, range.Expression.EndCol);
            var returnColumn = name == "VLOOKUP" ? checked((int)Number((CellValue)arguments[2])) - 1 : 0;
            if (returnColumn < 0 || returnColumn > lastColumn - firstColumn) return CellValue.FromError(CellErrorCode.Reference);
            var matchMode = name == "VLOOKUP" ? (arguments.Length < 4 || Boolean((CellValue)arguments[3]) ? 1 : 0)
                : arguments.Length < 3 ? 1 : Math.Sign(Number((CellValue)arguments[2]));
            CellValue match = CellValue.FromError(CellErrorCode.NotAvailable);
            if (name == "MATCH" && lastRow > firstRow && lastColumn > firstColumn) return CellValue.FromError(CellErrorCode.Value);
            for (var row = firstRow; row <= lastRow; row++)
            for (var column = firstColumn; column <= (name == "VLOOKUP" ? firstColumn : lastColumn); column++)
            {
                round.Check();
                var candidate = Read(row, column);
                if (candidate.Kind == CellValueKind.Error) return candidate;
                var comparison = Compare(candidate, sought);
                if (comparison == 0 || (matchMode > 0 && comparison < 0) || (matchMode < 0 && comparison > 0))
                {
                    match = name == "VLOOKUP" ? Read(row, firstColumn + returnColumn) : CellValue.FromNumber(lastRow == firstRow ? column - firstColumn + 1 : row - firstRow + 1);
                    if (comparison == 0) return match;
                }
            }
            return match;
        }
    }

    private static double Number(CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => 0,
        CellValueKind.Number => value.NumberValue,
        CellValueKind.Boolean => value.BooleanValue ? 1 : 0,
        CellValueKind.Text when value.TextValue.Length == 0 => 0,
        CellValueKind.Text => double.Parse(value.TextValue, CultureInfo.InvariantCulture),
        _ => throw new FormatException("该值不能转换为数字。")
    };

    private static bool Boolean(CellValue value) => value.Kind switch
    {
        CellValueKind.Boolean => value.BooleanValue,
        CellValueKind.Number => value.NumberValue != 0,
        CellValueKind.Text => double.TryParse(value.TextValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? number != 0 : string.Equals(value.TextValue, "TRUE", StringComparison.OrdinalIgnoreCase),
        _ => false
    };

    private static string Text(CellValue value) => value.Kind switch
    {
        CellValueKind.Text => value.TextValue,
        CellValueKind.Number => value.NumberValue.ToString("G15", CultureInfo.InvariantCulture),
        CellValueKind.Boolean => value.BooleanValue ? "TRUE" : "FALSE",
        _ => ""
    };

    private static object ToObject(CellValue value) => value.Kind switch
    {
        CellValueKind.Number => value.NumberValue,
        CellValueKind.Boolean => value.BooleanValue,
        CellValueKind.Text => value.TextValue,
        _ => ""
    };

    private static CellValue FromObject(object? value) => value switch
    {
        null => CellValue.Blank,
        string text => CellValue.FromText(text),
        bool boolean => CellValue.FromBoolean(boolean),
        double number => CellValue.FromNumber(number),
        int number => CellValue.FromNumber(number),
        _ => throw new InvalidOperationException("函数返回了不支持的结果类型。")
    };

    private static int Compare(CellValue left, CellValue right) => left.Kind == CellValueKind.Number && right.Kind == CellValueKind.Number
        ? left.NumberValue.CompareTo(right.NumberValue) : string.Compare(Text(left), Text(right), StringComparison.OrdinalIgnoreCase);

    private static CellErrorCode ParseError(string error) => error switch
    {
        "#NULL!" => CellErrorCode.Null, "#DIV/0!" => CellErrorCode.DivisionByZero,
        "#REF!" => CellErrorCode.Reference, "#NAME?" => CellErrorCode.Name,
        "#NUM!" => CellErrorCode.Number, "#N/A" => CellErrorCode.NotAvailable,
        "#CYCLE!" => CellErrorCode.Cycle, "#SCRIPT!" => CellErrorCode.Script, _ => CellErrorCode.Value
    };

    private static void Visit(Expr expression, Action<Expr> action)
    {
        var pending = new Stack<Expr>();
        pending.Push(expression);
        while (pending.TryPop(out var current))
        {
            if (current is null) continue;
            action(current);
            switch (current)
            {
                case BinaryExpr binary: pending.Push(binary.Left); pending.Push(binary.Right); break;
                case UnaryExpr unary: pending.Push(unary.Operand); break;
                case FunctionExpr function: foreach (var argument in function.Arguments) pending.Push(argument); break;
            }
        }
    }

    private static bool HasCrossSheetReference(string formula)
    {
        var quoted = false;
        for (var index = 0; index < formula.Length; index++)
        {
            var character = formula[index];
            if (character == '"') quoted = !quoted;
            // Error spellings contain '!'; they are not sheet references, even if the legacy
            // parser cannot parse an error literal in this position.
            if (!quoted && character == '#')
            {
                string[] errors = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#CYCLE!", "#SCRIPT!"];
                var error = errors.FirstOrDefault(error => formula.AsSpan(index).StartsWith(error, StringComparison.OrdinalIgnoreCase));
                if (error is not null) { index += error.Length - 1; continue; }
            }
            if (!quoted && character == '!') return true;
        }
        return false;
    }

    private static IReadOnlyDictionary<string, ExcelFunction> CreateBuiltins()
    {
        var engine = new FormulaEngine();
        if (XlCell.OnValueChanged is { } callbacks)
            foreach (var callback in callbacks.GetInvocationList())
                if (ReferenceEquals(callback.Target, engine)) XlCell.OnValueChanged -= (Action<XlCell>)callback;
        return engine.Functions.GetAll().ToDictionary(function => function.Name, StringComparer.OrdinalIgnoreCase);
    }

    private static string Address(CellPosition position)
    {
        var letters = "";
        for (var value = position.Column + 1; value > 0; value = (value - 1) / 26) letters = (char)('A' + (value - 1) % 26) + letters;
        return letters + (position.Row + 1).ToString(CultureInfo.InvariantCulture);
    }

    private sealed class Node(CellSnapshot cell)
    {
        public CellSnapshot Cell { get; } = cell;
        public Expr? Expression { get; set; }
        public bool Blocked { get; set; }
        public bool Processed { get; set; }
        public int Pending { get; set; }
        public List<CellPosition> Dependents { get; } = [];
        public List<string> Diagnostics { get; } = [];
    }

    private sealed record RangeValue(RangeRefExpr Expression);

    // Only context-free legacy delegates reach this object; lookups use sparse cells above.
    private sealed class ScalarBuiltinContext : IFormulaContext
    {
        public static ScalarBuiltinContext Instance { get; } = new();
        public XlWorksheet Worksheet => throw new NotSupportedException("This builtin requires a worksheet context.");
        public XlCell? GetCell(int row, int col) => throw new NotSupportedException();
        public XlWorksheet? GetSheet(string name) => throw new NotSupportedException();
    }
}

