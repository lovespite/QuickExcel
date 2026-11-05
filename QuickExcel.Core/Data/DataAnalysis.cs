using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed record AnalysisScenario(string Name, IReadOnlyList<double> Values);
public sealed record GoalSeekResult(double Input, double Result, int Evaluations, CleaningPreview Preview);
public enum ForecastMethod { Linear, Holt }

/// <summary>Analysis runs in immutable private drafts; no intermediate result enters document history or calculator reuse.</summary>
public static class DataAnalysis
{
    internal static CleaningPreview Output(WorkbookSnapshot book, SheetId sheetId, CellPosition start, IReadOnlyList<IReadOnlyList<CellValue>> values, params string[] notices)
    {
        if (values.Count == 0 || values[0].Count == 0 || values.Any(r => r.Count != values[0].Count)) throw new ArgumentException("输出表为空或列数不一致。");
        var range = new CellRectangle(start.Row, start.Column, checked(start.Row + values.Count - 1), checked(start.Column + values[0].Count - 1));
        if (range.Count > DataCleaning.MaxCells) throw new ArgumentException("分析输出超过 20,000 格。");
        var sheet = DataCleaning.RequireRange(book, sheetId, range);
        var edits = new List<CellEdit>(); var overwrites = new List<CellAddress>();
        for (var r = 0; r < values.Count; r++)
        for (var c = 0; c < values[r].Count; c++)
        {
            var row = start.Row + r; var col = start.Column + c; var address = new CellAddress(sheetId, row, col);
            var old = sheet.GetCell(row, col);
            var content = CellContent.FromLiteral(values[r][c], sheet.GetStyleId(row, col));
            if (old is not null && (old.Content.Kind == CellContentKind.Formula || old.Value.Kind != CellValueKind.Blank) && old.Content != content) overwrites.Add(address);
            edits.Add(new(address, content));
        }
        return new(book, new SetCellsCommand(edits).Prepare(book), overwrites, [], notices);
    }

    private static void RequireInput(WorkbookSnapshot book, CellAddress address)
    {
        var sheet = DataCleaning.RequireRange(book, address.SheetId, new(address.Row, address.Column, address.Row, address.Column));
        var cell = sheet.GetCell(address.Row, address.Column);
        if (cell?.SpillOwner is not null) throw new ArgumentException("可变单元格不能是动态数组的溢出结果。");
        if (cell is not null && (cell.Content.Kind != CellContentKind.Literal || cell.Value.Kind is not (CellValueKind.Number or CellValueKind.Blank)))
            throw new ArgumentException("可变单元格须为数值或空白，不能覆盖公式或文本。");
        if (!ProtectionSafety.CanEdit(book, sheet, address.Row, address.Column)) throw new InvalidOperationException("可变单元格已锁定，请先取消工作表保护。");
    }

    private sealed class Evaluator(WorkbookSnapshot book, IWorkbookCalculator calculator, CellAddress formula, CancellationToken cancellation)
    {
        private readonly Guid _session = Guid.NewGuid();
        public int Count { get; private set; }
        public double Evaluate(IReadOnlyList<CellAddress> inputs, IReadOnlyList<double> values)
        {
            cancellation.ThrowIfCancellationRequested();
            if (book.GetSheet(formula.SheetId).GetCell(formula.Row, formula.Column)?.Content.Kind != CellContentKind.Formula)
                throw new ArgumentException("目标单元格必须包含公式。");
            if (++Count > 400) throw new InvalidOperationException("模拟分析超过 400 次计算限制，未修改工作簿。");
            var draft = new SetCellsCommand(inputs.Select((address, i) => new CellEdit(address,
                CellContent.FromLiteral(CellValue.FromNumber(values[i]), book.GetSheet(address.SheetId).GetStyleId(address.Row, address.Column))))).Prepare(book).Snapshot;
            var stamp = new CalculationStamp(_session, book.Revision, Count, 0);
            var result = calculator.Calculate(new(stamp, draft, false) { IsProvisional = true, ForceFull = true, CancellationToken = cancellation }, cancellation);
            if (!result.Cells.TryGetValue(formula, out var cell) || cell.State != CalculationState.Fresh || cell.Value is not { Kind: CellValueKind.Number } value || !double.IsFinite(value.NumberValue))
                throw new InvalidOperationException("模拟公式未得到有效数值。请检查错误、循环引用或脚本函数；模拟分析仅使用内置公式。");
            return value.NumberValue;
        }
    }

    public static GoalSeekResult GoalSeek(WorkbookSnapshot book, IWorkbookCalculator calculator, CellAddress formula, CellAddress input,
        double target, double tolerance = 0.000001, CancellationToken cancellation = default)
    {
        RequireInput(book, input);
        if (!double.IsFinite(target) || !double.IsFinite(tolerance) || tolerance <= 0) throw new ArgumentException("目标值和误差必须为有限数，误差须大于零。");
        var evaluator = new Evaluator(book, calculator, formula, cancellation);
        double F(double x) => evaluator.Evaluate([input], [x]) - target;
        var x = book.GetSheet(input.SheetId).GetCell(input.Row, input.Column)?.Value is { Kind: CellValueKind.Number } initial ? initial.NumberValue : 0;
        var y = F(x); var previousX = x; var previousY = y;
        // Secant iteration with bounded steps. A bracket, once found, is kept and bisected when secant leaves it.
        double? low = null, high = null; double lowY = 0;
        for (var iteration = 0; iteration < 100 && Math.Abs(y) > tolerance; iteration++)
        {
            cancellation.ThrowIfCancellationRequested();
            double next;
            if (iteration == 0 || y == previousY) next = x + Math.Max(0.01, Math.Abs(x) * 0.01);
            else next = x - y * (x - previousX) / (y - previousY);
            if (low is not null && high is not null && (!double.IsFinite(next) || next <= low || next >= high)) next = low.Value + (high.Value - low.Value) / 2;
            if (!double.IsFinite(next) || Math.Abs(next) > 1e100) break;
            if (next == x) break;
            var nextY = F(next);
            if (low is not null)
            {
                if (Math.Sign(nextY) == Math.Sign(lowY)) { low = next; lowY = nextY; } else high = next;
            }
            else if (Math.Sign(y) != Math.Sign(nextY))
            { low = Math.Min(x, next); high = Math.Max(x, next); lowY = x < next ? y : nextY; }
            previousX = x; previousY = y; x = next; y = nextY;
        }
        if (Math.Abs(y) > tolerance) throw new InvalidOperationException("单变量求解未收敛。请调整初始值或目标值；原始数据保持不变。");
        var preview = Output(book, input.SheetId, new(input.Row, input.Column), [[CellValue.FromNumber(x)]],
            $"单变量求解：可变值 {x:G12}；公式结果 {y + target:G12}；目标 {target:G12}；误差 {Math.Abs(y):G4}；计算 {evaluator.Count} 次。");
        return new(x, y + target, evaluator.Count, preview);
    }

    public static CleaningPreview Scenarios(WorkbookSnapshot book, IWorkbookCalculator calculator, CellAddress formula, IReadOnlyList<CellAddress> inputs,
        IReadOnlyList<AnalysisScenario> scenarios, SheetId outputSheet, CellPosition output, CancellationToken cancellation = default)
    {
        if (inputs.Count is < 1 or > 16 || inputs.Distinct().Count() != inputs.Count || scenarios.Count is < 1 or > 200 ||
            scenarios.Any(s => s.Values.Count != inputs.Count || s.Values.Any(v => !double.IsFinite(v)) || string.IsNullOrWhiteSpace(s.Name)))
            throw new ArgumentException("请提供 1–16 个不同的可变单元格和 1–200 个完整数值方案。");
        foreach (var input in inputs) RequireInput(book, input);
        var evaluator = new Evaluator(book, calculator, formula, cancellation);
        var table = new List<IReadOnlyList<CellValue>> { new[] { CellValue.FromText("方案") }.Concat(inputs.Select(a => CellValue.FromText($"{book.GetSheet(a.SheetId).Name}!{new CellRectangle(a.Row, a.Column, a.Row, a.Column)}"))).Append(CellValue.FromText("结果")).ToArray() };
        foreach (var scenario in scenarios)
        {
            var value = evaluator.Evaluate(inputs, scenario.Values);
            table.Add(new[] { CellValue.FromText(scenario.Name) }.Concat(scenario.Values.Select(CellValue.FromNumber)).Append(CellValue.FromNumber(value)).ToArray());
        }
        RequireSeparateOutput(formula, inputs, outputSheet, output, table.Count, table[0].Count);
        return Output(book, outputSheet, output, table, "方案摘要为本次计算的值快照；源输入和公式保持原值。源数据改变后请重新运行。");
    }

    public static CleaningPreview DataTable(WorkbookSnapshot book, IWorkbookCalculator calculator, CellAddress formula,
        CellAddress columnInput, IReadOnlyList<double> columnValues, CellAddress? rowInput, IReadOnlyList<double> rowValues,
        SheetId outputSheet, CellPosition output, CancellationToken cancellation = default)
    {
        RequireInput(book, columnInput);
        if (rowInput is { } second) RequireInput(book, second);
        if (rowInput == columnInput || columnValues.Count == 0 || columnValues.Count > 200 || rowInput is not null && rowValues.Count == 0 ||
            columnValues.Count * (long)Math.Max(1, rowValues.Count) > 200 || columnValues.Concat(rowValues).Any(v => !double.IsFinite(v)))
            throw new ArgumentException("模拟运算表需要不同的输入单元格、有限数值和最多 200 种组合。");
        var evaluator = new Evaluator(book, calculator, formula, cancellation);
        var table = new List<IReadOnlyList<CellValue>>();
        if (rowInput is null)
        {
            table.Add([CellValue.FromText("输入"), CellValue.FromText("结果")]);
            foreach (var value in columnValues) table.Add([CellValue.FromNumber(value), CellValue.FromNumber(evaluator.Evaluate([columnInput], [value]))]);
        }
        else
        {
            table.Add(new[] { CellValue.FromText("列输入 / 行输入") }.Concat(rowValues.Select(CellValue.FromNumber)).ToArray());
            foreach (var column in columnValues)
                table.Add(new[] { CellValue.FromNumber(column) }.Concat(rowValues.Select(row => CellValue.FromNumber(evaluator.Evaluate([columnInput, rowInput.Value], [column, row])))).ToArray());
        }
        RequireSeparateOutput(formula, rowInput is null ? [columnInput] : [columnInput, rowInput.Value], outputSheet, output, table.Count, table[0].Count);
        return Output(book, outputSheet, output, table, "模拟运算表将保存为结果值，源输入保持原值。");
    }

    private static void RequireSeparateOutput(CellAddress formula, IReadOnlyList<CellAddress> inputs, SheetId sheet, CellPosition start, int rows, int columns)
    {
        var range = new CellRectangle(start.Row, start.Column, checked(start.Row + rows - 1), checked(start.Column + columns - 1));
        if (inputs.Append(formula).Any(a => a.SheetId == sheet && range.Contains(a.Row, a.Column))) throw new ArgumentException("输出区域不能覆盖可变单元格或目标公式。");
    }

    public static CleaningPreview Forecast(WorkbookSnapshot book, SheetId sheetId, CellRectangle source, bool hasHeader, int periods,
        CellPosition output, ForecastMethod method = ForecastMethod.Linear, double alpha = 0.5, double beta = 0.3)
    {
        var sheet = DataCleaning.RequireRange(book, sheetId, source);
        if (source.Columns != 2 || source.Rows - (hasHeader ? 1 : 0) < 3 || source.Rows > 5000 || periods is < 1 or > 1000 ||
            !Enum.IsDefined(method) || !double.IsFinite(alpha) || !double.IsFinite(beta) || alpha is <= 0 or > 1 || beta is <= 0 or > 1)
            throw new ArgumentException("预测需要两列时间/数值及至少 3 个数据点；最多 5,000 行，预测 1–1,000 期，平滑参数在 (0,1] 内。");
        double Number(int row, int col)
        {
            var cell = sheet.GetCell(row, col);
            if (cell?.Value is not { Kind: CellValueKind.Number } v || cell.Content.Kind == CellContentKind.Formula &&
                (cell.Calculation is not { State: CalculationState.Fresh } || cell.Calculation.InputRevision != book.Revision))
                throw new ArgumentException($"第 {row + 1} 行需要有效的时间和数值；不自动跳过缺失点或旧缓存。");
            return v.NumberValue;
        }
        var pairs = Enumerable.Range(source.Top + (hasHeader ? 1 : 0), source.Rows - (hasHeader ? 1 : 0)).Select(r => (X: Number(r, source.Left), Y: Number(r, source.Right))).ToArray();
        var step = pairs[1].X - pairs[0].X;
        if (step <= 0 || !double.IsFinite(step) || pairs.Zip(pairs.Skip(1)).Any(p => Math.Abs(p.Second.X - p.First.X - step) > Math.Max(1e-9, Math.Abs(step) * 1e-6)))
            throw new ArgumentException("时间轴必须严格递增且等间距；请先补齐缺失期并处理重复时间。");
        var meanX = pairs.Average(p => p.X); var meanY = pairs.Average(p => p.Y);
        var slope = pairs.Sum(p => (p.X - meanX) * (p.Y - meanY)) / pairs.Sum(p => (p.X - meanX) * (p.X - meanX));
        var level = pairs[0].Y; var trend = pairs[1].Y - pairs[0].Y;
        if (method == ForecastMethod.Holt)
            foreach (var p in pairs.Skip(1)) { var prior = level; level = alpha * p.Y + (1 - alpha) * (level + trend); trend = beta * (level - prior) + (1 - beta) * trend; }
        var table = new List<IReadOnlyList<CellValue>> { new[] { "时间", "历史值", "预测值" }.Select(CellValue.FromText).ToArray() };
        foreach (var p in pairs) table.Add([CellValue.FromNumber(p.X), CellValue.FromNumber(p.Y), CellValue.Blank]);
        for (var i = 1; i <= periods; i++)
        {
            var x = pairs[^1].X + i * step; var y = method == ForecastMethod.Linear ? meanY + slope * (x - meanX) : level + i * trend;
            if (!double.IsFinite(x) || !double.IsFinite(y)) throw new InvalidOperationException("预测值超出数值范围。");
            table.Add([CellValue.FromNumber(x), CellValue.Blank, CellValue.FromNumber(y)]);
        }
        var destination = new CellRectangle(output.Row, output.Column, checked(output.Row + table.Count - 1), checked(output.Column + 2));
        if (ValidationSafety.Overlaps(source, destination)) throw new ArgumentException("预测输出不能覆盖历史数据区域。");
        var preview = Output(book, sheetId, output, table, $"{(method == ForecastMethod.Linear ? "线性趋势" : "Holt 双指数平滑（无季节性）")}预测 {periods} 期；结果为值快照。时间间隔 {step:G8}；不提供置信区间或自动季节性推断。");
        // Carry the time axis format (including Excel date serial formatting) into the report.
        var format = sheet.GetStyleId(source.Top + (hasHeader ? 1 : 0), source.Left);
        var edits = preview.Changes.Select(c => new CellEdit(c.Address, c.Address.Column == output.Column && c.Address.Row > output.Row ? c.After.WithStyle(format) : c.After));
        return new(book, new SetCellsCommand(edits).Prepare(book), preview.Overwrites, [], preview.Notices);
    }
}
