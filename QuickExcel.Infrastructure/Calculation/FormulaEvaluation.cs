using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

public sealed record FormulaEvaluationStep(string Expression, string Value);
public sealed record FormulaEvaluationResult(IReadOnlyList<FormulaEvaluationStep> Steps, string? Notice);

/// <summary>Observes the production evaluator on a private snapshot. Never calculates dependencies or starts scripts.</summary>
public static class FormulaEvaluation
{
    public static FormulaEvaluationResult Inspect(WorkbookSnapshot snapshot, CellAddress origin, CancellationToken cancellationToken = default)
    {
        var steps = new List<FormulaEvaluationStep>();
        var sheet = snapshot.GetSheet(origin.SheetId); var cell = sheet.GetCell(origin.Row, origin.Column);
        if (cell?.Content.Kind != CellContentKind.Formula || ProtectionSafety.HidesFormula(snapshot, sheet, origin.Row, origin.Column))
            return new([], MessageText.Get("当前单元格没有可显示的公式。"));
        if (snapshot.FileState?.FormulaReadOnly == true)
            return new([], MessageText.Get("来源公式语义未完整支持；仅显示已有计算状态。"));
        if (cell.Calculation is { Value.Kind: CellValueKind.Error, Value.ErrorCode: CellErrorCode.Cycle })
            return new([], MessageText.Get("循环依赖无法逐步求值；请先修正引用。"));
        var work = 0;
        void Check() { cancellationToken.ThrowIfCancellationRequested(); if (++work > 1_000_000) throw new InvalidOperationException(MessageText.Get("逐步求值超过分析预算。")); }
        try
        {
            Check();
            var node = FormulaBinding.Bind(FormulaSyntax.Parse(cell.Content.FormulaText), snapshot, sheet.Id, new(origin.Row, origin.Column), cell.Content.DynamicArray);
            var functions = FormulaSyntax.Walk(node).OfType<FunctionNode>().ToArray();
            if (functions.Any(f => !StandardFunctions.Signatures.ContainsKey(f.Name) || f.Name is "NOW" or "TODAY"))
                return new([], MessageText.Get("脚本、未知函数和时钟函数仅显示当前计算状态，不重新执行。"));
            var inputs = new Dictionary<CellAddress, CellCalculation>();
            foreach (var s in snapshot.Sheets)
            foreach (var (position, input) in s.Cells)
            {
                Check();
                if (input.Calculation is { Value: not null, State: CalculationState.Fresh or CalculationState.Error } current && current.InputRevision == snapshot.Revision)
                    inputs[new(s.Id, position.Row, position.Column)] = current;
            }
            var evaluator = new TypedEvaluator(snapshot, inputs, null, new Dictionary<string, ScriptFunctionDescriptor>(), Check)
            {
                CurrentSheet = sheet.Id, RequireCurrentInputs = true,
                ArrayMode = cell.Content.DynamicArray || DynamicArrayFunctions.UsesArrays(node),
                Trace = (part, value) =>
                {
                    Check();
                    if (steps.Count >= 500) throw new InvalidOperationException(MessageText.Get("逐步求值超过分析预算。"));
                    if (part is not ValueNode)
                        steps.Add(new(FormulaSyntax.TryPrintPortable(part, out var expression) ? expression : MessageText.Get("展开后的子表达式"), value));
                }
            };
            object result;
            try { result = evaluator.ArrayMode ? evaluator.FormulaResult(node) : evaluator.CellResult(node); }
            catch (Exception ex) when (ex is FormulaValueException or FormatException or OverflowException or ArgumentException)
            { result = CellValue.FromError(ex is FormulaValueException error ? error.Code : CellErrorCode.Value); }
            steps.Add(new(MessageText.Get("公式结果（数组溢出布局另行检查）"), TypedEvaluator.Describe(result)));
            return new(steps, null);
        }
        catch (Exception ex) when (ex is FormatException or InvalidOperationException or ArgumentException or OverflowException)
        { return new([], ex.Message); } // Never expose a partial trace as a complete evaluation.
    }

    public static IReadOnlyList<FormulaAuditTarget> Errors(WorkbookSnapshot snapshot, CancellationToken cancellationToken = default, int limit = 500)
    {
        if (limit is < 1 or > 5000) throw new ArgumentOutOfRangeException(nameof(limit));
        var targets = new List<FormulaAuditTarget>();
        foreach (var sheet in snapshot.Sheets)
        foreach (var (position, cell) in sheet.Cells.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.Value.Kind != CellValueKind.Error) continue;
            if ((cell.Content.Kind == CellContentKind.Formula || cell.SpillOwner is not null) &&
                (cell.Calculation is not { State: CalculationState.Fresh or CalculationState.Error } calculation || calculation.InputRevision != snapshot.Revision)) continue;
            targets.Add(FormulaAudit.Target(snapshot, new(sheet.Id, position.Row, position.Column), new(position.Row, position.Column, position.Row, position.Column)));
            if (targets.Count == limit) return targets;
        }
        return targets.Take(limit).ToArray();
    }
}
