using System.Diagnostics;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

/// <summary>Discover spill footprints, then replan consumers against their producing anchors.
/// No intermediate pass is published. Scalar-only workbooks retain the incremental/batched path.</summary>
internal static class ArrayCalculationRound
{
    internal static CalculationResult Calculate(CalculationRequest request, IScriptRuntimeFactory factory, Action check, Stopwatch timer)
    {
        var snapshot = request.Snapshot;
        using var runtime = request.ScriptsEnabled && !snapshot.ScriptModules.EntriesFor("formula-functions").IsEmpty
            ? factory.Create(new ScriptRuntimeRequest(snapshot.ScriptModules, request.Stamp, WorkbookCalculator.TimeLimit - timer.Elapsed, request.CancellationToken) { RuntimeLifetime = request.RuntimeLifetime }) : null;
        var functions = new Dictionary<string, ScriptFunctionDescriptor>(StringComparer.OrdinalIgnoreCase);
        if (runtime is not null)
        foreach (var descriptor in runtime.FunctionDescriptors)
        {
            if (StandardFunctions.Signatures.ContainsKey(ExcelFormulaNames.CanonicalName(descriptor.Name)) || !functions.TryAdd(descriptor.Name, descriptor))
                throw new InvalidOperationException($"脚本函数 {descriptor.Name} 与内置函数或其他注册冲突。");
            descriptor.Validate();
        }
        IReadOnlyDictionary<CellAddress, FormulaSpill> planningSpills = new Dictionary<CellAddress, FormulaSpill>();
        var executeScripts = false;
        var evaluated = 0;
        for (var pass = 0; pass < 32; pass++)
        {
            check();
            var index = new WorkbookDependencyIndex(snapshot, check, planningSpills);
            var allocation = new FormulaSpillAllocator(snapshot);
            var cells = new Dictionary<CellAddress, CellCalculation>();
            var diagnostics = new List<string>();
            CellValue? Derived(CellAddress address) => allocation.Owners.TryGetValue(address, out var owner)
                ? allocation.Spills[owner].Array[address.Row - owner.Row, address.Column - owner.Column] : null;
            var evaluator = new TypedEvaluator(snapshot, cells, runtime, functions, check)
            { DerivedValue = Derived, DerivedPositions = () => allocation.Owners.Keys };
            void Store(CellAddress address, CellValue? value, CalculationState state, FormulaArray? array = null, string? note = null, FormulaArrayDimensions? blockedSpill = null)
            {
                cells[address] = new(value, state, request.Stamp.InputRevision, request.Stamp.CalculationGeneration,
                    request.Stamp.ScriptGeneration, note is null ? null : [note], array, blockedSpill);
                if (note is not null && diagnostics.Count < 100)
                    diagnostics.Add($"{snapshot.GetSheet(address.SheetId).Name}!{FormulaSyntax.ColumnName(address.Column)}{address.Row + 1}：{note}");
            }
            var unavailable = index.Nodes.Where(p => p.Value.VisibilityBlocked || FormulaSyntax.Walk(p.Value.Expression).OfType<FunctionNode>()
                .Any(f => !StandardFunctions.Signatures.ContainsKey(f.Name) && !functions.ContainsKey(f.Name))).Select(p => p.Key).ToArray();
            var scriptDependents = index.Closure(index.Nodes.Where(p => p.Value.Script).Select(p => p.Key), check);
            // A missing or stateful producer has no trustworthy footprint. Do not infer that its
            // former children are empty, or speculatively invoke a stateful UDF during replanning.
            if (index.Closure(unavailable, check).Concat(scriptDependents).Any(a =>
                index.Nodes[a].Cell.Content.DynamicArray || DynamicArrayFunctions.UsesArrays(index.Nodes[a].Expression)))
            {
                const string reason = "数组尺寸依赖不可用函数或脚本调用，不能安全规划溢出区域；保留旧缓存，整轮未重算。";
                return new(request.Stamp, index.Nodes.ToDictionary(p => p.Key, p => new CellCalculation(p.Value.Cell.Calculation?.Value,
                    CalculationState.Blocked, request.Stamp.InputRevision, request.Stamp.CalculationGeneration, request.Stamp.ScriptGeneration, [reason])), [reason]);
            }
            var blocked = index.Closure(unavailable, check);
            if (!executeScripts) blocked.UnionWith(scriptDependents);
            foreach (var address in index.CycleMembers)
                if (blocked.Contains(address)) Store(address, index.Nodes[address].Cell.Calculation?.Value, CalculationState.Blocked, note: "函数不可用，保留来源缓存。");
                else Store(address, CellValue.FromError(CellErrorCode.Cycle), CalculationState.Error, note: "公式或溢出结果存在循环依赖。");
            var pending = index.Nodes.ToDictionary(p => p.Key, p => p.Value.Dependencies.Count(d => !index.CycleMembers.Contains(d)));
            var ready = new PriorityQueue<CellAddress, int>();
            foreach (var (address, node) in index.Nodes)
                if (pending[address] == 0 && !index.CycleMembers.Contains(address)) ready.Enqueue(address, node.Order);
            while (ready.TryDequeue(out var address, out _))
            {
                check(); var node = index.Nodes[address];
                var dependencies = node.Dependencies.ToDictionary(d => d, d => cells[d]);
                if (blocked.Contains(address) || dependencies.Values.Any(d => d.State == CalculationState.Blocked))
                    Store(address, node.Cell.Calculation?.Value, CalculationState.Blocked, note: "函数不可用或上游受阻；旧缓存不是本轮新结果。");
                else if (dependencies.Values.Any(d => d.Value is { Kind: CellValueKind.Error, ErrorCode: CellErrorCode.Cycle }))
                    Store(address, CellValue.FromError(CellErrorCode.Cycle), CalculationState.Error, note: "上游存在循环依赖。");
                else
                {
                    object value;
                    evaluator.CurrentSheet = address.SheetId;
                    evaluator.ArrayMode = node.Cell.Content.DynamicArray || DynamicArrayFunctions.UsesArrays(node.Expression);
                    try
                    {
                        evaluated++;
                        value = evaluator.ArrayMode ? evaluator.FormulaResult(node.Expression) : evaluator.CellResult(node.Expression);
                    }
                    catch (Exception error) when (error is FormulaValueException or FormatException or OverflowException or ArgumentException)
                    { value = CellValue.FromError(error is FormulaValueException formulaError ? formulaError.Code : CellErrorCode.Value); }
                    if (value is FormulaArray array)
                    {
                        if (allocation.TryReserve(new(address, array), out var reason, check))
                            Store(address, array[0,0], array[0,0].Kind == CellValueKind.Error ? CalculationState.Error : CalculationState.Fresh, array);
                        else Store(address, CellValue.FromError(CellErrorCode.Spill), CalculationState.Error, note: reason, blockedSpill: new(array.Rows,array.Columns));
                    }
                    else
                    {
                        var scalar = (CellValue)value;
                        Store(address, scalar, scalar.Kind == CellValueKind.Error ? CalculationState.Error : CalculationState.Fresh);
                    }
                }
                foreach (var dependent in node.Dependents)
                    if (--pending[dependent] == 0 && !index.CycleMembers.Contains(dependent)) ready.Enqueue(dependent, index.Nodes[dependent].Order);
            }
            if (cells.Count != index.Nodes.Count) throw new InvalidOperationException("数组依赖计划未完成，整轮结果未提交。");
            // Retain the footprint which established an actual self/dependency cycle so removing
            // its error result cannot make the next pass forget the cycle and oscillate.
            var nextPlanning = allocation.Spills.ToDictionary();
            foreach (var anchor in index.CycleMembers)
                if (planningSpills.TryGetValue(anchor, out var footprint)) nextPlanning.TryAdd(anchor, footprint);
            if (SameFootprints(planningSpills, nextPlanning))
            {
                check();
                if (!executeScripts && scriptDependents.Count > 0 && runtime is not null)
                { executeScripts = true; planningSpills = nextPlanning; continue; }
                if (runtime is not null) diagnostics.AddRange(runtime.GetLogs().Take(Math.Max(0, 100 - diagnostics.Count)).Select(l => "JavaScript：" + l));
                return new(request.Stamp, cells, diagnostics)
                { Statistics = new(index.Nodes.Count, evaluated, 0, true, timer.Elapsed) };
            }
            if (executeScripts) throw new InvalidOperationException("脚本执行后数组区域发生变化，整轮结果未提交。");
            planningSpills = nextPlanning;
        }
        throw new InvalidOperationException("动态数组区域在 32 次依赖规划后仍未稳定，整轮结果未提交。");
    }

    private static bool SameFootprints(IReadOnlyDictionary<CellAddress, FormulaSpill> before, IReadOnlyDictionary<CellAddress, FormulaSpill> after) =>
        before.Count == after.Count && before.All(p => after.TryGetValue(p.Key, out var next) && p.Value.Range == next.Range);
}
