using System.Diagnostics;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

/// <summary>Workbook-scoped typed calculation with conservative script rounds and immutable result reuse.</summary>
public sealed class WorkbookCalculator(IScriptRuntimeFactory runtimeFactory) : IWorkbookCalculator, IWorkbookValidation
{
    public IReadOnlyList<ValidationFailure> Check(WorkbookSnapshot snapshot, IReadOnlyList<CellAddress> addresses) => new WorkbookValidation().Check(snapshot, addresses);
    public IReadOnlyList<CellValue> ListValues(WorkbookSnapshot snapshot, SheetId sheet, CellPosition position, WorksheetValidation rule) =>
        new WorkbookValidation().ListValues(snapshot, sheet, position, rule);
    public static readonly TimeSpan TimeLimit = TimeSpan.FromSeconds(5);
    private readonly IScriptRuntimeFactory _factory = runtimeFactory ?? throw new ArgumentNullException(nameof(runtimeFactory));
    private readonly object _gate = new();
    private Cache? _cache;
    private sealed record Cache(CalculationRequest Request, CalculationResult Result);

    public CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default)
    {
        using var lifetime = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, request.CancellationToken);
        cancellationToken = lifetime.Token;
        var timer = Stopwatch.StartNew();
        void Check()
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (timer.Elapsed > TimeLimit) throw new TimeoutException("重算达到 5 秒限制，本轮结果未提交。");
        }
        Check();
        if (request.Snapshot.FileState?.FormulaReadOnly == true)
        {
            var stamp = request.Stamp;
            var retained = new Dictionary<CellAddress, CellCalculation>();
            foreach (var sheet in request.Snapshot.Sheets)
            foreach (var (position, cell) in sheet.Cells)
            {
                Check();
                if (cell.Content.Kind != CellContentKind.Formula) continue;
                retained.Add(new(sheet.Id, position.Row, position.Column), new(cell.Calculation?.Value, CalculationState.Blocked,
                    stamp.InputRevision, stamp.CalculationGeneration, stamp.ScriptGeneration, ["只读公式及其受影响依赖范围未知；保留来源缓存，未重算。"]));
            }
            return new(stamp, retained, [WorkbookFileState.FormulaReadOnlyMessage]);
        }
        if (request.ScriptsEnabled && !request.Snapshot.ScriptModules.CanExecute)
            throw new NotSupportedException("多模块执行需要显式声明脚本 API v2。");
        var index = new WorkbookDependencyIndex(request.Snapshot, Check);
        Cache? prior; lock (_gate) prior = _cache;
        var reusable = !request.ForceFull && !request.ScriptsEnabled && !index.Nodes.Values.Any(n => n.Script)
            && prior is not null && !prior.Request.ScriptsEnabled
            && prior.Request.Stamp.SessionId == request.Stamp.SessionId
            && prior.Request.Stamp.ScriptGeneration == request.Stamp.ScriptGeneration
            && prior.Request.Snapshot.DocumentId == request.Snapshot.DocumentId
            && prior.Request.Snapshot.DateSystem == request.Snapshot.DateSystem
            && prior.Request.Snapshot.TextCompatibilityVersion == request.Snapshot.TextCompatibilityVersion
            && prior.Request.Snapshot.ScriptSource == request.Snapshot.ScriptSource
            && prior.Request.Snapshot.DefinedNames.SequenceEqual(request.Snapshot.DefinedNames)
            && prior.Request.Snapshot.Sheets.Select(s => (s.Id, s.Name)).SequenceEqual(request.Snapshot.Sheets.Select(s => (s.Id, s.Name)));
        var affected = reusable ? index.AffectedBy(Changed(prior!.Request.Snapshot, request.Snapshot, Check), Check) : index.Nodes.Keys.ToHashSet();
        affected.UnionWith(index.Closure(index.Nodes.Where(p => p.Value.Volatile).Select(p => p.Key), Check));
        var cells = new Dictionary<CellAddress, CellCalculation>();
        var diagnostics = new List<string>();
        var evaluated = 0; var reused = 0;
        void Store(CellAddress address, CellValue? value, CalculationState state, IEnumerable<string>? messages = null)
        {
            var stamp = request.Stamp;
            var notes = (messages ?? []).Take(20).Select(m => m.Length > 2000 ? m[..2000] : m).ToArray();
            cells[address] = new(value, state, stamp.InputRevision, stamp.CalculationGeneration, stamp.ScriptGeneration, notes);
            if (diagnostics.Count >= 100) return;
            var sheet = request.Snapshot.GetSheet(address.SheetId);
            foreach (var note in notes)
            {
                if (diagnostics.Count == 100) break;
                diagnostics.Add($"{sheet.Name}!{FormulaSyntax.ColumnName(address.Column)}{address.Row + 1}：{note}");
            }
        }
        using var runtime = request.ScriptsEnabled && !request.Snapshot.ScriptModules.EntriesFor("formula-functions").IsEmpty ? _factory.Create(new ScriptRuntimeRequest(request.Snapshot.ScriptModules,
            request.Stamp, TimeLimit - timer.Elapsed, cancellationToken)) : null;
        Check();
        var functions = new Dictionary<string, ScriptFunctionDescriptor>(StringComparer.OrdinalIgnoreCase);
        if (runtime is not null)
        {
            foreach (var descriptor in runtime.FunctionDescriptors)
            {
                if (StandardFunctions.Signatures.ContainsKey(ExcelFormulaNames.CanonicalName(descriptor.Name)) || !functions.TryAdd(descriptor.Name, descriptor))
                    throw new InvalidOperationException($"脚本函数 {descriptor.Name} 与内置函数或其他注册冲突。");
                descriptor.Validate();
            }
        }
        var evaluator = new TypedEvaluator(request.Snapshot, cells, runtime, functions, Check);
        var unavailable = index.Nodes.Where(p => p.Value.VisibilityBlocked || FormulaSyntax.Walk(p.Value.Expression).OfType<FunctionNode>().Any(f =>
            !StandardFunctions.Signatures.ContainsKey(f.Name) && (!functions.ContainsKey(f.Name) || f.Name.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase) || f.Name.StartsWith("_xlws.", StringComparison.OrdinalIgnoreCase))))
            .Select(p => p.Key);
        var blockedByUnavailable = index.Closure(unavailable, Check);
        var pending = index.Nodes.ToDictionary(p => p.Key, p => p.Value.Dependencies.Count(d => !index.CycleMembers.Contains(d)));
        var ready = new PriorityQueue<CellAddress, int>();
        foreach (var address in index.CycleMembers)
            if (blockedByUnavailable.Contains(address)) Store(address, index.Nodes[address].Cell.Calculation?.Value, CalculationState.Blocked, ["包含不可用函数，不能据部分依赖图确认计算语义；保留旧缓存。"]);
            else Store(address, CellValue.FromError(CellErrorCode.Cycle), CalculationState.Error, ["循环引用：此单元格属于循环强连通分量。"]);
        foreach (var (address, node) in index.Nodes)
            if (pending[address] == 0 && !index.CycleMembers.Contains(address)) ready.Enqueue(address, node.Order);
        IScriptInvocationBatch? batch = null;
        var queued = new List<(CellAddress Address, TypedEvaluator.PreparedScriptCall Call)>();
        var queuedAddresses = new HashSet<CellAddress>();
        void StoreEvaluated(CellAddress address, Func<CellValue> evaluate)
        {
            var notes = new List<string>();
            CellValue value;
            try
            {
                value = evaluate();
                if (value is { Kind: CellValueKind.Text } text && text.TextValue.Length > WorkbookLimits.MaxCellTextLength)
                    throw new FormatException("计算结果超过单元格文本长度限制。");
            }
            catch (Exception e) when (e is FormulaValueException or FormatException or OverflowException or ArgumentException)
            {
                value = CellValue.FromError(e is FormulaValueException formula ? formula.Code : CellErrorCode.Value);
                notes.Add(e.Message);
            }
            if (index.Nodes[address].ParseError is { } parseError) notes.Add(parseError);
            if (value.Kind == CellValueKind.Error && notes.Count == 0) notes.Add($"计算结果：{CellInput.DisplayText(value)}");
            Store(address, value, value.Kind == CellValueKind.Error ? CalculationState.Error : CalculationState.Fresh, notes);
        }
        void Flush()
        {
            if (batch is null || batch.Count == 0) return;
            Check();
            var replies = batch.Invoke();
            Check();
            if (replies.Count != queued.Count) throw new ScriptRuntimeFailureException("Protocol", "脚本批次结果数量不匹配，本轮未提交。");
            for (var i = 0; i < queued.Count; i++)
            {
                Check();
                var (address, call) = queued[i];
                StoreEvaluated(address, () => evaluator.CompleteScript(call, replies[i]));
            }
            batch = null; queued.Clear(); queuedAddresses.Clear();
        }
        while (ready.TryDequeue(out var address, out _))
        {
            Check();
            var node = index.Nodes[address];
            var canBatch = runtime is IScriptBatchRuntime && evaluator.CanBatch(node.Expression);
            // Release graph edges in exactly the original priority order, but flush before a
            // consumer reads a pending result or any non-batchable expression executes.
            if (!canBatch || node.Dependencies.Overlaps(queuedAddresses)) Flush();
            if (reusable && !affected.Contains(address) && prior!.Result.Cells.TryGetValue(address, out var cached)
                && cached.State is CalculationState.Fresh or CalculationState.Error)
            {
                Store(address, cached.Value, cached.State, cached.Diagnostics); reused++;
            }
            else
            {
                evaluated++;
                var notes = new List<string>();
                var dependencies = node.Dependencies.Select(d => cells[d]).ToArray();
                var blocked = blockedByUnavailable.Contains(address) || !request.ScriptsEnabled && node.Script || dependencies.Any(d => d.State == CalculationState.Blocked);
                CellValue? value;
                if (!blocked && dependencies.Any(d => d.State == CalculationState.Error && d.Value is { Kind: CellValueKind.Error, ErrorCode: CellErrorCode.Cycle }))
                {
                    value = CellValue.FromError(CellErrorCode.Cycle);
                    notes.Add("依赖循环引用：此单元格不属于循环分量，其上游包含循环。");
                    blocked = false;
                }
                else if (blocked)
                {
                    value = node.Cell.Calculation?.Value;
                    notes.Add("函数不可用、脚本未启用或上游结果受阻；保留旧缓存，不能视为本轮新结果。");
                }
                else
                {
                    evaluator.CurrentSheet = address.SheetId;
                    object? prepared = null;
                    try
                    {
                        if (canBatch) prepared = evaluator.PrepareScript(node.Expression);
                    }
                    catch (Exception e) when (e is FormulaValueException or FormatException or OverflowException or ArgumentException)
                    {
                        Flush();
                        StoreEvaluated(address, () => throw e);
                        goto ReleaseDependents;
                    }
                    if (prepared is TypedEvaluator.PreparedScriptCall call)
                    {
                        batch ??= ((IScriptBatchRuntime)runtime!).CreateBatch();
                        if (!batch.TryAdd(call.Descriptor.Name, call.Arguments))
                        {
                            Flush();
                            batch = ((IScriptBatchRuntime)runtime!).CreateBatch();
                            if (!batch.TryAdd(call.Descriptor.Name, call.Arguments))
                                throw new ScriptRuntimeFailureException("InputLimit", "单次脚本调用超过批次限额。");
                        }
                        queued.Add((address, call)); queuedAddresses.Add(address);
                    }
                    else
                    {
                        Flush();
                        StoreEvaluated(address, () => prepared is CellValue error ? error : evaluator.CellResult(node.Expression));
                    }
                    goto ReleaseDependents;
                }
                Flush();
                Store(address, value, blocked ? CalculationState.Blocked : value?.Kind == CellValueKind.Error ? CalculationState.Error : CalculationState.Fresh, notes);
            }
        ReleaseDependents:
            foreach (var dependent in node.Dependents)
                if (--pending[dependent] == 0 && !index.CycleMembers.Contains(dependent)) ready.Enqueue(dependent, index.Nodes[dependent].Order);
        }
        Flush();
        Check();
        if (cells.Count != index.Nodes.Count) throw new InvalidOperationException("计算计划未完成，拒绝发布部分结果。");
        if (runtime is not null) diagnostics.AddRange(runtime.GetLogs().Take(Math.Max(0, 100 - diagnostics.Count)).Select(l => "JavaScript：" + l));
        Check();
        var result = new CalculationResult(request.Stamp, cells, diagnostics)
        { Statistics = new(index.Nodes.Count, evaluated + index.CycleMembers.Count, reused, !reusable, timer.Elapsed) };
        // A cancelled/failed round never becomes a reuse source. An older concurrent round cannot replace a newer cache.
        lock (_gate)
            if (_cache is null || _cache.Request.Stamp.SessionId != request.Stamp.SessionId
                || _cache.Request.Stamp.CalculationGeneration <= request.Stamp.CalculationGeneration)
                _cache = new(request, result);
        return result;
    }

    private static IEnumerable<CellAddress> Changed(WorkbookSnapshot before, WorkbookSnapshot after, Action check)
    {
        foreach (var sheet in after.Sheets)
        {
            var old = before.GetSheet(sheet.Id);
            foreach (var position in sheet.Cells.Keys.Union(old.Cells.Keys))
            {
                check();
                if (sheet.Cells.GetValueOrDefault(position)?.Content != old.Cells.GetValueOrDefault(position)?.Content)
                    yield return new(sheet.Id, position.Row, position.Column);
            }
        }
    }
}
