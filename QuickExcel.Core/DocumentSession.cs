using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>
/// Owns a document's immutable state. Short commands and result commits share one serialized write gate;
/// calculation and storage run outside that gate against captured snapshots.
/// </summary>
public sealed class DocumentSession : IDisposable
{
    private readonly object _gate = new();
    private WorkbookSnapshot _snapshot;
    private readonly LinkedList<HistoryEntry> _undo = new();
    private readonly Stack<HistoryEntry> _redo = new();
    private long _historyBytes;
    private long _savedRevision;
    private long _calculationGeneration;
    private long _scriptGeneration;
    private CancellationTokenSource? _activeCalculation;
    private Guid _activeActionId;
    private bool _scriptsEnabled;
    private bool _closed;
    private ImmutableArray<string> _diagnostics = [];
    private readonly Queue<DocumentChange> _notifications = new();
    private bool _publishing;

    private readonly IWorkbookValidation? _validation;
    public DocumentSession(WorkbookSnapshot snapshot, long historyBudgetBytes = 32 * 1024 * 1024, IWorkbookValidation? validation = null)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(historyBudgetBytes);
        _snapshot = snapshot;
        _savedRevision = snapshot.FileState?.RequiresMigration == true ? -1 : snapshot.Revision;
        HistoryBudgetBytes = historyBudgetBytes;
        _validation = validation;
    }

    public Guid SessionId { get; } = Guid.NewGuid();
    public long HistoryBudgetBytes { get; }
    public int DiscardedHistoryCount { get; private set; }
    public event EventHandler<DocumentChange>? Changed;
    public WorkbookSnapshot Snapshot { get { lock (_gate) { ThrowIfClosed(); return _snapshot; } } }
    public long Revision { get { lock (_gate) return _snapshot.Revision; } }
    public long SavedRevision { get { lock (_gate) return _savedRevision; } }
    public bool IsDirty { get { lock (_gate) return _savedRevision != _snapshot.Revision; } }
    public bool CanUndo { get { lock (_gate) return !_closed && _undo.Count != 0; } }
    public bool CanRedo { get { lock (_gate) return !_closed && _redo.Count != 0; } }
    public long HistoryBytes { get { lock (_gate) return _historyBytes; } }
    public bool ScriptsEnabled { get { lock (_gate) return _scriptsEnabled; } }
    public bool IsClosed { get { lock (_gate) return _closed; } }
    public ImmutableArray<string> Diagnostics { get { lock (_gate) return _diagnostics; } }

    public CommandResult Execute(DocumentCommand command, long? expectedRevision = null, bool acceptValidationWarnings = false) =>
        ExecuteCore(command, expectedRevision, acceptValidationWarnings: acceptValidationWarnings);

    /// <summary>Navigation is captured by the next save without input revisions, undo entries or recalculation.
    /// Freeze/zoom/visibility changes must use commands instead.</summary>
    public void UpdateNavigation(SheetId sheetId, CellPosition cell, CellPosition? viewport = null)
    {
        lock (_gate)
        {
            ThrowIfClosed(); WorkbookLimits.ValidatePosition(cell.Row, cell.Column);
            var sheet = _snapshot.GetSheet(sheetId);
            if (sheet.Visibility != SheetVisibility.Visible) throw new InvalidOperationException("不能激活隐藏工作表。");
            var view = sheet.View.ActiveCell == cell ? sheet.View : sheet.View.WithActiveCell(cell);
            if (viewport is { } first)
            {
                WorkbookLimits.ValidatePosition(first.Row, first.Column);
                view = view.HasFreeze ? view with { PaneTopLeftCell = new(Math.Max(first.Row, view.FrozenRows), Math.Max(first.Column, view.FrozenColumns)) }
                    : view with { TopLeftCell = first };
            }
            if (view != sheet.View)
                _snapshot = _snapshot.WithSheets(_snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithView(view) : s));
            if (_snapshot.ActiveSheetId != sheetId) _snapshot = _snapshot.WithActiveSheet(sheetId);
        }
    }

    private CommandResult ExecuteCore(DocumentCommand command, long? expectedRevision, ScriptActionRequest? action = null, CancellationToken cancellationToken = default, ImmutableArray<string> diagnostics = default, bool acceptValidationWarnings = false)
    {
        ArgumentNullException.ThrowIfNull(command);
        DocumentChange? change = null;
        CommandResult result;
        lock (_gate)
        {
            ThrowIfClosed();
            if (action is not null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                action.CancellationToken.ThrowIfCancellationRequested();
                if (!_scriptsEnabled || !Matches(action.Stamp) || _activeActionId != action.OperationId || !_snapshot.ScriptModules.HasActions)
                    throw new InvalidOperationException("动作已过期或未授权，请重新运行。");
            }
            CheckRevision(expectedRevision);
            var before = _snapshot;
            if (before.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
            var draft = command.Prepare(before); // No session state is changed until every check succeeds.
            TableSafety.ValidateEdit(before, draft.Snapshot);
            var inputChanges = draft.Changes.Where(c => c.Before?.Content.WithStyle(0) != c.After?.Content.WithStyle(0)).Select(c => c.Address).ToArray();
            if (inputChanges.Length > 0 && before.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("dataValidation", StringComparison.OrdinalIgnoreCase)) == true)
                throw new NotSupportedException("工作簿含尚未支持的数据验证，不能确认输入限制；本次修改未应用。原规则随原文件保留。");
            var constrained = inputChanges.Where(a => draft.Snapshot.Sheets.FirstOrDefault(s => s.Id == a.SheetId)?.Data.Validations.Any(v => v.ShowError && v.Contains(new(a.Row, a.Column))) == true).ToArray();
            if (constrained.Length > 0)
            {
                if (_validation is null) throw new NotSupportedException("此会话未配置数据验证计算器，修改未应用。");
                var failures = _validation.Check(draft.Snapshot, constrained);
                if (failures.Count > 0 && (!acceptValidationWarnings || failures.Any(f => f.Alert == ValidationAlert.Stop)))
                    throw new DataValidationException(failures);
            }
            cancellationToken.ThrowIfCancellationRequested();
            if (ReferenceEquals(draft.Snapshot, before))
            {
                if (action is not null) { AdvanceCalculationGeneration(); _diagnostics = diagnostics; }
                return new(false, before.Revision);
            }
            var sourceChanged = command is UpdateScriptSourceCommand or UpdateScriptModulesCommand or UpdateScriptSignatureCommand;
            var revised = draft.Snapshot.WithRevision(checked(before.Revision + 1));
            var after = draft.CalculationUnchanged ? CarryCalculationRevision(revised, before.Revision) : Invalidate(revised);
            var entry = sourceChanged ? null : new HistoryEntry(before, after, draft.Changes, draft.StructureChanged, draft.CalculationUnchanged);
            if (entry is not null && entry.Bytes > HistoryBudgetBytes)
                throw new InvalidOperationException("此操作超过撤销历史容量，尚未修改工作簿。请缩小操作范围。");
            cancellationToken.ThrowIfCancellationRequested();
            if (!sourceChanged)
            {
                foreach (var redo in _redo) _historyBytes -= redo.Bytes;
                _redo.Clear();
            }
            if (entry is not null)
            {
                _undo.AddLast(entry);
                _historyBytes += entry.Bytes;
                TrimHistory();
            }
            if (sourceChanged)
            {
                _scriptsEnabled = false;
                _scriptGeneration = checked(_scriptGeneration + 1);
            }
            AdvanceCalculationGeneration();
            _snapshot = after;
            if (!draft.CalculationUnchanged) _diagnostics = diagnostics.IsDefault ? [] : diagnostics;
            change = new(after.Revision, ActualChanges(before, after, draft.Changes), draft.StructureChanged)
            { CalculationUnchanged = draft.CalculationUnchanged };
            QueueNotification(change);
            result = new(true, after.Revision) { CalculationUnchanged = draft.CalculationUnchanged };
        }
        PublishNotifications();
        return result;
    }

    public Task<CommandResult> ExecuteAsync(DocumentCommand command, long? expectedRevision = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(Execute(command, expectedRevision));
    }

    public ScriptActionRequest BeginAction(string? name)
    {
        lock (_gate)
        {
            ThrowIfClosed();
            if (!_scriptsEnabled || !_snapshot.ScriptModules.HasActions) throw new InvalidOperationException("请先启用包含 actions 绑定的脚本。");
            if (name is not null && (string.IsNullOrWhiteSpace(name) || name.Length > 128 || name.Any(char.IsControl))) throw new ArgumentException("动作名称无效。");
            AdvanceCalculationGeneration();
            _activeCalculation = new CancellationTokenSource();
            _activeActionId = Guid.NewGuid();
            return new(new(SessionId, _snapshot.Revision, _calculationGeneration, _scriptGeneration), _activeActionId, name, _snapshot, _activeCalculation.Token);
        }
    }

    public bool CompleteActionDiscovery(ScriptActionRequest request)
    {
        lock (_gate)
        {
            if (!_scriptsEnabled || !Matches(request.Stamp) || _activeActionId != request.OperationId || request.Name is not null) return false;
            AdvanceCalculationGeneration();
            return true;
        }
    }

    public CommandResult CommitAction(ScriptActionRequest request, ScriptActionResult result, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(result);
        if (request.Name is null) throw new InvalidOperationException("动作发现请求不能提交变更。");
        try { return ExecuteCore(new ApplyActionPatchCommand(request, result), request.Stamp.InputRevision, request, cancellationToken, result.Logs); }
        catch (Exception error)
        {
            lock (_gate)
            {
                if (Matches(request.Stamp) && _activeActionId == request.OperationId)
                {
                    AdvanceCalculationGeneration();
                    _diagnostics = [error.Message];
                }
            }
            throw;
        }
    }

    public bool Undo(long? expectedRevision = null) => UndoWithResult(expectedRevision).Changed;
    public bool Redo(long? expectedRevision = null) => RedoWithResult(expectedRevision).Changed;
    public CommandResult UndoWithResult(long? expectedRevision = null) => MoveHistory(redo: false, expectedRevision);
    public CommandResult RedoWithResult(long? expectedRevision = null) => MoveHistory(redo: true, expectedRevision);

    private CommandResult MoveHistory(bool redo, long? expectedRevision)
    {
        DocumentChange change;
        CommandResult result;
        lock (_gate)
        {
            ThrowIfClosed();
            CheckRevision(expectedRevision);
            if (redo ? _redo.Count == 0 : _undo.Count == 0) return new(false, _snapshot.Revision);
            var before = _snapshot;
            var entry = redo ? _redo.Pop() : _undo.Last!.Value;
            if (!redo) _undo.RemoveLast();
            // Script text has its own editor history. Cell undo never reverts source or grants execution.
            var restored = (redo ? entry.After : entry.Before).WithScriptModules(before.ScriptModules)
                .WithRevision(checked(before.Revision + 1));
            // Presentation history must use the current caches, not historical results that
            // may predate a script change, a failed calculation or a newer manual calculation.
            _snapshot = entry.CalculationUnchanged
                ? CarryCalculationRevision(RestoreCurrentCalculations(restored, before), before.Revision)
                : Invalidate(restored);
            AdvanceCalculationGeneration();
            if (redo) _undo.AddLast(entry); else _redo.Push(entry);
            if (!entry.CalculationUnchanged) _diagnostics = [];
            change = new(_snapshot.Revision, ActualChanges(before, _snapshot, entry.Changes), entry.StructureChanged, IsUndoOrRedo: true)
            { CalculationUnchanged = entry.CalculationUnchanged };
            QueueNotification(change);
            result = new(true, _snapshot.Revision) { CalculationUnchanged = entry.CalculationUnchanged };
        }
        PublishNotifications();
        return result;
    }

    public void SetScriptsEnabled(bool enabled)
    {
        DocumentChange change;
        lock (_gate)
        {
            ThrowIfClosed();
            if (enabled && _snapshot.FileState?.FormulaReadOnly == true) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
            if (enabled && !_snapshot.ScriptModules.CanExecute) throw new NotSupportedException("多模块执行需要显式声明脚本 API v2。");
            // Reapply also invalidates an in-flight compile of the same source.
            _scriptsEnabled = enabled;
            _scriptGeneration = checked(_scriptGeneration + 1);
            AdvanceCalculationGeneration();
            _snapshot = Invalidate(_snapshot);
            _diagnostics = [];
            change = new(_snapshot.Revision, [], false, CalculationOnly: true);
            QueueNotification(change);
        }
        PublishNotifications();
    }

    public CalculationRequest BeginCalculation()
    {
        lock (_gate)
        {
            ThrowIfClosed();
            AdvanceCalculationGeneration();
            _activeCalculation = new CancellationTokenSource();
            return new(new(SessionId, _snapshot.Revision, _calculationGeneration, _scriptGeneration), _snapshot, _scriptsEnabled)
            { CancellationToken = _activeCalculation.Token };
        }
    }

    public bool TryCommitCalculation(CalculationResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        DocumentChange change;
        lock (_gate)
        {
            if (!Matches(result.Stamp)) return false;
            var sheets = _snapshot.Sheets.ToArray();
            var builders = new Dictionary<SheetId, ImmutableDictionary<CellPosition, CellSnapshot>.Builder>();
            var changes = ImmutableArray.CreateBuilder<CellChange>();
            foreach (var (address, calculation) in result.Cells)
            {
                WorkbookLimits.ValidatePosition(address.Row, address.Column);
                var sheet = _snapshot.GetSheet(address.SheetId);
                var before = sheet.GetCell(address.Row, address.Column);
                if (before?.Content.Kind != CellContentKind.Formula)
                    throw new InvalidOperationException("计算结果只能提交到本轮快照的公式单元格。");
                if (calculation.InputRevision != result.Stamp.InputRevision || calculation.CalculationGeneration != result.Stamp.CalculationGeneration || calculation.ScriptGeneration != result.Stamp.ScriptGeneration)
                    throw new InvalidOperationException("计算单元格的修订或代次不匹配。");
                if (calculation.Value is { Kind: CellValueKind.Text } text && text.TextValue.Length > WorkbookLimits.MaxCellTextLength)
                    throw new InvalidOperationException("计算结果超过单元格文本容量。");
                if (_snapshot.FileState?.FormulaReadOnly == true && (calculation.State != CalculationState.Blocked || calculation.Value != before.Calculation?.Value))
                    throw new InvalidOperationException("只读公式工作簿只能保留原缓存并标记受阻，不能发布新结果。");
                if (!builders.TryGetValue(sheet.Id, out var builder)) builders[sheet.Id] = builder = sheet.Cells.ToBuilder();
                var after = new CellSnapshot(before.Content, calculation);
                builder[new(address.Row, address.Column)] = after;
                changes.Add(new(address, before, after));
            }
            for (var index = 0; index < sheets.Length; index++)
                if (builders.TryGetValue(sheets[index].Id, out var builder)) sheets[index] = sheets[index].WithCells(builder.ToImmutable());
            _snapshot = _snapshot.WithSheets(sheets);
            _diagnostics = result.Diagnostics;
            AdvanceCalculationGeneration(); // A completed request is single-use.
            // Keep the latest result for redo, without changing any input revision or execution permission.
            if (_undo.Last is { } last && last.Value.After.Revision == _snapshot.Revision)
            {
                _historyBytes -= last.Value.Bytes;
                last.Value = last.Value with { After = _snapshot };
                _historyBytes += last.Value.Bytes;
                TrimHistory();
            }
            change = new(_snapshot.Revision, changes.ToImmutable(), false, CalculationOnly: true);
            QueueNotification(change);
        }
        PublishNotifications();
        return true;
    }

    public bool TryFailCalculation(CalculationStamp stamp, string message, bool disableScripts)
    {
        DocumentChange change;
        lock (_gate)
        {
            if (!Matches(stamp)) return false;
            _diagnostics = [message.Length > 4000 ? message[..4000] : message];
            if (disableScripts)
            {
                _scriptsEnabled = false;
                _scriptGeneration = checked(_scriptGeneration + 1);
            }
            AdvanceCalculationGeneration();
            _snapshot = Invalidate(_snapshot);
            change = new(_snapshot.Revision, [], false, CalculationOnly: true);
            QueueNotification(change);
        }
        PublishNotifications();
        return true;
    }

    public bool MarkSaved(Guid sessionId, long capturedRevision)
    {
        lock (_gate)
        {
            if (_closed || sessionId != SessionId || capturedRevision > _snapshot.Revision || capturedRevision < _savedRevision) return false;
            _savedRevision = capturedRevision;
            QueueNotification(new(_snapshot.Revision, [], false, CalculationOnly: true));
        }
        PublishNotifications();
        return true;
    }

    private bool Matches(CalculationStamp stamp) => !_closed && stamp.SessionId == SessionId && stamp.InputRevision == _snapshot.Revision
        && stamp.CalculationGeneration == _calculationGeneration && stamp.ScriptGeneration == _scriptGeneration;

    private static WorkbookSnapshot Invalidate(WorkbookSnapshot snapshot)
    {
        var sheets = snapshot.Sheets.Select(sheet =>
        {
            var cells = sheet.Cells.ToBuilder();
            var changed = false;
            foreach (var (position, cell) in sheet.Cells)
            {
                if (cell.Content.Kind != CellContentKind.Formula) continue;
                var old = cell.Calculation;
                var state = old?.State == CalculationState.Blocked ? CalculationState.Blocked
                    : old?.Value is null ? CalculationState.Dirty : CalculationState.Stale;
                cells[position] = new(cell.Content, new(old?.Value, state, old?.InputRevision ?? snapshot.Revision,
                    old?.CalculationGeneration ?? 0, old?.ScriptGeneration ?? 0, old?.Diagnostics));
                changed = true;
            }
            return changed ? sheet.WithCells(cells.ToImmutable()) : sheet;
        });
        return snapshot.WithSheets(sheets);
    }

    private static WorkbookSnapshot RestoreCurrentCalculations(WorkbookSnapshot restored, WorkbookSnapshot current) =>
        restored.WithSheets(restored.Sheets.Select(sheet =>
        {
            var currentSheet = current.GetSheet(sheet.Id);
            var cells = sheet.Cells.ToBuilder();
            foreach (var (position, cell) in sheet.Cells)
                if (cell.Content.Kind == CellContentKind.Formula)
                    cells[position] = new(cell.Content, currentSheet.GetCell(position.Row, position.Column)?.Calculation);
            return sheet.WithCells(cells.ToImmutable());
        }));

    // Presentation edits cannot change formula inputs. Carry only results valid for the previous
    // revision; stale/blocked/missing caches stay that way. Generation advancement still rejects
    // late calculations/actions. Content/structure undo still invalidates calculations.
    private static WorkbookSnapshot CarryCalculationRevision(WorkbookSnapshot snapshot, long previousRevision)
    {
        return snapshot.WithSheets(snapshot.Sheets.Select(sheet =>
        {
            ImmutableDictionary<CellPosition, CellSnapshot>.Builder? cells = null;
            foreach (var (position, cell) in sheet.Cells)
                if (cell.Content.Kind == CellContentKind.Formula && cell.Calculation is { } cache &&
                    cache.InputRevision == previousRevision && cache.State is CalculationState.Fresh or CalculationState.Error)
                {
                    cells ??= sheet.Cells.ToBuilder();
                    cells[position] = new(cell.Content, new(cache.Value, cache.State, snapshot.Revision,
                        cache.CalculationGeneration, cache.ScriptGeneration, cache.Diagnostics));
                }
            return cells is null ? sheet : sheet.WithCells(cells.ToImmutable());
        }));
    }

    private void TrimHistory()
    {
        while (_historyBytes > HistoryBudgetBytes && _undo.First is { } first)
        {
            _historyBytes -= first.Value.Bytes;
            _undo.RemoveFirst();
            DiscardedHistoryCount++;
        }
    }

    private static ImmutableArray<CellChange> ActualChanges(WorkbookSnapshot before, WorkbookSnapshot after, ImmutableArray<CellChange> changes) =>
        changes.Select(change => new CellChange(change.Address,
            before.Sheets.FirstOrDefault(s => s.Id == change.Address.SheetId)?.GetCell(change.Address.Row, change.Address.Column),
            after.Sheets.FirstOrDefault(s => s.Id == change.Address.SheetId)?.GetCell(change.Address.Row, change.Address.Column))).ToImmutableArray();

    private void CheckRevision(long? expected)
    {
        if (expected.HasValue && expected != _snapshot.Revision) throw new InvalidOperationException("文档已发生变化，请重新提交操作。");
    }
    private void AdvanceCalculationGeneration()
    {
        _activeActionId = Guid.Empty;
        _calculationGeneration = checked(_calculationGeneration + 1);
        CancelActiveCalculation();
    }

    private void CancelActiveCalculation()
    {
        var active = _activeCalculation;
        _activeCalculation = null;
        if (active is null) return;
        try { active.Cancel(); }
        catch (AggregateException) { System.Diagnostics.Trace.TraceError("Calculation cancellation callback failed."); }
        finally { active.Dispose(); }
    }

    private void ThrowIfClosed() => ObjectDisposedException.ThrowIf(_closed, this);
    private void QueueNotification(DocumentChange change) => _notifications.Enqueue(change with { Snapshot = _snapshot, SessionId = SessionId });

    private void PublishNotifications()
    {
        lock (_gate)
        {
            if (_publishing) return;
            _publishing = true;
        }
        while (true)
        {
            DocumentChange change;
            Delegate[] handlers;
            lock (_gate)
            {
                if (!_notifications.TryDequeue(out change!)) { _publishing = false; return; }
                handlers = Changed?.GetInvocationList() ?? [];
            }
            foreach (EventHandler<DocumentChange> handler in handlers)
            {
                // A faulty observer cannot turn an already committed transaction into a reported failure.
                try { handler(this, change); }
                catch (Exception exception) { System.Diagnostics.Trace.TraceError($"Document observer failed: {exception.GetType().Name}"); }
            }
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_closed) return;
            _closed = true;
            CancelActiveCalculation();
            _scriptsEnabled = false;
            _undo.Clear(); _redo.Clear(); _historyBytes = 0;
            _notifications.Clear();
            Changed = null;
        }
    }

    private sealed record HistoryEntry(WorkbookSnapshot Before, WorkbookSnapshot After, ImmutableArray<CellChange> Changes, bool StructureChanged, bool CalculationUnchanged)
    {
        // Unchanged immutable input and metadata are shared. Count changed payloads and copied formula caches.
        public long Bytes => 1024L + Changes.Sum(change => Estimate(change.Before) + Estimate(change.After) + 256L)
            + FormulaCaches(Before) + FormulaCaches(After)
            + Before.DefinedNames.Sum(n => 128L + 2L * (n.Name.Length + n.Formula.Length))
            + After.DefinedNames.Sum(n => 128L + 2L * (n.Name.Length + n.Formula.Length))
            + RemovedSheetMetadata(Before, After) + RemovedSheetMetadata(After, Before)
            + ChangedDataMetadata(Before, After) + ChangedDataMetadata(After, Before)
            + (StructureChanged ? 512L * Math.Abs(After.Sheets.Length - Before.Sheets.Length) : 0);
        private static long ChangedDataMetadata(WorkbookSnapshot from, WorkbookSnapshot to) => from.Sheets
            .Where(s => !ReferenceEquals(s.Data, to.Sheets.FirstOrDefault(other => other.Id == s.Id)?.Data))
            .Sum(s => 128L + s.Data.Validations.Sum(v => v.EstimatedBytes) + (s.Data.Filter?.Columns.Values.Sum(c => 256L + (c.Values?.Sum(v => 32L + v.Length * 2L) ?? 0) +
                ((c.First?.Operand.Length ?? 0) + (c.Second?.Operand.Length ?? 0)) * 2L) ?? 0) +
                (s.Data.Sort?.Keys.Sum(k => 128L + k.CustomOrder.Sum(v => 32L + v.Length * 2L)) ?? 0) +
                s.Data.Tables.Sum(t => 1024L + t.Style.Elements.Length * 512L + t.Columns.Sum(c => 512L + 2L * (c.Name.Length + (c.CalculatedFormula?.Length ?? 0) + (c.TotalsFormula?.Length ?? 0) + (c.TotalsLabel?.Length ?? 0))) +
                    (t.Filter?.Columns.Values.Sum(c => 256L + (c.Values?.Sum(v => 32L + v.Length * 2L) ?? 0) + ((c.First?.Operand.Length ?? 0) + (c.Second?.Operand.Length ?? 0)) * 2L) ?? 0) +
                    (t.Sort?.Keys.Sum(k => 128L + k.CustomOrder.Sum(v => 32L + v.Length * 2L)) ?? 0)));
        private static long RemovedSheetMetadata(WorkbookSnapshot from, WorkbookSnapshot to) => from.Sheets
            .Where(s => !to.Sheets.Any(other => other.Id == s.Id))
            .Sum(s => 96L * (s.Rows.Count + s.Columns.Count) + s.Images.Sum(image => 128L + image.Data.Length)
                + s.MergedCells.Sum(merge => 32L + merge.Length * 2L));
        private static long FormulaCaches(WorkbookSnapshot snapshot) => snapshot.Sheets.Sum(sheet => sheet.Cells.Values
            .Where(cell => cell.Content.Kind == CellContentKind.Formula).Sum(cell => 256L
                + (cell.Calculation?.Value is { Kind: CellValueKind.Text } value ? value.TextValue.Length * 2L : 0)
                + (cell.Calculation?.Diagnostics.Sum(message => 32L + message.Length * 2L) ?? 0)));
        private static long Estimate(CellSnapshot? cell) => cell is null ? 0 : 256L
            + (cell.Content.Kind == CellContentKind.Formula ? cell.Content.FormulaText.Length * 2L
                : cell.Content.LiteralValue.Kind == CellValueKind.Text ? cell.Content.LiteralValue.TextValue.Length * 2L : 0);
    }
}
