using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>
/// Owns a document's immutable state. Short commands and result commits share one serialized write gate;
/// calculation and storage run outside that gate against captured snapshots.
/// </summary>
public sealed partial class DocumentSession : IDisposable
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
    private CancellationTokenSource _runtimeLifetime = new();
    private Guid _activeActionId;
    private bool _scriptsEnabled;
    private bool _closed;
    private ImmutableArray<string> _diagnostics = [];
    private readonly Queue<DocumentChange> _notifications = new();
    private bool _publishing;

    private readonly IWorkbookValidation? _validation;
    private readonly IWorkbookCalculator? _calculator;
    private readonly IWorkbookVirtualFiles? _virtualFiles;
    public DocumentSession(WorkbookSnapshot snapshot, long historyBudgetBytes = 32 * 1024 * 1024, IWorkbookValidation? validation = null, IWorkbookVirtualFiles? virtualFiles = null, IWorkbookCalculator? calculator = null)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(historyBudgetBytes);
        _snapshot = snapshot;
        _savedRevision = snapshot.FileState?.RequiresMigration == true ? -1 : snapshot.Revision;
        HistoryBudgetBytes = historyBudgetBytes;
        _validation = validation;
        _calculator = calculator;
        _virtualFiles = virtualFiles;
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

    // Calculated trial values are literal cells in the saved workbook. Validate them
    // through the ordinary edit path, but let one undo restore the initiating edit.
    public CommandResult ApplyAutomaticDataTableResults(SetCellsCommand results, long expectedRevision) =>
        ExecuteCore(results, expectedRevision, includeWithPreviousEdit: true);

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
            if (_snapshot.ActiveSheetId != sheetId)
            {
                EnqueueScriptEvent("sheet.deactivate", _snapshot.GetSheet(_snapshot.ActiveSheetId));
                _snapshot = _snapshot.WithActiveSheet(sheetId);
                EnqueueScriptEvent("sheet.activate", sheet);
            }
        }
    }

    private CommandResult ExecuteCore(DocumentCommand command, long? expectedRevision, ScriptActionRequest? action = null, CancellationToken cancellationToken = default, ImmutableArray<string> diagnostics = default, bool acceptValidationWarnings = false, VirtualFileSet? transientFiles = null, bool includeWithPreviousEdit = false)
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
                if (!Matches(action.Stamp) || _activeActionId != action.OperationId ||
                    (action.ConsoleSource is null ? !_scriptsEnabled || !_snapshot.ScriptModules.HasActions || _activeConsoleSource is not null
                        : action.ConsoleSource != _activeConsoleSource))
                    throw new InvalidOperationException("动作已过期或未授权，请重新运行。");
            }
            CheckRevision(expectedRevision);
            var before = _snapshot;
            if (before.FileState?.FormulaReadOnly == true && action?.ConsoleSource is null) throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
            var draft = command.Prepare(before); // No session state is changed until every check succeeds.
            ExtendedObjectSafety.ValidateEdit(before, draft.Snapshot, command, draft.StructureChanged);
            PivotSafety.ValidateEdit(before,draft.Snapshot,command,draft.StructureChanged);
            draft = draft with { Snapshot = PivotSafety.MarkSourceChanges(before,draft.Snapshot) };
            FormulaSpillSafety.ValidateEdit(before, draft.Snapshot, draft.StructureChanged);
            if (draft.StructureChanged && command is not (CreatePivotCommand or RefreshPivotCommand or RemovePivotCommand) && !draft.CalculationUnchanged && !ReferenceEquals(before,draft.Snapshot))
                draft = draft with { Snapshot = FormulaSpillSafety.DropDerivedAfterStructure(draft.Snapshot) };
            if (command is not ApplyActionPatchCommand { ProtectionValidatedByModel: true })
                ProtectionSafety.Validate(before, draft.Snapshot, command);
            if (draft.Snapshot.Sheets.Any(s => WorksheetOutline.HasGroups(s) && s.Data.AllFilters.Any(f => f.Columns.Count > 0)))
                throw new NotSupportedException("分组与活动筛选的混合编辑暂未对齐 Excel；请先展开并取消分组，或清除筛选条件。");
            // Visibility aggregates also depend on filter metadata and displayed filter labels.
            // Conservatively recalculate these workbooks for presentation-only commands.
            if (draft.CalculationUnchanged && !ReferenceEquals(draft.Snapshot, before) && Formulas.SubtotalFunctions.Present(draft.Snapshot))
                draft = draft with { CalculationUnchanged = false };
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
                if (action is not null) { AdvanceCalculationGeneration(); _diagnostics = diagnostics; AppendConsole(diagnostics); if (transientFiles is not null) _transientFiles = transientFiles; }
                return new(false, before.Revision);
            }
            else
            {
                var sourceChanged = command is UpdateScriptSourceCommand or UpdateScriptModulesCommand or UpdateScriptSignatureCommand || !ReferenceEquals(before.ScriptModules, draft.Snapshot.ScriptModules);
                var revised = draft.Snapshot.WithRevision(checked(before.Revision + 1));
                var after = draft.CalculationUnchanged ? CarryCalculationRevision(revised, before.Revision) : Invalidate(revised);
                var entry = sourceChanged ? null : new HistoryEntry(before, after, draft.Changes, draft.StructureChanged, draft.CalculationUnchanged);
                var previous = includeWithPreviousEdit && _redo.Count == 0 && before.Revision > _savedRevision &&
                    _undo.Last?.Value.After.Revision == before.Revision ? _undo.Last : null;
                if (previous is not null && entry is not null)
                {
                    var first = previous.Value;
                    var addresses = first.Changes.Concat(entry.Changes).DistinctBy(c => c.Address).ToImmutableArray();
                    entry = new(first.Before, after, ActualChanges(first.Before, after, addresses),
                        first.StructureChanged || entry.StructureChanged, false);
                }
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
                    if (previous is not null) { _historyBytes -= previous.Value.Bytes; _undo.Remove(previous); }
                    _undo.AddLast(entry);
                    _historyBytes += entry.Bytes;
                    TrimHistory();
                }
                if (sourceChanged)
                {
                    ResetScriptRuntime();
                    _scriptsEnabled = false;
                    _scriptEvents.Clear();
                    _scriptGeneration = checked(_scriptGeneration + 1);
                }
                AdvanceCalculationGeneration();
                _snapshot = after;
                if (transientFiles is not null) _transientFiles = transientFiles;
                if (action is null) QueueChangeEvents(before, after, draft.Changes);
                if (action is not null || !draft.CalculationUnchanged) { _diagnostics = diagnostics.IsDefault ? [] : diagnostics; AppendConsole(_diagnostics); }
                change = new(after.Revision, ActualChanges(before, after, draft.Changes), draft.StructureChanged)
                { CalculationUnchanged = draft.CalculationUnchanged };
                QueueNotification(change);
                result = new(true, after.Revision) { CalculationUnchanged = draft.CalculationUnchanged };
            }
        }
        PublishNotifications();
        return result;
    }

    public Task<CommandResult> ExecuteAsync(DocumentCommand command, long? expectedRevision = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(Execute(command, expectedRevision));
    }

    public ScriptActionRequest BeginAction(string? name, ScriptEvent? scriptEvent = null) => BeginActionCore(name, scriptEvent, null);
    private ScriptActionRequest BeginActionCore(string? name, ScriptEvent? scriptEvent, string? consoleSource)
    {
        ScriptActionRequest request;
        lock (_gate)
        {
            ThrowIfClosed();
            if (consoleSource is null && (!_scriptsEnabled || !_snapshot.ScriptModules.HasActions)) throw new InvalidOperationException("请先启用包含 actions 绑定的脚本。");
            if (name is not null && (string.IsNullOrWhiteSpace(name) || name.Length > 128 || name.Any(char.IsControl))) throw new ArgumentException("动作名称无效。");
            if (scriptEvent is not null && (scriptEvent.SessionId != SessionId || !ScriptEventNames.All.Contains(scriptEvent.Name) || name != ScriptEventNames.Prefix + scriptEvent.Name) ||
                scriptEvent is null && name?.StartsWith(ScriptEventNames.Prefix, StringComparison.Ordinal) == true) throw new ArgumentException("事件调用身份无效。");
            AdvanceCalculationGeneration();
            _activeCalculation = new CancellationTokenSource();
            _activeActionId = Guid.NewGuid();
            _activeConsoleSource = consoleSource;
            if (consoleSource is not null) AppendConsole([consoleSource], "input");
            request = new(new(SessionId, _snapshot.Revision, _calculationGeneration, _scriptGeneration), _activeActionId, name, _snapshot, _activeCalculation.Token)
            { Files = _transientFiles, Event = scriptEvent, ConsoleSource = consoleSource,
                RuntimeLifetime = consoleSource is null ? _runtimeLifetime.Token : default };
            if (name is not null) request = request with { Model = CreateScriptModel(request) };
        }
        try
        {
            if (_virtualFiles is not null && name is not null)
            {
                request.CancellationToken.ThrowIfCancellationRequested();
                var captured = request;
                request = request with { DeferredFiles = new(() => new VirtualFileSet(_virtualFiles.Capture(captured.Snapshot).Entries.Values.Concat(captured.Files.Entries.Values))) };
                request.CancellationToken.ThrowIfCancellationRequested();
            }
            return request;
        }
        catch (Exception error)
        {
            TryFailCalculation(request.Stamp, "文件快照失败：" + error.Message, disableScripts: false);
            if (error is NotSupportedException) throw new InvalidOperationException("当前工作簿不能构建安全的脚本文件快照：" + error.Message, error);
            throw;
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
        try
        {
            if (result.FileEdits.IsDefault) throw new InvalidOperationException("缺少文件变更。");
            if (request.ConsoleSource is not null && request.Snapshot.FileState?.FormulaReadOnly == true && (!result.Edits.IsEmpty || !result.FileEdits.IsEmpty))
                throw new InvalidOperationException(WorkbookFileState.FormulaReadOnlyMessage);
            var transient = result.FileEdits.IsEmpty ? null : new VirtualFileSet(request.Files.Apply(result.FileEdits).Entries.Values.Where(e => !e.Path.StartsWith("//xlsq/", StringComparison.Ordinal)));
            return ExecuteCore(new ApplyActionPatchCommand(request, result, _virtualFiles), request.Stamp.InputRevision, request, cancellationToken, result.Logs, transientFiles: transient);
        }
        catch (Exception error)
        {
            lock (_gate)
            {
                if (request.RuntimeLifetime.CanBeCanceled && request.RuntimeLifetime == _runtimeLifetime.Token)
                {
                    ResetScriptRuntime();
                    // A rejected transaction discards its execution contexts, including
                    // any UDF work performed by provisional formula reads.
                    _snapshot = Invalidate(_snapshot);
                    QueueNotification(new(_snapshot.Revision, [], false, CalculationOnly: true));
                }
                if (Matches(request.Stamp) && _activeActionId == request.OperationId)
                {
                    AdvanceCalculationGeneration();
                    _diagnostics = [error.Message];
                    if (!result.Logs.IsDefault) AppendConsole(result.Logs);
                    AppendConsole([error.Message], level: ScriptConsoleLevel.Error);
                }
            }
            PublishNotifications();
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
            QueueChangeEvents(before, _snapshot, change.Cells, history: true);
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
            if (enabled && !_snapshot.ScriptModules.CanExecute) throw new NotSupportedException(ScriptModuleSet.MigrationMessage);
            // Reapply also invalidates an in-flight compile of the same source.
            ResetScriptRuntime();
            _scriptsEnabled = enabled;
            _registeredEvents = null;
            _scriptEvents.Clear();
            if (enabled && _snapshot.ScriptModules.HasActions && !_openedEvent) { _openedEvent = true; EnqueueScriptEvent("workbook.open"); EnqueueScriptEvent("sheet.activate", _snapshot.GetSheet(_snapshot.ActiveSheetId)); }
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
            { CancellationToken = _activeCalculation.Token, RuntimeLifetime = _runtimeLifetime.Token };
        }
    }

    public bool TryCommitCalculation(CalculationResult result)
    {
        ArgumentNullException.ThrowIfNull(result);
        DocumentChange change;
        lock (_gate)
        {
            if (!Matches(result.Stamp)) return false;
            _snapshot = ApplyCalculation(_snapshot, result, out var changes);
            _diagnostics = result.Diagnostics;
            AppendConsole(_diagnostics);
            AdvanceCalculationGeneration(); // A completed request is single-use.
            // Keep the latest result for redo, without changing any input revision or execution permission.
            if (_undo.Last is { } last && last.Value.After.Revision == _snapshot.Revision)
            {
                _historyBytes -= last.Value.Bytes;
                last.Value = last.Value with { After = _snapshot };
                _historyBytes += last.Value.Bytes;
                TrimHistory();
            }
            EnqueueScriptEvent("workbook.calculated");
            change = new(_snapshot.Revision, changes, false, CalculationOnly: true);
            QueueNotification(change);
        }
        PublishNotifications();
        return true;
    }

    public bool TryFailCalculation(CalculationStamp stamp, string message, bool disableScripts, IEnumerable<string>? logs = null)
    {
        DocumentChange change;
        lock (_gate)
        {
            if (!Matches(stamp)) return false;
            ResetScriptRuntime();
            _diagnostics = [message.Length > 4000 ? message[..4000] : message];
            if (logs is not null) AppendConsole(logs.Take(100));
            AppendConsole(_diagnostics, level: ScriptConsoleLevel.Error);
            if (disableScripts)
            {
                _scriptsEnabled = false;
                _scriptEvents.Clear();
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
            EnqueueScriptEvent("workbook.saved");
            QueueNotification(new(_snapshot.Revision, [], false, CalculationOnly: true));
        }
        PublishNotifications();
        return true;
    }

    private bool Matches(CalculationStamp stamp) => !_closed && stamp.SessionId == SessionId && stamp.InputRevision == _snapshot.Revision
        && stamp.CalculationGeneration == _calculationGeneration && stamp.ScriptGeneration == _scriptGeneration;

    private static WorkbookSnapshot Invalidate(WorkbookSnapshot snapshot)
    {
        var anyChanged = false;
        var sheets = snapshot.Sheets.Select(sheet =>
        {
            var cells = sheet.Cells.ToBuilder();
            var changed = false;
            foreach (var (position, cell) in sheet.Cells)
            {
                if (cell.Content.Kind != CellContentKind.Formula && cell.SpillOwner is null) continue;
                if (cell.SpillOwner is { } owner && sheet.GetCell(owner.Row, owner.Column)?.Content.Kind != CellContentKind.Formula)
                {
                    if (cell.Content.StyleId == sheet.GetInheritedStyleId(position.Row, position.Column)) cells.Remove(position);
                    else cells[position] = new(cell.Content);
                    changed = true; continue;
                }
                var old = cell.Calculation;
                var state = old?.State == CalculationState.Blocked ? CalculationState.Blocked
                    : old?.Value is null ? CalculationState.Dirty : CalculationState.Stale;
                cells[position] = new(cell.Content, new(old?.Value, state, old?.InputRevision ?? snapshot.Revision,
                    old?.CalculationGeneration ?? 0, old?.ScriptGeneration ?? 0, old?.Diagnostics, old?.Array, old?.BlockedSpill), cell.SpillOwner);
                changed = true;
            }
            anyChanged |= changed;
            return changed ? sheet.WithCells(cells.ToImmutable()) : sheet;
        }).ToArray();
        return anyChanged ? snapshot.WithSheets(sheets) : snapshot;
    }

    private static WorkbookSnapshot RestoreCurrentCalculations(WorkbookSnapshot restored, WorkbookSnapshot current) =>
        restored.WithSheets(restored.Sheets.Select(sheet =>
        {
            var currentSheet = current.GetSheet(sheet.Id);
            var cells = sheet.Cells.ToBuilder();
            foreach (var (position, cell) in sheet.Cells)
                if (cell.Content.Kind == CellContentKind.Formula || cell.SpillOwner is not null)
                    cells[position] = new(cell.Content, currentSheet.GetCell(position.Row, position.Column)?.Calculation, cell.SpillOwner);
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
                if ((cell.Content.Kind == CellContentKind.Formula || cell.SpillOwner is not null) && cell.Calculation is { } cache &&
                    cache.InputRevision == previousRevision && cache.State is CalculationState.Fresh or CalculationState.Error)
                {
                    cells ??= sheet.Cells.ToBuilder();
                    cells[position] = new(cell.Content, new(cache.Value, cache.State, snapshot.Revision,
                        cache.CalculationGeneration, cache.ScriptGeneration, cache.Diagnostics, cache.Array, cache.BlockedSpill), cell.SpillOwner);
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
        _activeConsoleSource = null;
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
            _runtimeLifetime.Cancel(); _runtimeLifetime.Dispose();
            CancelActiveCalculation();
            _scriptsEnabled = false;
            _undo.Clear(); _redo.Clear(); _historyBytes = 0;
            _notifications.Clear();
            Changed = null;
            _scriptEvents.Clear(); _transientFiles = VirtualFileSet.Empty;
        }
    }

    private void ResetScriptRuntime()
    {
        var previous = _runtimeLifetime; _runtimeLifetime = new();
        previous.Cancel(); previous.Dispose();
    }

    internal void InvalidateScriptRuntime(CancellationToken lifetime)
    {
        lock (_gate)
            if (!_closed && lifetime.CanBeCanceled && lifetime == _runtimeLifetime.Token) ResetScriptRuntime();
    }

    private sealed record HistoryEntry(WorkbookSnapshot Before, WorkbookSnapshot After, ImmutableArray<CellChange> Changes, bool StructureChanged, bool CalculationUnchanged)
    {
        // Unchanged immutable input and metadata are shared. Count changed payloads and copied formula caches.
        public long Bytes => 1024L + Changes.Sum(change => Estimate(change.Before) + Estimate(change.After) + 256L)
            + FormulaCaches(Before) + FormulaCaches(After)
            + (ReferenceEquals(Before.PackageFiles, After.PackageFiles) ? 0 : Before.PackageFiles.Bytes + After.PackageFiles.Bytes)
            + (ReferenceEquals(Before.FileState, After.FileState) ? 0 : (Before.FileState?.OriginalWorkbook.Length ?? 0L) + (After.FileState?.OriginalWorkbook.Length ?? 0L))
            + Before.DefinedNames.Sum(n => 128L + 2L * (n.Name.Length + n.Formula.Length))
            + After.DefinedNames.Sum(n => 128L + 2L * (n.Name.Length + n.Formula.Length))
            + RemovedSheetMetadata(Before, After) + RemovedSheetMetadata(After, Before)
            + ChangedDataMetadata(Before, After) + ChangedDataMetadata(After, Before)
            + ChangedPresentationMetadata(Before, After) + ChangedPresentationMetadata(After, Before)
            + (StructureChanged ? 512L * Math.Abs(After.Sheets.Length - Before.Sheets.Length) : 0);
        private static long ChangedDataMetadata(WorkbookSnapshot from, WorkbookSnapshot to) => from.Sheets
            .Where(s => !ReferenceEquals(s.Data, to.Sheets.FirstOrDefault(other => other.Id == s.Id)?.Data))
            .Sum(s => 128L + s.Data.ThreadComments.Sum(c => 256L + 2L * (c.Content.Length + c.Author.Length + (c.UserId?.Length ?? 0) + (c.ProviderId?.Length ?? 0))) + s.Data.Pivots.Sum(p => p.EstimatedBytes) + s.Data.Charts.Sum(c => c.EstimatedBytes) + s.Data.Validations.Sum(v => v.EstimatedBytes) + s.Data.ConditionalFormats.Sum(v => v.EstimatedBytes) + (s.Data.Filter?.Columns.Values.Sum(c => 256L + (c.Values?.Sum(v => 32L + v.Length * 2L) ?? 0) +
                ((c.First?.Operand.Length ?? 0) + (c.Second?.Operand.Length ?? 0)) * 2L) ?? 0) +
                (s.Data.Sort?.Keys.Sum(k => 128L + k.CustomOrder.Sum(v => 32L + v.Length * 2L)) ?? 0) +
                s.Data.Tables.Sum(t => 1024L + t.Style.Elements.Length * 512L + t.Columns.Sum(c => 512L + 2L * (c.Name.Length + (c.CalculatedFormula?.Length ?? 0) + (c.TotalsFormula?.Length ?? 0) + (c.TotalsLabel?.Length ?? 0))) +
                    (t.Filter?.Columns.Values.Sum(c => 256L + (c.Values?.Sum(v => 32L + v.Length * 2L) ?? 0) + ((c.First?.Operand.Length ?? 0) + (c.Second?.Operand.Length ?? 0)) * 2L) ?? 0) +
                    (t.Sort?.Keys.Sum(k => 128L + k.CustomOrder.Sum(v => 32L + v.Length * 2L)) ?? 0)));
        private static long RemovedSheetMetadata(WorkbookSnapshot from, WorkbookSnapshot to) => from.Sheets
            .Where(s => !to.Sheets.Any(other => other.Id == s.Id))
            .Sum(s => 96L * (s.Rows.Count + s.Columns.Count) + s.Images.Sum(image => 128L + image.Data.Length)
                + s.MergedCells.Sum(merge => 32L + merge.Length * 2L) + Notes(s) + Printing(s));
        private static long Notes(WorksheetSnapshot s) => s.Annotations.NoteDrawing.Length + s.Annotations.Notes.Sum(n => 256L + n.Value.Author.Length * 2L + n.Value.Text.Text.Length * 2L + n.Value.Text.Runs.Length * 128L);
        private static long Printing(WorksheetSnapshot s) => 512L + s.Printing.PrinterSettings.Length + 48L * (s.Printing.Areas.Length + s.Printing.Titles.Length + s.Printing.RowBreaks.Length + s.Printing.ColumnBreaks.Length) +
            (s.Printing.HeaderFooter is { } h ? new[] { h.OddHeader, h.EvenHeader, h.FirstHeader, h.OddFooter, h.EvenFooter, h.FirstFooter }.Sum(t => (t?.Length ?? 0) * 2L) : 0);
        private static long ChangedPresentationMetadata(WorkbookSnapshot from, WorkbookSnapshot to)
        {
            long bytes = from.Styles.Length > to.Styles.Length ? (from.Styles.Length - to.Styles.Length) * 2048L : 0;
            foreach (var s in from.Sheets)
            {
                var other = to.Sheets.FirstOrDefault(t => t.Id == s.Id); if (other is null) continue;
                if (!s.Images.Equals(other.Images)) bytes += s.Images.Sum(i => 256L + (other.Images.Any(j => i.Data.Equals(j.Data)) ? 0 : i.Data.Length));
                bytes += s.Rows.Count(p => !other.Rows.TryGetValue(p.Key, out var value) || value != p.Value) * 192L;
                bytes += s.Columns.Count(p => !other.Columns.TryGetValue(p.Key, out var value) || value != p.Value) * 192L;
                if (!ReferenceEquals(s.Annotations, other.Annotations)) bytes += Notes(s);
                if (!ReferenceEquals(s.Printing, other.Printing)) bytes += Printing(s);
            }
            return bytes;
        }
        private static long FormulaCaches(WorkbookSnapshot snapshot) => snapshot.Sheets.Sum(sheet => sheet.Cells.Values
            .Where(cell => cell.Content.Kind == CellContentKind.Formula || cell.SpillOwner is not null).Sum(cell => 256L
                + (cell.Calculation?.Value is { Kind: CellValueKind.Text } value ? value.TextValue.Length * 2L : 0)
                + (cell.Calculation?.Array is { } array ? array.Values.Length * 32L + array.Values.Where(v => v.Kind == CellValueKind.Text).Sum(v => v.TextValue.Length * 2L) : 0)
                + (cell.Calculation?.Diagnostics.Sum(message => 32L + message.Length * 2L) ?? 0)));
        private static long Estimate(CellSnapshot? cell) => cell is null ? 0 : 256L
            + (cell.Content.Kind == CellContentKind.Formula ? cell.Content.FormulaText.Length * 2L
                : cell.Content.LiteralValue.Kind == CellValueKind.Text ? cell.Content.LiteralValue.TextValue.Length * 2L : 0);
    }
}
