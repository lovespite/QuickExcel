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
    private bool _scriptsEnabled;
    private bool _closed;
    private ImmutableArray<string> _diagnostics = [];
    private readonly Queue<DocumentChange> _notifications = new();
    private bool _publishing;

    public DocumentSession(WorkbookSnapshot snapshot, long historyBudgetBytes = 32 * 1024 * 1024)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(historyBudgetBytes);
        _snapshot = snapshot;
        _savedRevision = snapshot.Revision;
        HistoryBudgetBytes = historyBudgetBytes;
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

    public CommandResult Execute(DocumentCommand command, long? expectedRevision = null)
    {
        ArgumentNullException.ThrowIfNull(command);
        DocumentChange? change = null;
        CommandResult result;
        lock (_gate)
        {
            ThrowIfClosed();
            CheckRevision(expectedRevision);
            var before = _snapshot;
            var draft = command.Prepare(before); // No session state is changed until every check succeeds.
            if (ReferenceEquals(draft.Snapshot, before)) return new(false, before.Revision);
            var sourceChanged = command is UpdateScriptSourceCommand;
            var after = Invalidate(draft.Snapshot.WithRevision(checked(before.Revision + 1)));
            var entry = sourceChanged ? null : new HistoryEntry(before, after, draft.Changes, draft.StructureChanged);
            if (entry is not null && entry.Bytes > HistoryBudgetBytes)
                throw new InvalidOperationException("此操作超过撤销历史容量，尚未修改工作簿。请缩小操作范围。");
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
            _calculationGeneration = checked(_calculationGeneration + 1);
            _snapshot = after;
            _diagnostics = [];
            change = new(after.Revision, ActualChanges(before, after, draft.Changes), draft.StructureChanged);
            QueueNotification(change);
            result = new(true, after.Revision);
        }
        PublishNotifications();
        return result;
    }

    public Task<CommandResult> ExecuteAsync(DocumentCommand command, long? expectedRevision = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(Execute(command, expectedRevision));
    }

    public bool Undo(long? expectedRevision = null) => MoveHistory(redo: false, expectedRevision);
    public bool Redo(long? expectedRevision = null) => MoveHistory(redo: true, expectedRevision);

    private bool MoveHistory(bool redo, long? expectedRevision)
    {
        DocumentChange change;
        lock (_gate)
        {
            ThrowIfClosed();
            CheckRevision(expectedRevision);
            if (redo ? _redo.Count == 0 : _undo.Count == 0) return false;
            var before = _snapshot;
            var entry = redo ? _redo.Pop() : _undo.Last!.Value;
            if (!redo) _undo.RemoveLast();
            // Script text has its own editor history. Cell undo never reverts source or grants execution.
            var restored = (redo ? entry.After : entry.Before).WithScriptSource(before.ScriptSource)
                .WithRevision(checked(before.Revision + 1));
            _snapshot = Invalidate(restored);
            _calculationGeneration = checked(_calculationGeneration + 1);
            if (redo) _undo.AddLast(entry); else _redo.Push(entry);
            _diagnostics = [];
            change = new(_snapshot.Revision, ActualChanges(before, _snapshot, entry.Changes), entry.StructureChanged, IsUndoOrRedo: true);
            QueueNotification(change);
        }
        PublishNotifications();
        return true;
    }

    public void SetScriptsEnabled(bool enabled)
    {
        DocumentChange change;
        lock (_gate)
        {
            ThrowIfClosed();
            // Reapply also invalidates an in-flight compile of the same source.
            _scriptsEnabled = enabled;
            _scriptGeneration = checked(_scriptGeneration + 1);
            _calculationGeneration = checked(_calculationGeneration + 1);
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
            _calculationGeneration = checked(_calculationGeneration + 1);
            return new(new(SessionId, _snapshot.Revision, _calculationGeneration, _scriptGeneration), _snapshot, _scriptsEnabled);
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
                if (!builders.TryGetValue(sheet.Id, out var builder)) builders[sheet.Id] = builder = sheet.Cells.ToBuilder();
                var after = new CellSnapshot(before.Content, calculation);
                builder[new(address.Row, address.Column)] = after;
                changes.Add(new(address, before, after));
            }
            for (var index = 0; index < sheets.Length; index++)
                if (builders.TryGetValue(sheets[index].Id, out var builder)) sheets[index] = sheets[index].WithCells(builder.ToImmutable());
            _snapshot = _snapshot.WithSheets(sheets);
            _diagnostics = result.Diagnostics;
            _calculationGeneration = checked(_calculationGeneration + 1); // A completed request is single-use.
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
            _calculationGeneration = checked(_calculationGeneration + 1);
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
            before.GetSheet(change.Address.SheetId).GetCell(change.Address.Row, change.Address.Column),
            after.GetSheet(change.Address.SheetId).GetCell(change.Address.Row, change.Address.Column))).ToImmutableArray();

    private void CheckRevision(long? expected)
    {
        if (expected.HasValue && expected != _snapshot.Revision) throw new InvalidOperationException("文档已发生变化，请重新提交操作。");
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
            _scriptsEnabled = false;
            _undo.Clear(); _redo.Clear(); _historyBytes = 0;
            _notifications.Clear();
            Changed = null;
        }
    }

    private sealed record HistoryEntry(WorkbookSnapshot Before, WorkbookSnapshot After, ImmutableArray<CellChange> Changes, bool StructureChanged)
    {
        // Unchanged immutable input and metadata are shared. Count changed payloads and copied formula caches.
        public long Bytes => 1024L + Changes.Sum(change => Estimate(change.Before) + Estimate(change.After) + 256L)
            + FormulaCaches(Before) + FormulaCaches(After)
            + (StructureChanged ? 512L * Math.Abs(After.Sheets.Length - Before.Sheets.Length) : 0);
        private static long FormulaCaches(WorkbookSnapshot snapshot) => snapshot.Sheets.Sum(sheet => sheet.Cells.Values
            .Where(cell => cell.Content.Kind == CellContentKind.Formula).Sum(cell => 256L
                + (cell.Calculation?.Value is { Kind: CellValueKind.Text } value ? value.TextValue.Length * 2L : 0)
                + (cell.Calculation?.Diagnostics.Sum(message => 32L + message.Length * 2L) ?? 0)));
        private static long Estimate(CellSnapshot? cell) => cell is null ? 0 : 256L
            + (cell.Content.Kind == CellContentKind.Formula ? cell.Content.FormulaText.Length * 2L
                : cell.Content.LiteralValue.Kind == CellValueKind.Text ? cell.Content.LiteralValue.TextValue.Length * 2L : 0);
    }
}
