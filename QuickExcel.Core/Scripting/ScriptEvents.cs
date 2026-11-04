using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed record ScriptEvent(string Name, Guid SessionId, long Sequence, long Revision,
    string? SheetId, string? SheetName, ImmutableArray<string> Ranges, bool IsUndoOrRedo = false);

public static class ScriptEventNames
{
    public const string Prefix = "@event/";
    public static readonly ImmutableHashSet<string> All = new[] {
        "workbook.open", "workbook.beforeClose", "workbook.close", "workbook.activate", "workbook.deactivate",
        "workbook.beforeSave", "workbook.saved", "workbook.calculated", "sheet.added", "sheet.deleted", "sheet.renamed",
        "sheet.activate", "sheet.deactivate", "selection.changed", "cell.click", "cell.edit", "cell.commit", "cells.changed"
    }.ToImmutableHashSet(StringComparer.Ordinal);
}

public sealed partial class DocumentSession
{
    private readonly Queue<ScriptEvent> _scriptEvents = new();
    private long _eventSequence;
    private bool _openedEvent;
    private HashSet<string>? _registeredEvents;
    private VirtualFileSet _transientFiles = VirtualFileSet.Empty;
    public VirtualFileSet TransientFiles { get { lock (_gate) { ThrowIfClosed(); return _transientFiles; } } }

    /// <summary>UI notifications are queued, never executed on a command's publication stack.</summary>
    public void RaiseScriptEvent(string name, SheetId? sheetId = null, IEnumerable<string>? ranges = null)
    {
        lock (_gate)
        {
            ThrowIfClosed();
            if (!ScriptEventNames.All.Contains(name)) throw new ArgumentException("未知脚本事件。");
            var sheet = sheetId is { } id ? _snapshot.GetSheet(id) : null;
            var addresses = (ranges ?? []).Take(65).ToImmutableArray();
            if (addresses.Length > 64) throw new ArgumentException("事件最多携带 64 个选区。");
            foreach (var address in addresses) _ = CellRectangle.Parse(address);
            EnqueueScriptEvent(name, sheet, addresses);
        }
    }

    private void EnqueueScriptEvent(string name, WorksheetSnapshot? sheet = null, ImmutableArray<string> ranges = default, bool history = false)
    {
        if (!_scriptsEnabled || !_snapshot.ScriptModules.HasActions || _registeredEvents is not null && !_registeredEvents.Contains(name)) return;
        if (_scriptEvents.Count >= 256)
        {
            _scriptEvents.Clear(); _diagnostics = _diagnostics.Add("脚本事件队列超过 256 条，待处理事件已丢弃。");
        }
        _scriptEvents.Enqueue(new(name, SessionId, ++_eventSequence, _snapshot.Revision, sheet?.Id.ToString(), sheet?.Name,
            ranges.IsDefault ? [] : ranges, history));
    }

    public bool HasScriptEvents { get { lock (_gate) return !_closed && _scriptEvents.Count > 0; } }

    public void ConfigureScriptEvents(ScriptModuleSet modules, IEnumerable<string> names)
    {
        lock (_gate)
        {
            ThrowIfClosed();
            if (!ReferenceEquals(modules, _snapshot.ScriptModules) || !_scriptsEnabled) return;
            _registeredEvents = names.Where(ScriptEventNames.All.Contains).ToHashSet(StringComparer.Ordinal);
            var retained = _scriptEvents.Where(e => _registeredEvents.Contains(e.Name)).ToArray();
            _scriptEvents.Clear(); foreach (var item in retained) _scriptEvents.Enqueue(item);
        }
    }

    public void QueueCloseEvent()
    {
        lock (_gate) { ThrowIfClosed(); _scriptEvents.Clear(); EnqueueScriptEvent("workbook.close"); }
    }

    public ScriptEvent? TakeScriptEvent()
    {
        lock (_gate) { ThrowIfClosed(); return _scriptEvents.TryDequeue(out var item) ? item : null; }
    }

    private void QueueChangeEvents(WorkbookSnapshot before, WorkbookSnapshot after, ImmutableArray<CellChange> cells, bool history = false)
    {
        foreach (var sheet in before.Sheets.Where(s => after.Sheets.All(a => a.Id != s.Id))) EnqueueScriptEvent("sheet.deleted", sheet, history: history);
        foreach (var sheet in after.Sheets)
        {
            var old = before.Sheets.FirstOrDefault(s => s.Id == sheet.Id);
            if (old is null) EnqueueScriptEvent("sheet.added", sheet, history: history);
            else if (old.Name != sheet.Name) EnqueueScriptEvent("sheet.renamed", sheet, history: history);
        }
        if (before.ActiveSheetId != after.ActiveSheetId)
        {
            EnqueueScriptEvent("sheet.deactivate", before.GetSheet(before.ActiveSheetId));
            EnqueueScriptEvent("sheet.activate", after.GetSheet(after.ActiveSheetId));
        }
        foreach (var group in cells.Where(c => c.Before?.Content != c.After?.Content).GroupBy(c => c.Address.SheetId))
        {
            var sheet = after.Sheets.FirstOrDefault(s => s.Id == group.Key);
            if (sheet is null) continue;
            var range = new CellRectangle(group.Min(c => c.Address.Row), group.Min(c => c.Address.Column), group.Max(c => c.Address.Row), group.Max(c => c.Address.Column));
            EnqueueScriptEvent("cells.changed", sheet, [range.ToString()], history);
        }
    }
}
