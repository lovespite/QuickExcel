using System.Collections.Immutable;
using System.Text;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed record CellEdit(CellAddress Address, CellContent Content);
public sealed record CellChange(CellAddress Address, CellSnapshot? Before, CellSnapshot? After);

/// <summary>Commands are immutable requests; every item is validated before the session changes.</summary>
public abstract class DocumentCommand
{
    internal abstract CommandDraft Prepare(WorkbookSnapshot snapshot);
}

public sealed class SetCellsCommand : DocumentCommand
{
    public SetCellsCommand(IEnumerable<CellEdit> edits)
    {
        ArgumentNullException.ThrowIfNull(edits);
        Edits = edits.Take(WorkbookLimits.MaxCells + 1).ToImmutableArray();
        if (Edits.Length > WorkbookLimits.MaxCells) throw new ArgumentException("编辑数量超过工作簿容量。", nameof(edits));
    }
    public ImmutableArray<CellEdit> Edits { get; }

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var addresses = new HashSet<CellAddress>();
        var changes = new List<CellChange>();
        var sheets = snapshot.Sheets.ToArray();
        var builders = new Dictionary<SheetId, ImmutableDictionary<CellPosition, CellSnapshot>.Builder>();
        foreach (var edit in Edits)
        {
            ArgumentNullException.ThrowIfNull(edit);
            ArgumentNullException.ThrowIfNull(edit.Content);
            WorkbookLimits.ValidatePosition(edit.Address.Row, edit.Address.Column);
            if (!addresses.Add(edit.Address)) throw new ArgumentException("同一事务不能多次指定同一单元格。");
            var sheet = snapshot.GetSheet(edit.Address.SheetId);
            if (edit.Content.StyleId >= snapshot.Styles.Length) throw new ArgumentException("单元格样式不存在。");
            var text = edit.Content.Kind == CellContentKind.Formula ? edit.Content.FormulaText
                : edit.Content.LiteralValue.Kind == CellValueKind.Text ? edit.Content.LiteralValue.TextValue : "";
            if (text.Length > WorkbookLimits.MaxCellTextLength) throw new ArgumentException("单元格最多支持 32767 个字符，整批编辑尚未应用。");
            var before = sheet.GetCell(edit.Address.Row, edit.Address.Column);
            if (before?.Content == edit.Content) continue;
            var after = edit.Content.Kind == CellContentKind.Literal && edit.Content.LiteralValue.Kind == CellValueKind.Blank && edit.Content.StyleId == 0
                ? null : new CellSnapshot(edit.Content);
            if (before is null && after is null) continue;
            if (!builders.TryGetValue(sheet.Id, out var builder)) builders[sheet.Id] = builder = sheet.Cells.ToBuilder();
            var position = new CellPosition(edit.Address.Row, edit.Address.Column);
            if (after is null) builder.Remove(position); else builder[position] = after;
            changes.Add(new(edit.Address, before, after));
        }
        if (changes.Count == 0) return new(snapshot, [], false);
        for (var index = 0; index < sheets.Length; index++)
            if (builders.TryGetValue(sheets[index].Id, out var builder)) sheets[index] = sheets[index].WithCells(builder.ToImmutable());
        return new(snapshot.WithSheets(sheets), changes.ToImmutableArray(), false);
    }
}

public sealed class ClearCellsCommand : DocumentCommand
{
    public ClearCellsCommand(IEnumerable<CellAddress> addresses) => Addresses = addresses.Take(WorkbookLimits.MaxCells + 1).ToImmutableArray();
    public ImmutableArray<CellAddress> Addresses { get; }
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot) => new SetCellsCommand(Addresses.Select(address =>
        new CellEdit(address, CellContent.FromLiteral(CellValue.Blank, snapshot.GetSheet(address.SheetId).GetCell(address.Row, address.Column)?.Content.StyleId ?? 0)))).Prepare(snapshot);
}

public sealed class AddSheetCommand : DocumentCommand
{
    public AddSheetCommand(string name) { Name = name; Id = SheetId.New(); }
    public string Name { get; }
    public SheetId Id { get; }
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot) =>
        new(snapshot.WithSheets(snapshot.Sheets.Add(new WorksheetSnapshot(Id, Name))), [], true);
}

public sealed class UpdateScriptSourceCommand : DocumentCommand
{
    public UpdateScriptSourceCommand(string source)
    {
        ArgumentNullException.ThrowIfNull(source);
        Source = source;
    }
    public string Source { get; }
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if ((long)Encoding.UTF8.GetByteCount(Source) + snapshot.ScriptModules.Modules.Where(m => m.Key != snapshot.ScriptModules.EntryModules[0]).Sum(m => (long)Encoding.UTF8.GetByteCount(m.Value.Source)) > 1024 * 1024) throw new ArgumentException("脚本大小不能超过 1 MiB。");
        return new(snapshot.ScriptSource == Source ? snapshot : snapshot.WithScriptSource(Source), [], false);
    }
}

/// <summary>Replaces package source/bindings without carrying execution authorization or editor history.</summary>
public sealed class UpdateScriptModulesCommand(ScriptModuleSet modules) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(modules);
        if (modules.Modules.Sum(m => (long)Encoding.UTF8.GetByteCount(m.Value.Source)) > 1024 * 1024)
            throw new ArgumentException("脚本总量不能超过 1 MiB。");
        var prior = snapshot.ScriptModules;
        var sameBindings = prior.ScriptApiVersion == modules.ScriptApiVersion && prior.EntryModules.SequenceEqual(modules.EntryModules)
            && prior.Modules.Count == modules.Modules.Count && prior.Modules.All(m => modules.Modules.TryGetValue(m.Key, out var value) && value.Path == m.Value.Path);
        if (sameBindings && prior.Modules.All(m => modules.Modules[m.Key].Source == m.Value.Source)) return new(snapshot, [], false);
        var updated = new ScriptModuleSet(modules.Modules, modules.EntryModules, checked(prior.Revision + 1),
            sameBindings ? prior.BindingRevision : checked(prior.BindingRevision + 1), modules.ScriptApiVersion);
        return new(snapshot.WithScriptModules(updated), [], false);
    }
}

internal sealed record CommandDraft(WorkbookSnapshot Snapshot, ImmutableArray<CellChange> Changes, bool StructureChanged);

public sealed record DocumentChange(long Revision, ImmutableArray<CellChange> Cells, bool StructureChanged,
    bool IsUndoOrRedo = false, bool CalculationOnly = false)
{
    /// <summary>The exact committed state for this event, even if another writer has already advanced the session.</summary>
    public WorkbookSnapshot? Snapshot { get; init; }
    public Guid SessionId { get; init; }
}

public sealed record CommandResult(bool Changed, long Revision);
