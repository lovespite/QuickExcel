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
    public SetCellsCommand(IEnumerable<CellEdit> edits, bool allowRichTextLoss = false)
    {
        ArgumentNullException.ThrowIfNull(edits);
        Edits = edits.Take(WorkbookLimits.MaxCells + 1).ToImmutableArray();
        if (Edits.Length > WorkbookLimits.MaxCells) throw new ArgumentException(MessageText.Get("编辑数量超过工作簿容量。"), nameof(edits));
        _allowRichTextLoss = allowRichTextLoss;
    }
    public ImmutableArray<CellEdit> Edits { get; }
    private readonly bool _allowRichTextLoss;

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
            if (!addresses.Add(edit.Address)) throw new ArgumentException(MessageText.Get("同一事务不能多次指定同一单元格。"));
            var sheet = snapshot.GetSheet(edit.Address.SheetId);
            if (edit.Content.StyleId >= snapshot.Styles.Length) throw new ArgumentException(MessageText.Get("单元格样式不存在。"));
            var text = edit.Content.Kind == CellContentKind.Formula ? edit.Content.FormulaText
                : edit.Content.LiteralValue.Kind == CellValueKind.Text ? edit.Content.LiteralValue.TextValue : "";
            if (text.Length > WorkbookLimits.MaxCellTextLength) throw new ArgumentException(MessageText.Get("单元格最多支持 32767 个字符，整批编辑尚未应用。"));
            var before = sheet.GetCell(edit.Address.Row, edit.Address.Column);
            if (before?.SpillOwner is { } owner && (before.Content.StyleId == edit.Content.StyleId || before.Content.WithStyle(0) != edit.Content.WithStyle(0)) &&
                !Edits.Any(e => e.Address == owner && e.Content.WithStyle(0) != sheet.GetCell(owner.Row, owner.Column)?.Content.WithStyle(0)))
                throw new InvalidOperationException(MessageText.Get("不能单独修改溢出结果，请编辑数组左上角公式；整批修改未应用。"));
            if (!_allowRichTextLoss && before?.Content.RichText is not null && edit.Content.RichText is null)
            {
                if (edit.Content.Kind == CellContentKind.Literal && edit.Content.LiteralValue == before.Value && edit.Content.StyleId == before.Content.StyleId) continue;
                throw new InvalidOperationException(MessageText.Get("目标包含富文本。请先明确转换为普通文本，再编辑或粘贴；整批操作未应用。"));
            }
            if (before?.Content == edit.Content) continue;
            if (edit.Content.Kind != CellContentKind.Literal || edit.Content.LiteralValue.Kind != CellValueKind.Blank)
                foreach (var merge in sheet.MergedCells.Select(CellRectangle.Parse))
                    if (merge.Contains(edit.Address.Row, edit.Address.Column) && (edit.Address.Row != merge.Top || edit.Address.Column != merge.Left))
                        throw new InvalidOperationException(MessageText.Get("不能写入合并区域的非主单元格，请先取消合并；整批操作未应用。"));
            var inheritedStyle = sheet.GetInheritedStyleId(edit.Address.Row, edit.Address.Column);
            var after = before?.SpillOwner is null && edit.Content.Kind == CellContentKind.Literal && edit.Content.LiteralValue.Kind == CellValueKind.Blank && edit.Content.StyleId == 0 && inheritedStyle == 0
                ? null : new CellSnapshot(edit.Content,
                    before?.SpillOwner is not null && before.Content.WithStyle(0) == edit.Content.WithStyle(0) ? before.Calculation :
                    before?.Content.Kind == CellContentKind.Formula && edit.Content.Kind == CellContentKind.Formula &&
                    before.Content.FormulaText == edit.Content.FormulaText ? before.Calculation : null,
                    before?.SpillOwner is not null && before.Content.WithStyle(0) == edit.Content.WithStyle(0) ? before.SpillOwner : null);
            if (before is null && after is null) continue;
            if (!builders.TryGetValue(sheet.Id, out var builder)) builders[sheet.Id] = builder = sheet.Cells.ToBuilder();
            var position = new CellPosition(edit.Address.Row, edit.Address.Column);
            if (after is null) builder.Remove(position); else builder[position] = after;
            changes.Add(new(edit.Address, before, after));
        }
        if (changes.Count == 0) return new(snapshot, [], false);
        for (var index = 0; index < sheets.Length; index++)
            if (builders.TryGetValue(sheets[index].Id, out var builder)) sheets[index] = sheets[index].WithCells(builder.ToImmutable());
        var activated = sheets.SelectMany(s => s.Data.Tables.Where(t => t.InsertRow && changes.Any(p => p.Address.SheetId == s.Id && t.Range.Contains(p.Address.Row, p.Address.Column) &&
            p.Address.Row >= t.Range.Top + (t.HeaderRow ? 1 : 0) && p.After is { } cell && (cell.Content.Kind == CellContentKind.Formula || cell.Value.Kind != CellValueKind.Blank))).Select(t => (Sheet: s.Id, Table: t))).ToArray();
        if (activated.Length > 0)
        {
            var after = snapshot.WithSheets(sheets.Select(s => s.WithData(s.Data with { Tables = s.Data.Tables.Select(t => activated.Any(a => a.Table.Id == t.Id) ? t.WithInsertRow(false) : t).ToImmutableArray() })));
            foreach (var group in activated.GroupBy(a => a.Sheet))
                after = TableStructure.FillInserted(after, snapshot.GetSheet(group.Key), p => group.Any(a => a.Table.Range.Contains(p.Row, p.Column) && p.Row >= a.Table.Range.Top + (a.Table.HeaderRow ? 1 : 0)));
            return StructureEdits.Draft(snapshot, after);
        }
        return new(snapshot.WithSheets(sheets), changes.ToImmutableArray(), false);
    }
}

public sealed class ClearCellsCommand : DocumentCommand
{
    public ClearCellsCommand(IEnumerable<CellAddress> addresses) => Addresses = addresses.Take(WorkbookLimits.MaxCells + 1).ToImmutableArray();
    public ImmutableArray<CellAddress> Addresses { get; }
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (Addresses.Length > WorkbookLimits.MaxCells) throw new ArgumentException(MessageText.Get("编辑数量超过工作簿容量。"));
        var seen = new HashSet<CellAddress>();
        var edits = new List<CellEdit>();
        foreach (var address in Addresses)
        {
            if (!seen.Add(address)) throw new ArgumentException(MessageText.Get("同一事务不能多次指定同一单元格。"));
            var cell = snapshot.GetSheet(address.SheetId).GetCell(address.Row, address.Column);
            // Clearing an absent cell must not create a General override of row/column formatting.
            if (cell is not null) edits.Add(new(address, CellContent.FromLiteral(CellValue.Blank, cell.Content.StyleId)));
        }
        return new SetCellsCommand(edits, allowRichTextLoss: true).Prepare(snapshot);
    }
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
        if ((long)Encoding.UTF8.GetByteCount(Source) + snapshot.ScriptModules.Modules.Where(m => m.Key != snapshot.ScriptModules.EntryModules[0]).Sum(m => (long)Encoding.UTF8.GetByteCount(m.Value.Source)) > 1024 * 1024) throw new ArgumentException(MessageText.Get("脚本大小不能超过 1 MiB。"));
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
            throw new ArgumentException(MessageText.Get("脚本总量不能超过 1 MiB。"));
        var prior = snapshot.ScriptModules;
        var sameBindings = prior.ScriptApiVersion == modules.ScriptApiVersion && prior.EntryModules.SequenceEqual(modules.EntryModules)
            && prior.Bindings.ToHashSet().SetEquals(modules.Bindings)
            && prior.Modules.Count == modules.Modules.Count && prior.Modules.All(m => modules.Modules.TryGetValue(m.Key, out var value) && value.Path == m.Value.Path && value.Language == m.Value.Language);
        if (sameBindings && prior.Modules.All(m => modules.Modules[m.Key].Source == m.Value.Source)) return new(snapshot, [], false);
        var updated = new ScriptModuleSet(modules.Modules, modules.EntryModules, checked(prior.Revision + 1),
            sameBindings ? prior.BindingRevision : checked(prior.BindingRevision + 1), modules.ScriptApiVersion, modules.Bindings);
        return new(snapshot.WithScriptModules(updated), [], false);
    }
}

/// <summary>Attaches a signature to the exact source capture that was signed; never grants execution.</summary>
public sealed class UpdateScriptSignatureCommand(ScriptModuleSet expectedModules, ScriptSignature? signature) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (!ReferenceEquals(snapshot.ScriptModules, expectedModules)) throw new InvalidOperationException(MessageText.Get("脚本已改变，请重新签名。"));
        return new(snapshot.ScriptModules.Signature == signature ? snapshot : snapshot.WithScriptModules(expectedModules.WithSignature(signature)), [], false);
    }
}

internal sealed record CommandDraft(WorkbookSnapshot Snapshot, ImmutableArray<CellChange> Changes, bool StructureChanged,
    bool CalculationUnchanged = false);

public sealed record DocumentChange(long Revision, ImmutableArray<CellChange> Cells, bool StructureChanged,
    bool IsUndoOrRedo = false, bool CalculationOnly = false)
{
    /// <summary>The exact committed state for this event, even if another writer has already advanced the session.</summary>
    public WorkbookSnapshot? Snapshot { get; init; }
    public Guid SessionId { get; init; }
    public bool CalculationUnchanged { get; init; }
}

public sealed record CommandResult(bool Changed, long Revision)
{
    public bool CalculationUnchanged { get; init; }
}
