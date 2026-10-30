using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public sealed class ReorderSheetCommand(SheetId sheetId, int destination) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (destination < 0 || destination >= snapshot.Sheets.Length) throw new ArgumentOutOfRangeException(nameof(destination));
        var sheet = snapshot.GetSheet(sheetId); var index = snapshot.Sheets.IndexOf(sheet);
        if (index == destination) return new(snapshot, [], false);
        return new(snapshot.WithSheets(snapshot.Sheets.RemoveAt(index).Insert(destination, sheet)), [], true);
    }
}

public sealed class CopySheetCommand : DocumentCommand
{
    public CopySheetCommand(SheetId source, string name, int destination) { Source = source; Name = name; Destination = destination; Id = SheetId.New(); }
    public SheetId Source { get; }
    public string Name { get; }
    public int Destination { get; }
    public SheetId Id { get; }
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (Destination < 0 || Destination > snapshot.Sheets.Length) throw new ArgumentOutOfRangeException(nameof(Destination));
        WorkbookLimits.ValidateSheetName(Name);
        if (snapshot.Sheets.Any(s => s.Name.Equals(Name, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("工作表名称已存在。");
        if (snapshot.Sheets.Length >= WorkbookLimits.MaxSheets) throw new InvalidOperationException("工作表数量超过 64。");
        StructureEdits.RequireTransformableFunctions(snapshot);
        var source = snapshot.GetSheet(Source);
        if (snapshot.Sheets.Sum(s => (long)s.Cells.Count) + source.Cells.Count > WorkbookLimits.MaxCells) throw new InvalidOperationException("复制工作表超过单元格总容量。");
        var expressions = snapshot.DefinedNames.ToDictionary(n => n, n => FormulaSyntax.Parse(n.Formula));
        var dependents = new Dictionary<DefinedName, List<DefinedName>>(); var relevant = new HashSet<DefinedName>(); var pending = new Queue<DefinedName>();
        foreach (var definition in snapshot.DefinedNames)
        {
            var nodes = FormulaSyntax.Walk(expressions[definition]).ToArray();
            if (nodes.OfType<ReferenceNode>().Any(r => string.Equals(r.Sheet, source.Name, StringComparison.OrdinalIgnoreCase)))
            { relevant.Add(definition); pending.Enqueue(definition); }
            foreach (var node in nodes.OfType<NameNode>())
                if (FormulaBinding.ResolveName(snapshot, definition.Scope ?? Source, node, inDefinition: true) is { } dependency)
                {
                    if (!dependents.TryGetValue(dependency, out var list)) dependents[dependency] = list = [];
                    list.Add(definition);
                }
        }
        while (pending.TryDequeue(out var dependency))
            foreach (var dependent in dependents.GetValueOrDefault(dependency) ?? [])
                if (relevant.Add(dependent)) pending.Enqueue(dependent);
        var localNames = snapshot.DefinedNames.Where(n => n.Scope == Source).ToArray();
        var copiedNames = localNames.Concat(snapshot.DefinedNames.Where(n => n.Scope is null && relevant.Contains(n) &&
            !localNames.Any(l => l.Name.Equals(n.Name, StringComparison.OrdinalIgnoreCase)))).ToHashSet();
        FormulaNode Remap(FormulaNode node, bool definition, SheetId context)
        {
            if (node is ReferenceNode r && string.Equals(r.Sheet, source.Name, StringComparison.OrdinalIgnoreCase)) return r with { Sheet = Name };
            if (node is NameNode n)
            {
                if (string.Equals(n.Sheet, source.Name, StringComparison.OrdinalIgnoreCase)) return n with { Sheet = Name };
                if (definition && n.Sheet is null && FormulaBinding.ResolveName(snapshot, context, n, inDefinition: true) is { } bound && copiedNames.Contains(bound))
                    return n with { Sheet = Name };
            }
            return node;
        }
        string Formula(string text, bool definition, SheetId context)
        {
            var changed = false;
            var expression = FormulaSyntax.Map(FormulaSyntax.Parse(text), node => { var next = Remap(node, definition, context); changed |= next != node; return next; });
            return changed ? "=" + FormulaSyntax.Print(expression) : text;
        }
        var cells = source.Cells.Select(p => new KeyValuePair<CellPosition, CellSnapshot>(p.Key,
            p.Value.Content.Kind == CellContentKind.Formula ? new(CellContent.FromFormula(Formula(p.Value.Content.FormulaText, false, Source), p.Value.Content.StyleId)) : p.Value)).ToImmutableDictionary();
        var copy = new WorksheetSnapshot(Id, Name, cells, source.Rows, source.Columns, source.MergedCells, source.Images,
            source.ShowGridLines, source.DefaultRowHeight, source.TabColor, source.View, source.Visibility, source.Printing, source.Annotations);
        var names = snapshot.DefinedNames.Concat(copiedNames.Select(n => new DefinedName(n.Name, Formula(n.Formula, true, n.Scope ?? Source), Id)));
        var after = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, snapshot.Sheets.Insert(Destination, copy), snapshot.ScriptSource,
            snapshot.Styles, snapshot.ScriptModules, snapshot.FileState, names, copy.Visibility == SheetVisibility.Visible ? Id : snapshot.ActiveSheetId, snapshot.Window);
        return new(after, cells.Select(p => new CellChange(new(Id, p.Key.Row, p.Key.Column), null, p.Value)).ToImmutableArray(), true);
    }
}

public sealed class RenameDefinedNameCommand(string name, SheetId? scope, string newName) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var definition = snapshot.DefinedNames.SingleOrDefault(n => n.Scope == scope && n.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("名称不存在。");
        if (definition.Name == newName) return new(snapshot, [], false);
        if (snapshot.DefinedNames.Any(n => n != definition && n.Name.Equals(newName, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException("新名称已在某个作用域存在；为避免改变其他引用的绑定，本次改名未应用。");
        StructureEdits.RequireTransformableFunctions(snapshot);
        string Formula(string text, SheetId context, bool inDefinition)
        {
            var changed = false;
            var expression = FormulaSyntax.Map(FormulaSyntax.Parse(text), node =>
            {
                if (node is NameNode n && FormulaBinding.ResolveName(snapshot, context, n, inDefinition) == definition)
                { changed = true; return n with { Name = newName }; }
                return node;
            });
            return changed ? "=" + FormulaSyntax.Print(expression) : text;
        }
        var sheets = snapshot.Sheets.Select(sheet => sheet.WithCells(sheet.Cells.ToImmutableDictionary(p => p.Key, p =>
        {
            if (p.Value.Content.Kind != CellContentKind.Formula) return p.Value;
            var next = Formula(p.Value.Content.FormulaText, sheet.Id, false);
            return next == p.Value.Content.FormulaText ? p.Value : new(CellContent.FromFormula(next, p.Value.Content.StyleId));
        })));
        var names = snapshot.DefinedNames.Select(n => n with { Name = n == definition ? newName : n.Name, Formula = Formula(n.Formula, n.Scope ?? snapshot.Sheets[0].Id, true) });
        var after = new WorkbookSnapshot(snapshot.DocumentId, snapshot.Revision, snapshot.DateSystem, sheets, snapshot.ScriptSource,
            snapshot.Styles, snapshot.ScriptModules, snapshot.FileState, names, snapshot.ActiveSheetId, snapshot.Window);
        return StructureEdits.Draft(snapshot, after);
    }
}

public sealed class DeleteDefinedNameCommand(string name, SheetId? scope) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var definition = snapshot.DefinedNames.SingleOrDefault(n => n.Scope == scope && n.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("名称不存在。");
        if (scope is { } sheetId)
        {
            var sheet = snapshot.GetSheet(sheetId);
            var formulas = snapshot.Sheets.SelectMany(s => s.Cells.Values).Where(c => c.Content.Kind == CellContentKind.Formula).Select(c => c.Content.FormulaText)
                .Concat(snapshot.DefinedNames.Where(n => n != definition).Select(n => n.Formula));
            if (formulas.SelectMany(f => FormulaSyntax.Walk(FormulaSyntax.Parse(f))).OfType<NameNode>().Any(n => n.Name.Equals(name, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(n.Sheet, sheet.Name, StringComparison.OrdinalIgnoreCase)))
                throw new NotSupportedException("该局部名称仍有带工作表限定的引用。Excel 删除后的工作簿限定名称尚未建模，请先修改这些引用；名称未删除。");
        }
        return StructureEdits.Draft(snapshot, snapshot.WithDefinedNames(snapshot.DefinedNames.Where(n => n != definition)));
    }
}

/// <summary>Rename and replace a definition in the same history entry; scope stays fixed.</summary>
public sealed class EditDefinedNameCommand(string oldName, SheetId? scope, string name, string formula) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var original = snapshot.DefinedNames.SingleOrDefault(n => n.Scope == scope && n.Name.Equals(oldName, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("名称不存在。");
        var renamed = new RenameDefinedNameCommand(oldName, scope, name).Prepare(snapshot).Snapshot;
        var nextFormula = formula == original.Formula ? renamed.DefinedNames.Single(n => n.Scope == scope && n.Name == name).Formula : formula;
        var updated = new DefineNameCommand(new(name, nextFormula, scope)).Prepare(renamed).Snapshot;
        return ReferenceEquals(snapshot, updated) ? new(snapshot, [], false) : StructureEdits.Draft(snapshot, updated);
    }
}
