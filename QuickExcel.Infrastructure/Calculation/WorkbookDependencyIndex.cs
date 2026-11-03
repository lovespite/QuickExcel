using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Calculation;

/// <summary>Sparse reference index. Rectangle intervals are indexed by row; no empty cells are allocated.</summary>
public sealed class WorkbookDependencyIndex
{
    internal sealed class Node(CellSnapshot cell, int order)
    {
        public CellSnapshot Cell { get; } = cell;
        public int Order { get; } = order;
        public FormulaNode Expression { get; set; } = new ValueNode(CellValue.FromError(CellErrorCode.Value));
        public string? ParseError { get; set; }
        public bool Script { get; set; }
        public bool Volatile { get; set; }
        public bool VisibilityBlocked { get; set; }
        public string? PortableFormula { get; set; }
        public HashSet<CellAddress> Dependencies { get; } = [];
        public HashSet<CellAddress> Dependents { get; } = [];
    }

    internal Dictionary<CellAddress, Node> Nodes { get; } = [];
    private readonly Dictionary<SheetId, Interval?> _ranges = [];
    public IReadOnlySet<CellAddress> CycleMembers { get; }
    public IReadOnlySet<CellAddress> NonPortableCells { get; }

    public WorkbookDependencyIndex(WorkbookSnapshot snapshot, CancellationToken cancellationToken = default)
        : this(snapshot, () => cancellationToken.ThrowIfCancellationRequested()) { }

    internal WorkbookDependencyIndex(WorkbookSnapshot snapshot, Action check)
        : this(snapshot, check, null) { }

    internal WorkbookDependencyIndex(WorkbookSnapshot snapshot, Action check, IReadOnlyDictionary<CellAddress, QuickExcel.Core.FormulaSpill>? spills)
    {
        var sheets = snapshot.Sheets.ToDictionary(s => s.Name, StringComparer.OrdinalIgnoreCase);
        var references = new Dictionary<SheetId, List<Entry>>();
        var order = 0; var referenceCount = 0;
        foreach (var sheet in snapshot.Sheets)
        foreach (var (position, cell) in sheet.Cells.OrderBy(p => p.Key.Row).ThenBy(p => p.Key.Column))
        {
            check();
            if (cell.Content.Kind != CellContentKind.Formula) continue;
            var address = new CellAddress(sheet.Id, position.Row, position.Column);
            var node = new Node(cell, order++); Nodes.Add(address, node);
            try
            {
                node.Expression = FormulaBinding.Bind(FormulaSyntax.Parse(cell.Content.FormulaText), snapshot, sheet.Id, position, cell.Content.DynamicArray);
                node.VisibilityBlocked = SubtotalFunctions.Contains(node.Expression) && SubtotalFunctions.HasUnmodeledFilters(snapshot);
                var depthCheck = new Stack<(FormulaNode Node, int Depth)>();
                depthCheck.Push((node.Expression, 0));
                while (depthCheck.TryPop(out var item))
                {
                    check();
                    if (item.Depth > 128) throw new FormatException("名称展开后公式嵌套超过 128 层。");
                    switch (item.Node)
                    {
                        case UnaryNode u: depthCheck.Push((u.Operand, item.Depth + 1)); break;
                        case BinaryNode b: depthCheck.Push((b.Left, item.Depth + 1)); depthCheck.Push((b.Right, item.Depth + 1)); break;
                        case FunctionNode f: foreach (var a in f.Arguments) depthCheck.Push((a, item.Depth + 1)); break;
                        case LetNode let:
                            depthCheck.Push((let.Body, item.Depth + 1));
                            foreach (var b in let.Bindings) depthCheck.Push((b.Value, item.Depth + 1));
                            break;
                    }
                }
                foreach (var part in FormulaSyntax.Walk(ConditionalRanges.ForDependencies(node.Expression)).Concat(SubtotalFunctions.FilterDependencies(node.Expression, snapshot)))
                {
                    check();
                    if (part is FunctionNode function)
                    {
                        node.Script |= !StandardFunctions.Signatures.ContainsKey(function.Name);
                        node.Volatile |= function.Name is "TODAY" or "NOW" or "SUBTOTAL";
                    }
                    if (part is not ReferenceNode reference) continue;
                    node.Volatile |= reference.SourceExpression is FunctionNode { Name: "OFFSET" or "INDIRECT" };
                    var target = sheets[reference.Sheet!].Id;
                    if (!references.TryGetValue(target, out var entries)) references[target] = entries = [];
                    if (++referenceCount > 2_000_000) throw new InvalidOperationException("公式引用索引超过 2,000,000 项。");
                    entries.Add(new(reference, address));
                }
                if (FormulaSyntax.TryPrintPortable(node.Expression, out var portable)) node.PortableFormula = portable;
            }
            catch (Exception e) when (e is FormatException or OverflowException)
            { node.ParseError = e.Message; node.Expression = new ValueNode(CellValue.FromError(CellErrorCode.Value)); }
        }
        foreach (var (sheet, entries) in references) { check(); _ranges[sheet] = Interval.Build(entries); }
        var edges = 0;
        foreach (var (address, node) in Nodes)
        {
            check();
            foreach (var dependent in DirectDependents(address, check))
            {
                if (!node.Dependents.Add(dependent)) continue;
                Nodes[dependent].Dependencies.Add(address);
                if (++edges > 2_000_000) throw new InvalidOperationException("公式依赖超过 2,000,000 条，未发布本轮结果。");
            }
        }
        if (spills is not null)
        foreach (var spill in spills.Values)
        {
            if (!Nodes.TryGetValue(spill.Anchor, out var producer)) continue;
            for (var row = 0; row < spill.Array.Rows; row++)
            for (var column = 0; column < spill.Array.Columns; column++)
            {
                check();
                if (row == 0 && column == 0) continue;
                foreach (var dependent in DirectDependents(new(spill.Anchor.SheetId, spill.Anchor.Row + row, spill.Anchor.Column + column), check))
                {
                    if (!producer.Dependents.Add(dependent)) continue;
                    Nodes[dependent].Dependencies.Add(spill.Anchor);
                    if (++edges > 2_000_000) throw new InvalidOperationException("数组公式依赖超过 2,000,000 条。");
                }
            }
        }
        CycleMembers = FindCycles(check);
        NonPortableCells = Closure(Nodes.Where(p => p.Value.Script || p.Value.VisibilityBlocked || p.Value.ParseError is not null || p.Value.PortableFormula is null || !p.Value.Cell.Content.DynamicArray && p.Value.Expression is ReferenceNode { IsRange: true }
            || FormulaSyntax.Walk(p.Value.Expression).OfType<FunctionNode>().Any(f => !StandardFunctions.AcceptsArguments(f.Name, f.Arguments.Count))
            || FormulaSyntax.Walk(p.Value.Expression).OfType<ValueNode>().Any(n => n.Value.Kind == CellValueKind.Error && n.Value.ErrorCode is CellErrorCode.Cycle or CellErrorCode.Script))
            .Select(p => p.Key).Concat(CycleMembers), check);
    }

    public IReadOnlySet<CellAddress> AffectedBy(IEnumerable<CellAddress> edits) => AffectedBy(edits, () => { });
    internal HashSet<CellAddress> AffectedBy(IEnumerable<CellAddress> edits, Action check)
    {
        var seeds = new HashSet<CellAddress>();
        foreach (var edit in edits)
        {
            check();
            if (Nodes.ContainsKey(edit)) seeds.Add(edit);
            seeds.UnionWith(DirectDependents(edit, check));
        }
        return Closure(seeds, check);
    }

    internal HashSet<CellAddress> Closure(IEnumerable<CellAddress> seeds, Action check)
    {
        var found = seeds.ToHashSet(); var queue = new Queue<CellAddress>(found);
        while (queue.TryDequeue(out var address))
        {
            check();
            if (!Nodes.TryGetValue(address, out var node)) continue;
            foreach (var dependent in node.Dependents) if (found.Add(dependent)) queue.Enqueue(dependent);
        }
        return found;
    }

    private IEnumerable<CellAddress> DirectDependents(CellAddress address, Action check) =>
        _ranges.TryGetValue(address.SheetId, out var root) && root is not null
            ? root.Query(address.Row, address.Column, check) : [];

    // Iterative Kosaraju avoids overflowing the managed stack on long workbook chains.
    private HashSet<CellAddress> FindCycles(Action check)
    {
        var visited = new HashSet<CellAddress>(); var finish = new List<CellAddress>();
        foreach (var start in Nodes.Keys)
        {
            if (!visited.Add(start)) continue;
            var stack = new Stack<(CellAddress Address, IEnumerator<CellAddress> Edges)>();
            stack.Push((start, Nodes[start].Dependencies.GetEnumerator()));
            while (stack.TryPeek(out var frame))
            {
                check();
                if (frame.Edges.MoveNext())
                {
                    var next = frame.Edges.Current;
                    if (visited.Add(next)) stack.Push((next, Nodes[next].Dependencies.GetEnumerator()));
                }
                else { frame.Edges.Dispose(); stack.Pop(); finish.Add(frame.Address); }
            }
        }
        visited.Clear(); var cycles = new HashSet<CellAddress>();
        for (var i = finish.Count - 1; i >= 0; i--)
        {
            var start = finish[i]; if (!visited.Add(start)) continue;
            var component = new List<CellAddress>(); var stack = new Stack<CellAddress>(); stack.Push(start);
            while (stack.TryPop(out var current))
            {
                check(); component.Add(current);
                foreach (var next in Nodes[current].Dependents) if (visited.Add(next)) stack.Push(next);
            }
            if (component.Count > 1 || Nodes[start].Dependencies.Contains(start)) cycles.UnionWith(component);
        }
        return cycles;
    }

    private sealed record Entry(ReferenceNode Range, CellAddress Dependent);
    private sealed class Interval(int center, Entry[] crossing, Interval? left, Interval? right)
    {
        public static Interval? Build(List<Entry> entries)
        {
            if (entries.Count == 0) return null;
            var center = entries.Select(e => e.Range.Row).Order().ElementAt(entries.Count / 2);
            return new(center, entries.Where(e => e.Range.Row <= center && e.Range.EndRow >= center).ToArray(),
                Build(entries.Where(e => e.Range.EndRow < center).ToList()),
                Build(entries.Where(e => e.Range.Row > center).ToList()));
        }
        public IEnumerable<CellAddress> Query(int row, int column, Action check)
        {
            foreach (var entry in crossing) { check(); if (entry.Range.Contains(row, column)) yield return entry.Dependent; }
            var next = row < center ? left : row > center ? right : null;
            if (next is not null) foreach (var address in next.Query(row, column, check)) yield return address;
        }
    }
}
