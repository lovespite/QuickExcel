using QuickExcel.Domain;
using System.Collections.Immutable;

namespace QuickExcel.Core;

/// <summary>Commit all tabs of the format dialog as one history entry.</summary>
public sealed class SetCellFormatCommand(IEnumerable<CellAddress> addresses, CellStylePatch style, CellBorderPatch? borders) : DocumentCommand
{
    private readonly CellAddress[] _addresses = addresses.Take(20_001).ToArray();
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var first = new SetCellStyleCommand(_addresses, style).Prepare(snapshot);
        if (borders is null) return first;
        var second = new SetCellBordersCommand(_addresses, borders).Prepare(first.Snapshot);
        var changes = first.Changes.Concat(second.Changes).GroupBy(c => c.Address)
            .Select(g => new CellChange(g.Key, g.First().Before, g.Last().After)).ToImmutableArray();
        return second with { Changes = changes, CalculationUnchanged = true };
    }
}
