using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Explicit removal of run formatting. Text and base style remain unchanged, as one undoable edit.</summary>
public sealed class FlattenRichTextCommand(IEnumerable<CellAddress> addresses) : DocumentCommand
{
    private readonly CellAddress[] _addresses = addresses.Take(WorkbookLimits.MaxCells + 1).Distinct().ToArray();
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        if (_addresses.Length > WorkbookLimits.MaxCells) throw new ArgumentException("编辑数量超过容量。");
        return new SetCellsCommand(_addresses.Select(a => (Address: a, Cell: snapshot.GetSheet(a.SheetId).GetCell(a.Row, a.Column)))
            .Where(p => p.Cell?.Content.RichText is not null)
            .Select(p => new CellEdit(p.Address, CellContent.FromLiteral(p.Cell!.Value, p.Cell.Content.StyleId))), allowRichTextLoss: true).Prepare(snapshot);
    }
}
