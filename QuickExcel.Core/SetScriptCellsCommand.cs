using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Shared by live model writes and legacy patches, including final transaction replay.</summary>
internal sealed class SetScriptCellsCommand(IEnumerable<CellEdit> edits) : DocumentCommand
{
    private readonly SetCellsCommand _command = new(edits);

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        foreach (var group in _command.Edits.GroupBy(e => e.Address.SheetId))
        {
            var imageCells = snapshot.GetSheet(group.Key).Images.Where(i => i.PlaceInCell)
                .Select(i => new CellPosition(i.RowIndex, i.ColumnIndex)).ToHashSet();
            if (group.Any(e => imageCells.Contains(new(e.Address.Row, e.Address.Column))))
                throw new InvalidOperationException("脚本不能改写含单元格内图片的目标单元格；整批操作未应用。");
        }
        // Content edits preserve merges and attachments. The ordinary command permits merge
        // anchors and rejects nonblank writes to covered cells, rich-text loss and spill edits.
        return _command.Prepare(snapshot);
    }
}
