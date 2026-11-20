using Avalonia.Controls;
using QuickExcel.Core;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task DeleteCellsAsync()
    {
        if (!await CommitEditorsAsync()) return;
        var range = SelectedRange(); var rectangle = Rectangle(range);
        var wholeRow = SelectionPlanning.IsWholeRow(rectangle);
        var wholeColumn = SelectionPlanning.IsWholeColumn(rectangle);
        if (wholeRow != wholeColumn)
        {
            await ShiftAxisAsync(wholeRow ? SheetAxis.Row : SheetAxis.Column, true);
            return;
        }
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId;
        var choice = await ShowOwnedAsync<int?>(new DeleteCellsDialog());
        if (choice is null) return;
        if (session != Document.Session.SessionId || snapshot.ActiveSheetId != Document.Snapshot.ActiveSheetId ||
            snapshot.Revision != Document.Snapshot.Revision || range != SelectedRange())
            throw new InvalidOperationException(UiText.Get("文档或选区已改变，请重新打开删除对话框。"));
        DocumentCommand command = choice.Value switch
        {
            0 => new DeleteCellsCommand(snapshot.ActiveSheetId, rectangle, CellDeleteShift.Left),
            1 => new DeleteCellsCommand(snapshot.ActiveSheetId, rectangle, CellDeleteShift.Up),
            2 => new ShiftAxisCommand(snapshot.ActiveSheetId, SheetAxis.Row, range.Top, range.Rows, true),
            _ => new ShiftAxisCommand(snapshot.ActiveSheetId, SheetAxis.Column, range.Left, range.Columns, true)
        };
        if (await ExecuteAsync(command, snapshot.Revision))
        {
            _grid.Select(new(range.Top, range.Left));
            await CalculatePendingAsync();
            _status.Text = UiText.Get("已删除，可一次撤销。");
        }
    }
}
