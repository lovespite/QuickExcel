using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task CheckDeleteCellsAsync(Action<bool, string> check, Func<Control, string, Task> capture)
    {
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static void Click(Window window, string name) => window.GetVisualDescendants().OfType<Button>()
            .Single(c => c.Name == name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        _workspace.New(); ShowBook(); var id = Document.Snapshot.ActiveSheetId;
        await ExecuteAsync(new SetCellsCommand([
            new(new(id, 1, 1), CellInput.Parse("11")), new(new(id, 1, 2), CellInput.Parse("22")),
            new(new(id, 1, 3), CellInput.Parse("33")), new(new(id, 2, 1), CellInput.Parse("44")),
            new(new(id, 3, 1), CellInput.Parse("55")), new(new(id, 0, 5), CellInput.Parse("=D2"))]));
        await CalculateAsync();
        try
        {
            for (var choice = 0; choice < 4; choice++)
            {
                _grid.Select(new(1, 1)); var before = Document.Snapshot;
                _dialogDriver = async window =>
                {
                    await Layout(); check(window is DeleteCellsDialog, "Delete cells opens four-option dialog");
                    var radios = window.GetVisualDescendants().OfType<RadioButton>().ToArray();
                    check(radios.Length == 4 && radios[0].IsChecked == true, "Delete defaults to shifting cells left");
                    radios[choice].IsChecked = true;
                    if (choice == 0) await capture(window, "delete-cells");
                    Click(window, "DeleteCellsConfirm");
                };
                await DeleteCellsAsync();
                check(Document.GetCellInput(0, 1, 1) == (choice is 0 or 3 ? "22" : "44"), "Delete option moves the appropriate neighboring cells: " + choice);
                check(Document.GetCellInput(0, 0, choice == 3 ? 4 : 5) == (choice is 0 or 3 ? "=C2" : choice == 2 ? "=#REF!" : "=D2"), "Delete option updates references: " + choice);
                await History(false);
                check(Document.GetCellInput(0, 1, 1) == "11" && Document.GetCellInput(0, 1, 3) == "33" && Document.GetCellInput(0, 3, 1) == "55", "One undo restores deleted cells: " + choice);
            }
            _grid.Select(new(1, 1)); var revision = Document.Snapshot.Revision;
            _dialogDriver = async window => { await Layout(); Click(window, "DeleteCellsCancel"); };
            await DeleteCellsAsync(); check(Document.Snapshot.Revision == revision, "Cancel delete leaves workbook unchanged");
            _dialogDriver = window =>
            {
                _workspace.Editor.SetCell(0, 8, 0, "newer"); window.Close((int?)0); return Task.CompletedTask;
            };
            try { await DeleteCellsAsync(); throw new Exception("Stale delete unexpectedly succeeded"); }
            catch (InvalidOperationException) { check(Document.GetCellInput(0, 1, 1) == "11", "Stale delete rejects without shifting cells"); }
        }
        finally { _dialogDriver = null; }
    }
}
