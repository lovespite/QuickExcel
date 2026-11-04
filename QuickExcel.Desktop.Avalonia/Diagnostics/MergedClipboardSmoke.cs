using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunMergedClipboardChecksAsync(Action<bool, string> check)
    {
        var originalClipboard = _clipboard;
        try
        {
            var memory = new MemoryClipboard(); _clipboard = memory;
            _workspace.New(); ShowBook(); var sheet = Document.Snapshot.ActiveSheetId;
            await ExecuteAsync(new SetCellsCommand([
                new(new(sheet, 0, 0), CellInput.Parse("合并来源")),
                new(new(sheet, 3, 3), CellInput.Parse("目标")), new(new(sheet, 3, 4), CellInput.Parse("保留内容")),
                new(new(sheet, 4, 3), CellInput.Parse("123")), new(new(sheet, 4, 4), CellInput.Parse("=2+3"))]));
            await ExecuteAsync(new MergeCellsCommand(sheet, new(0, 0, 1, 1)));
            _grid.Select(new(0, 0)); await CopyAsync(false);
            check(memory.Value.Token is not null && _clipboardCells?.MergedCells.Length == 1,
                "Copy handler captures the complete selected merge with typed clipboard data");
            _grid.Select(new(4, 6)); await PasteAsync("all");
            check(Document.Snapshot.Sheets[0].MergedCells.Contains("G5:H6") && Document.GetCellInput(0, 4, 6) == "合并来源",
                "Paste handler copies merged geometry and anchor content to an ordinary range");
            await History(false);
            check(!Document.Snapshot.Sheets[0].MergedCells.Contains("G5:H6") && Document.GetCellInput(0, 4, 6) == "",
                "One UI undo removes both pasted cells and their merge");
            await History(true);
            check(Document.Snapshot.Sheets[0].MergedCells.Contains("G5:H6"), "UI redo restores pasted merge geometry");
            _grid.Select(new(0, 0)); await CopyAsync(false);
            _grid.Select(new(3, 3)); await PasteAsync("formats");
            check(Document.Snapshot.Sheets[0].MergedCells.Contains("D4:E5") &&
                Document.GetCellInput(0, 3, 3) == "目标" && Document.GetCellInput(0, 3, 4) == "保留内容" && Document.GetCellInput(0, 4, 4) == "=2+3",
                "Format paste merges a filled target while preserving its covered contents and formulas");
            _grid.Select(new(0, 10)); await CopyAsync(false);
            _grid.Select(new(3, 3)); await PasteAsync("formats");
            check(!Document.Snapshot.Sheets[0].MergedCells.Contains("D4:E5") &&
                Document.GetCellInput(0, 3, 4) == "保留内容" && Document.GetCellInput(0, 4, 4) == "=2+3",
                "Single unmerged cell format paste unmerges the selected target and reveals its retained contents");
            await History(false);
            check(Document.Snapshot.Sheets[0].MergedCells.Contains("D4:E5"), "Format-paste unmerge is one undoable UI operation");
        }
        finally { _clipboard = originalClipboard; }
    }
}
