using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunMergedAxisInsertionChecksAsync(Action<bool, string> check)
    {
        async Task Layout() => await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        async Task Prepare(CellRectangle merge)
        {
            _workspace.New(); ShowBook(); Synchronize();
            var id = Document.Snapshot.ActiveSheetId;
            await ExecuteAsync(new SetCellsCommand([
                new(new(id, merge.Top, merge.Left), CellInput.Parse("合并标题")), new(new(id, 8, 8), CellInput.Parse("相邻数据"))]));
            await ExecuteAsync(new MergeCellsCommand(id, merge));
            await Layout();
        }
        foreach (var axis in new[] { SheetAxis.Row, SheetAxis.Column })
        foreach (var merge in new[] { new CellRectangle(1, 1, 3, 3), new(1, 1, 1, 3), new(1, 1, 3, 1) })
        {
            await Prepare(merge);
            _grid.Select(new(merge.Bottom, merge.Right));
            var row = axis == SheetAxis.Row;
            for (var count = 1; count <= 2; count++)
            {
                // Exercise both public UI entry points, including a second insertion without reselection.
                if (row)
                {
                    var button = _ribbon.GetVisualDescendants().OfType<Button>().Single(b => AutomationProperties.GetName(b) == UiText.Get("插入行"));
                    button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                }
                else
                {
                    _grid.RaiseEvent(new ContextRequestedEventArgs { RoutedEvent = InputElement.ContextRequestedEvent });
                    var insert = _grid.ContextMenu!.Items.OfType<MenuItem>().Single(m => Equals(m.Header, UiText.Get("插入")));
                    var item = insert.Items.OfType<MenuItem>().Single(m => Equals(m.Header, UiText.Get("插入列")));
                    _grid.ContextMenu.Close(); item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                }
                await _operation; await Layout();
                var expected = row ? merge with { Bottom = merge.Bottom + count } : merge with { Right = merge.Right + count };
                check(Document.Snapshot.Sheets[0].MergedCells.Single() == expected.ToString(), $"{axis} insertion {count} expands {merge} by exactly one");
                check(Document.GetCellInput(0, merge.Top, merge.Left) == "合并标题" && Document.GetCellInput(0, row ? 8 + count : 8, row ? 8 : 8 + count) == "相邻数据",
                    "Merged insertion retains anchor text and shifts surrounding cells by one");
                check(Rectangle(_grid.SelectionRanges.Single()) == expected && _grid.Current == new CellPosition(merge.Top, merge.Left),
                    "Expanded merged selection is ready for another insertion");
            }
            await History(false); await History(false);
            check(Document.Snapshot.Sheets[0].MergedCells.Single() == merge.ToString() && Document.GetCellInput(0, 8, 8) == "相邻数据",
                "Two undos restore two insertions including merge geometry and surrounding data");
            check(Rectangle(_grid.SelectionRanges.Single()) == merge, "Undo shrinks the selection with its merged cell");
            await History(true);
            check(Document.Snapshot.Sheets[0].MergedCells.Single() == (row ? merge with { Bottom = merge.Bottom + 1 } : merge with { Right = merge.Right + 1 }).ToString(),
                "Redo restores one merged insertion");
            await ShiftAxisAsync(axis, false);
            check(Document.Snapshot.Sheets[0].MergedCells.Single() == (row ? merge with { Bottom = merge.Bottom + 2 } : merge with { Right = merge.Right + 2 }).ToString(),
                "Insertion after undo/redo still adds only one row/column");
        }

        foreach (var axis in new[] { SheetAxis.Row, SheetAxis.Column })
        {
            var row = axis == SheetAxis.Row; var merge = new CellRectangle(1, 1, 3, 3);
            await Prepare(merge); _grid.Select(new(6, 6));
            var canvas = _grid.GetVisualDescendants().OfType<Canvas>().Single();
            var cell = _grid.GetCellBounds(row ? new(2, 0) : new(0, 2))!.Value;
            var point = row ? new Point(20, cell.Center.Y) : new Point(cell.Center.X, 18);
            using var pointer = new Pointer(row ? 221 : 222, PointerType.Mouse, true);
            canvas.RaiseEvent(new PointerPressedEventArgs(canvas, pointer, this, canvas.TranslatePoint(point, this)!.Value, 0,
                new PointerPointProperties(RawInputModifiers.RightMouseButton, PointerUpdateKind.RightButtonPressed), KeyModifiers.None, 1));
            var selected = _grid.SelectionRanges.Single();
            check(row ? selected.Top == 2 && selected.Rows == 1 : selected.Left == 2 && selected.Columns == 1,
                "Right-click inside a merged row/column retains one explicit header boundary");
            await ShiftAxisAsync(axis, false);
            check(Document.Snapshot.Sheets[0].MergedCells.Single() == (row ? "B2:D5" : "B2:E4") &&
                Document.GetCellInput(0, row ? 9 : 8, row ? 8 : 9) == "相邻数据", "Header insertion adds one row/column inside the merge");
            await History(false);

            // Explicit multi-axis selections still determine the insertion count and location.
            _grid.SelectRange(row ? new(2, 0, 3, WorkbookLimits.MaxColumns - 1) : new(0, 2, WorkbookLimits.MaxRows - 1, 3), new(2, 2));
            await ShiftAxisAsync(axis, false);
            check(Document.Snapshot.Sheets[0].MergedCells.Single() == (row ? "B2:D6" : "B2:F4") &&
                Document.GetCellInput(0, row ? 10 : 8, row ? 8 : 10) == "相邻数据", "Explicit two-axis selection inserts exactly two at its selected boundary");
            await History(false);
            _grid.SelectRange(row ? new(1, 0, 1, WorkbookLimits.MaxColumns - 1) : new(0, 1, WorkbookLimits.MaxRows - 1, 1), new(1, 1));
            await ShiftAxisAsync(axis, false);
            check(Document.Snapshot.Sheets[0].MergedCells.Single() == (row ? "B3:D5" : "C2:E4"), "Header insertion at the merge's outer edge shifts the whole merge");
        }
    }
}
