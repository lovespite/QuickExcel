using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunGridResizeSmokeAsync(string folder, Action<bool, string> check)
    {
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(() => UpdateLayout(), DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(_grid.Control);
        }
        double Width(int column) => _grid.Control.GetModelCellBounds(0, column)!.Value.Width;
        double Height(int row) => _grid.Control.GetModelCellBounds(row, 0)!.Value.Height;
        bool Near(double a, double b) => Math.Abs(a - b) < 1;
        async Task Resize(SheetAxis axis, int index, double displacement, bool cancel = false)
        {
            check(BeginGridResize(axis, index), "Header resize enters the shared transaction path");
            PreviewGridResize(displacement); await Idle(); CompleteGridResize(cancel); await Idle();
        }
        var path = Path.Combine(folder, "resized.xlsq");
        InstallDocument(WorkbookDocument.CreateBlank(_services), path, "尺寸回归.xlsq"); await Idle();
        GridSurface.SelectRange(new(2, 2, 3, 3), new(2, 2)); await Idle();
        var selection = GridSurface.Selection.ToArray();
        var revision = _document.Snapshot.Revision;
        var bounds = _grid.Control.GetModelCellBounds(0, 0)!.Value;
        check(FindGridResize(new(bounds.Right, 2)) == (SheetAxis.Column, 0) &&
            FindGridResize(new(2, bounds.Bottom)) == (SheetAxis.Row, 0), "Native header geometry identifies both resize edges");
        check(BeginGridResize(SheetAxis.Column, 0), "Column resize starts");
        for (var i = 1; i <= 8; i++) PreviewGridResize(i * 2);
        check(Near(Width(0), 160), "Pointer bursts queue a preview without repeatedly changing geometry");
        await Idle();
        check(Near(Width(0), 176) && _document.Snapshot.Revision == revision, "Only the latest pointer position previews, without history");
        PreviewGridResize(10); await Idle();
        check(Near(Width(0), 170), "Reversing pointer displacement does not sum overlapping deltas");
        PreviewGridResize(35); CompleteGridResize(false); await Idle();
        check(_document.Snapshot.Revision == revision + 1 && Near(Width(0), 195) &&
            Math.Abs(_document.Snapshot.Sheets[0].Columns[0].Width!.Value - (195d - 5) / 7) < .01,
            "Resize commits one transaction using character-width units");
        check(selection.SequenceEqual(GridSurface.Selection) && _grid.Current == new CellPosition(2, 2),
            "Header resizing preserves selection and active cell");
        check(_document.IsDirty && Title.StartsWith("*"), "Resize marks the workbook dirty");
        await UndoAsync(false); await Idle(); check(Near(Width(0), 160) && !_document.Session.CanUndo, "One undo restores column width");
        await UndoAsync(true); await Idle(); check(Near(Width(0), 195), "Redo restores column width");
        await Resize(SheetAxis.Row, 0, 31);
        check(Near(Height(0), 62) && Near(Height(1), 31), "Row resizing does not leak height to neighbors");
        GridSurface.Select(new(80, 0)); await Idle(); GridSurface.Select(new(0, 0)); await Idle();
        check(Near(Height(0), 62), "Resized row survives distant scrolling");
        await UndoAsync(false); await Idle(); check(Near(Height(0), 31), "Row height has one undo");
        await UndoAsync(true); await Idle(); check(Near(Height(0), 62), "Row height has one redo");
        revision = _document.Snapshot.Revision;
        await Resize(SheetAxis.Row, 0, 20, true); await Resize(SheetAxis.Column, 0, 20, true);
        check(_document.Snapshot.Revision == revision && Near(Height(0), 62) && Near(Width(0), 195),
            "Cancel restores both axes without history");
        BeginGridResize(SheetAxis.Row, 0); PreviewGridResize(20); CompleteGridResize(true); await Idle();
        check(_document.Snapshot.Revision == revision && Near(Height(0), 62), "Cancel aborts an unrendered preview");
        BeginGridResize(SheetAxis.Column, 0); CompleteGridResize(false);
        check(_document.Snapshot.Revision == revision, "No movement creates no history");
        _editState.StartComposition();
        try { check(!BeginGridResize(SheetAxis.Column, 0) && _gridResize is null, "IME composition blocks resizing"); }
        finally { _editState.EndComposition(); }

        GridSurface.Select(new(1, 1)); Freeze_Click(new MenuItem(), new()); await Idle();
        await Resize(SheetAxis.Row, 0, 7); await Resize(SheetAxis.Column, 0, 21);
        check(Near(Height(0), 69) && Near(Width(0), 216) && _grid.Control.FrozenRowCount == 1 &&
            _grid.Control.FrozenColumnCount == 1, "Resizing preserves native frozen panes");
        check(await SaveAsync(false), "Resized workbook saves");
        await OpenPathAsync(path); await Idle();
        check(Near(Height(0), 69) && Near(Width(0), 216), "Save/reopen preserves both dimensions");
        GridSurface.Select(new(2, 2)); FormulaBar.Text = "resize draft";
        await Resize(SheetAxis.Row, 0, 10);
        check(_document.GetCellInput(0, 2, 2) == "resize draft", "Resize commits the pending draft at its original address");
        await UndoAsync(false); await Idle();
        check(Near(Height(0), 69) && _document.GetCellInput(0, 2, 2) == "resize draft", "Undo size preserves separately committed input");
        await UndoAsync(false); check(_document.GetCellInput(0, 2, 2) == "", "Previous undo restores input independently");
        ApplyValues([(0, 0, "Header"), (1, 0, "keep"), (2, 0, "hide")]);
        ApplyFilter(new(0, 0, 2, 0), 0, "keep"); await Idle();
        await Resize(SheetAxis.Row, 1, 10);
        check(CurrentView.FilterRange is not null && Near(Height(1), 41) && _grid.Control.GetModelCellBounds(2, 0) is null,
            "Resize preserves filtered row visibility");
        ClearFilter_Click(this, new()); await Idle();
        await Resize(SheetAxis.Column, 1, -1000);
        check(Near(Width(1), 12) && _document.Snapshot.Sheets[0].Columns[1].Width == 1, "Narrow column clamps and round-trips");
        await Resize(SheetAxis.Column, 1, 20000); check(Near(Width(1), 10000), "Column width clamps to the rendering safety limit");
        await Resize(SheetAxis.Row, 0, 20000); check(Near(Height(0), 10000), "Row height clamps to the rendering safety limit");
        await Resize(SheetAxis.Row, 0, -20000); check(Near(Height(0), 1), "Row height clamps to minimum");
        InstallDocument(WorkbookDocument.CreateBlank(_services), null, "尺寸重绘回归.xlsq");
        ApplyValues([(0, 0, "=1+2")]); await RecalculateAsync(); await Idle();
        var model = _grid.Control.Model; var editor = _grid.Editor; var calculations = 0;
        void Changed(object? sender, DocumentChange change) { if (change.CalculationOnly) calculations++; }
        _document.Session.Changed += Changed;
        try
        {
            await Resize(SheetAxis.Column, 0, 16); await Resize(SheetAxis.Row, 0, 10);
            check(calculations == 0 && !_calculationPending && ReferenceEquals(model, _grid.Control.Model) && ReferenceEquals(editor, _grid.Editor),
                "Pure resize keeps model and editor attached without recalculation");
            var cache = _document.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation!;
            check(cache.State == CalculationState.Fresh && cache.InputRevision == _document.Snapshot.Revision,
                "Resize retains fresh formula caches at the current revision");
        }
        finally { _document.Session.Changed -= Changed; }
    }
}
