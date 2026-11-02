using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunAutoFitSmokeAsync(string folder, Action<bool, string> check)
    {
        var path = Path.Combine(folder, "auto-fit.xlsq");
        InstallDocument(WorkbookDocument.CreateBlank(_services), path, "自动尺寸回归.xlsq");
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(_grid.Control);
        }
        double Width(int column) => _document.Snapshot.Sheets[0].Columns.GetValueOrDefault(column)?.Width is { } units ? units * 7 + 5 : column == 0 ? 160 : 112;
        double Height(int row) => _document.Snapshot.Sheets[0].Rows.GetValueOrDefault(row)?.Height ?? 31;
        async Task Fit(SheetAxis axis, int index) { AutoFitGridAxis(axis, index); await Idle(); }
        var longText = "自动调整列宽：包含视口外的完整文字 ABC 1234567890";
        ApplyValues([(0, 0, "短"), (80, 0, longText), (0, 1, "=123456789+1"), (1, 1, "0.125"),
            (2, 1, "46000"), (3, 1, "加粗大字号"), (4, 1, "换行文字需要根据现有列宽自动计算完整行高。"),
            (5, 1, "第一行\n第二行\n第三行"), (6, 1, "styled first\nstyled second")]);
        var id = _document.Snapshot.Sheets[0].Id;
        void Style(int row, CellStylePatch patch) => ExecuteWorkbookCommand(new SetCellStyleCommand([new(id, row, 1)], patch));
        Style(1, new(NumberFormatId: 10)); Style(2, new(NumberFormatId: 14));
        Style(3, new(FontSize: 28, Bold: true, Italic: true));
        Style(4, new(WrapText: true, FontSize: 16)); Style(6, new(WrapText: true));
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Column, 1, 1, 12));
        await RecalculateAsync(); await Idle();
        GridSurface.SelectRange(new(1, 2, 2, 3), new(1, 2)); await Idle();
        var selection = GridSurface.Selection.ToArray();
        var revision = _document.Snapshot.Revision;
        var calculations = 0;
        void Changed(object? sender, DocumentChange change) { if (change.CalculationOnly) calculations++; }
        _document.Session.Changed += Changed;
        try
        {
            check(_grid.Control.GetModelCellBounds(80, 0) is null, "Auto-fit fixture includes text outside the viewport");
            await Fit(SheetAxis.Column, 0);
            var fittedWidth = Width(0);
            check(fittedWidth > 160 && _grid.Control.GetTextWidth(longText, false, false) / _grid.Control.RenderScaleX + 4 <= fittedWidth,
                "Column auto-fit includes the complete offscreen display text and padding");
            check(_document.Snapshot.Revision == revision + 1 && selection.SequenceEqual(GridSurface.Selection) && _grid.Current == new CellPosition(1, 2),
                "Auto-fit commits one dimension transaction without changing selection");
            await Fit(SheetAxis.Column, 0);
            check(_document.Snapshot.Revision == revision + 1, "Repeating auto-fit at the same size creates no history");
            await UndoAsync(false); await Idle(); check(Width(0) == 160, "One undo restores the width before auto-fit");
            await UndoAsync(true); await Idle(); check(Width(0) == fittedWidth, "Redo restores auto-fit width");
            await Fit(SheetAxis.Row, 3); check(Height(3) > 31, "Row auto-fit respects a larger bold italic font");
            await Fit(SheetAxis.Row, 4); var narrowHeight = Height(4);
            check(narrowHeight > 60 && Width(1) == 89, "Wrapped row auto-fit expands height without changing column width");
            ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Column, 1, 1, 35));
            await Fit(SheetAxis.Row, 4);
            check(Height(4) < narrowHeight, "Wrapped row height shrinks when the existing column is wider");
            await Fit(SheetAxis.Row, 5); await Fit(SheetAxis.Row, 6);
            check(Height(5) > 50 && Height(6) > 31, "Explicit line breaks are measured for default and wrapped styled cells");
            check(calculations == 0 && !_calculationPending && _document.Snapshot.Sheets[0].GetCell(0, 1)!.Calculation is { State: CalculationState.Fresh } cache &&
                cache.InputRevision == _document.Snapshot.Revision, "Auto-fit and undo/redo preserve fresh formula results without recalculation");
        }
        finally { _document.Session.Changed -= Changed; }

        // Equal displayed strings must have equal fitted widths, regardless of raw value or formula source.
        foreach (var row in new[] { 0, 1, 2 })
        {
            var text = _grid.DisplayText(row, 1);
            ApplyValues([(row, 4, "'" + text)]);
            ExecuteWorkbookCommand(new SetCellStyleCommand([new(id, row, 4)], row == 1 ? new(NumberFormatId: 10) : row == 2 ? new(NumberFormatId: 14) : new()));
            // Measure only this row's entries by using separate columns in a fresh snapshot surface.
            var sheet = new WorksheetSnapshot(SheetId.New(), "Display", new Dictionary<CellPosition, CellSnapshot>
            { [new(0, 0)] = _document.Snapshot.Sheets[0].GetCell(row, 1)!, [new(0, 1)] = _document.Snapshot.Sheets[0].GetCell(row, 4)! });
            var surface = new FastGridSurface((_, _) => false);
            surface.Bind(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: _document.Snapshot.Styles), 0, 10, 10);
            check(Math.Abs(surface.MeasureAutoFit(SheetAxis.Column, 0) - surface.MeasureAutoFit(SheetAxis.Column, 1)) < 1,
                "Auto-fit measures formatted numbers, dates and formula results instead of raw inputs");
        }
        await RecalculateAsync(); await Idle();
        SetSmokeFreeze(1, 1);
        ApplyFilter(new(0, 0, 80, 0), 0, "absent"); await Idle();
        await Fit(SheetAxis.Column, 0);
        check(Width(0) < 160 && CurrentView.FilterRange is not null && _grid.Control.FrozenRowCount == 1 && _grid.Control.FrozenColumnCount == 1,
            "Auto-fit excludes filtered data and preserves frozen panes and filter state");
        ClearFilter_Click(this, new()); await Idle();
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Row, 80, 1, hidden: true));
        await Fit(SheetAxis.Column, 0); check(Width(0) < 160, "Manually hidden rows do not expand auto-fit columns");
        var before = _document.Snapshot.Revision;
        _editState.StartComposition();
        try { await Fit(SheetAxis.Column, 0); check(_document.Snapshot.Revision == before && _gridResize is null, "IME composition blocks automatic resizing"); }
        finally { _editState.EndComposition(); }
        GridSurface.Select(new(8, 0)); FormulaBar.Text = "提交草稿以后再自动测量这一段文字";
        await Fit(SheetAxis.Column, 0);
        check(_document.GetCellInput(0, 8, 0) == "提交草稿以后再自动测量这一段文字" && Width(0) > 160,
            "Auto-fit commits the original-cell draft before measuring its new content");
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Column, 6, 1, 50));
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Row, 10, 1, 100));
        await Fit(SheetAxis.Column, 6); await Fit(SheetAxis.Row, 10);
        check(Math.Abs(Width(6) - 112) < .01 && Height(10) == 31, "Empty axes reset to default dimensions");
        ApplyValues([(12, 6, new string('W', 300)), (14, 6, "short")]);
        ExecuteWorkbookCommand(new MergeCellsCommand(id, new(12, 6, 13, 7)));
        await Fit(SheetAxis.Column, 6); await Fit(SheetAxis.Row, 12);
        check(Width(6) < 112 && Height(12) == 31, "Merged content does not expand one member row or column");
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Column, 1, 1, hidden: true));
        await Fit(SheetAxis.Row, 4);
        check(Height(4) == 31, "Hidden columns do not expand auto-fit row heights");
        ExecuteWorkbookCommand(new SetAxisMetadataCommand(id, SheetAxis.Column, 1, 1, hidden: false));
        await Fit(SheetAxis.Row, 4);
        ApplyValues([(15, 8, new string('W', 32767))]);
        await Fit(SheetAxis.Column, 8);
        check(Math.Abs(Width(8) - 10000) < .01, "Extremely long text clamps to the dimension safety limit");
        await UndoAsync(false); await UndoAsync(false); await Idle();
        var savedWidth = Width(0); var savedHeight = Height(4);
        check(await SaveAsync(false), "Auto-fitted dimensions save");
        await OpenPathAsync(path); await Idle();
        check(Math.Abs(Width(0) - savedWidth) < .01 && Height(4) == savedHeight, "Saved automatic dimensions survive reopening");
        CaptureGridEvidence(Path.Combine(folder, "auto-fit.png"));
    }
}
