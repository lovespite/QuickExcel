using System.IO;
using System.Windows.Media;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunRowStyleInheritanceSmokeAsync(string folder, Action<bool, string> check)
    {
        async Task Idle()
        {
            await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(_grid.Control);
        }
        var sheet = new WorksheetSnapshot(SheetId.New(), "RowStyles",
            cells: [new(new(0, 0), new(CellContent.FromLiteral(CellValue.Blank, 2)))],
            rows: [new(0, new(Height: 30, ApplyStyle: true)), new(1, new(Height: 30, StyleId: 2, ApplyStyle: false)),
                new(2, new(Height: 30)), new(3, new(Height: 30, StyleId: 2))],
            columns: Enumerable.Range(0, 3).Select(c => new KeyValuePair<int, ColumnMetadata>(c, new(Width: 18, StyleId: 1))));
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [
            new() { FillColor = "FFF2F2F2" },
            new() { FillColor = "FFE6F1EB", NumberFormatCode = "yyyy-mm-dd" },
            new() { FillColor = "FFDDEBFF", Bold = true }
        ]);
        var storage = new TypedWorkbookStorage();
        var input = Path.Combine(folder, "row-styles.xlsx"); storage.Export(input, snapshot);
        InstallDocument(WorkbookDocument.Open(input, _services), null, "行列样式验收");
        GridSurface.Select(new(8, 5)); await Idle();
        void CheckPixels(string stage)
        {
            check(ProductCellPixel(SheetGrid, 0, 1) == Color.FromRgb(242, 242, 242), stage + ": active row style 0 overrides column fill on an absent cell");
            check(ProductCellPixel(SheetGrid, 1, 1) == Color.FromRgb(230, 241, 235), stage + ": disabled row style yields to column fill");
            check(ProductCellPixel(SheetGrid, 2, 1) == Color.FromRgb(230, 241, 235), stage + ": height-only row keeps column fill");
            check(ProductCellPixel(SheetGrid, 3, 1) == Color.FromRgb(221, 235, 255), stage + ": enabled nonzero row style overrides column fill");
            check(ProductCellPixel(SheetGrid, 0, 0) == Color.FromRgb(221, 235, 255), stage + ": explicit cell style overrides row and column styles");
        }
        CheckPixels("Imported XLSX");
        var beforeClear = _document.Snapshot;
        _editor.ClearCells(0, [(0, 1), (1, 1), (2, 1), (3, 1)]);
        check(ReferenceEquals(beforeClear, _document.Snapshot), "Clearing absent cells through EditorController does not create style overrides or undo history");
        _editor.SetCell(0, 0, 1, "61"); _editor.SetCell(0, 1, 1, "61"); RefreshSheet(); await Idle();
        check(_document.GetCellValue(0, 0, 1) == "61" && _document.GetCellValue(0, 1, 1) == "1900-03-01",
            "New desktop input uses the same row/column style selection as empty-cell rendering");
        _editor.ClearCells(0, [(0, 1), (1, 1)]); RefreshSheet();
        check(_document.GetCellInput(0, 0, 1) == "" && _document.GetCellInput(0, 1, 1) == "" &&
            _document.Snapshot.Sheets[0].GetStyleId(0, 1) == 0 && _document.Snapshot.Sheets[0].GetStyleId(1, 1) == 1,
            "Desktop clear removes both values in one transaction while retaining their effective formats");
        await UndoAsync(false);
        check(_document.GetCellValue(0, 0, 1) == "61" && _document.GetCellValue(0, 1, 1) == "1900-03-01",
            "One desktop undo restores the entire clear transaction");
        await UndoAsync(false); await UndoAsync(false); GridSurface.Select(new(8, 5)); await Idle();
        check(_document.Snapshot.Sheets[0].GetCell(0, 1) is null && _document.Snapshot.Sheets[0].GetCell(1, 1) is null,
            "Undo restores sparse cells and their inherited styles");
        _editor.SetCell(0, 6, 5, "force workbook rewrite");
        var output = Path.Combine(folder, "row-styles-edited.xlsq");
        storage.Save(output, _document.Snapshot);
        InstallDocument(WorkbookDocument.Open(output, _services), output, Path.GetFileName(output));
        GridSurface.Select(new(8, 5)); await Idle();
        CheckPixels("Saved and reopened XLSQ");
        CaptureGridEvidence(Path.Combine(folder, "row-style-inheritance.png"));
    }
}
