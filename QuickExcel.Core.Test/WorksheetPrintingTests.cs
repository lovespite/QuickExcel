using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class WorksheetPrintingTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-printing-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public WorksheetPrintingTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private static WorksheetPrinting Printing() => new()
    {
        Margins = new(.4, .5, .6, .7, .2, .3), Properties = new(false, true),
        Setup = new() { PaperSize = 9, Orientation = "landscape", FitToWidth = 1, FitToHeight = 0, Scale = 75, FirstPageNumber = 2, UseFirstPageNumber = true,
            PaperWidth = "210mm", PaperHeight = "297mm", HorizontalDpi = 600, VerticalDpi = 600, BlackAndWhite = true, Draft = false, Errors = "dash", CellComments = "atEnd", Copies = 2, PageOrder = "overThenDown" },
        Options = new(true, false, true, true, true),
        HeaderFooter = new() { DifferentFirst = true, DifferentOddEven = true, AlignWithMargins = true, ScaleWithDoc = false,
            OddHeader = "&L中文 &C&B报价&R&P/&N", OddFooter = "&D &T &F &A", EvenHeader = "偶数", EvenFooter = "偶尾", FirstHeader = "首页", FirstFooter = "首尾" },
        Areas = [new(PrintRangeKind.Cells, 1, 1, 30, 8), new(PrintRangeKind.Cells, 40, 0, 50, 5)],
        Titles = [new(PrintRangeKind.Rows, 0, 0, 1, PrintRange.MaxColumns - 1), new(PrintRangeKind.Columns, 0, 0, PrintRange.MaxRows - 1, 1)],
        RowBreaks = [new(20, 0, PrintRange.MaxColumns - 1, true)], ColumnBreaks = [new(6, 0, PrintRange.MaxRows - 1, true)],
        PrinterSettings = [1, 2, 3, 4]
    };
    private static WorkbookSnapshot Book() => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [new(SheetId.New(), "报价,'版本", printing: Printing()), new(SheetId.New(), "Other")]);
    private static byte[] Part(string path, string name)
    { using var zip = ZipFile.OpenRead(path); using var stream = zip.GetEntry(name)!.Open(); using var buffer = new MemoryStream(); stream.CopyTo(buffer); return buffer.ToArray(); }
    private static void Mutate(string path, string part, Action<XElement> edit)
    { var xml = XDocument.Parse(Encoding.UTF8.GetString(Part(path, part))); edit(xml.Root!); using var zip = ZipFile.Open(path, ZipArchiveMode.Update); zip.GetEntry(part)!.Delete(); using var stream = zip.CreateEntry(part).Open(); xml.Save(stream); }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void PageSettingsAndNamesSurviveValueStyleRenameAndReopen(string extension)
    {
        var source = Book(); var path = Path.Combine(_folder, "source" + extension); _storage.Save(path, source);
        var opened = _storage.Open(path); Assert.Empty(opened.FileState!.CompatibilityIssues);
        Assert.True(source.Sheets[0].Printing.ContentEquals(opened.Sheets[0].Printing));
        using var session = new DocumentSession(opened); var id = opened.Sheets[0].Id;
        session.Execute(new SetCellsCommand([new(new(id, 3, 2), CellContent.FromLiteral(CellValue.FromNumber(42)))]));
        session.Execute(new SetCellStyleCommand([new(id, 3, 2)], new(Bold: true)));
        session.Execute(new RenameSheetCommand(id, "改名'打印"));
        _storage.Save(path, session.Snapshot); var saved = _storage.Open(path);
        Assert.True(source.Sheets[0].Printing.ContentEquals(saved.Sheets[0].Printing)); Assert.Equal("改名'打印", saved.Sheets[0].Name);
        Assert.Empty(saved.DefinedNames); Assert.Equal(42, saved.Sheets[0].GetCell(3, 2)!.Value.NumberValue);
        if (extension == ".xlsq") Assert.Contains("worksheet-printing", Encoding.UTF8.GetString(Part(path, "manifest.json")));
        else Assert.Equal(new byte[] { 1, 2, 3, 4 }, Part(path, "xl/printerSettings/printerSettings1.bin"));
    }

    [Fact]
    public void PrintingReferencesAndBreaksFollowAxesAndUndoWithoutAllocatingFullRows()
    {
        using var session = new DocumentSession(Book()); var id = session.Snapshot.Sheets[0].Id; var original = session.Snapshot.Sheets[0].Printing;
        session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 0, 1)); var rows = session.Snapshot.Sheets[0].Printing;
        Assert.Equal(2, rows.Areas[0].Top); Assert.Equal(2, rows.Titles[0].Bottom); Assert.Equal(original.Titles[1], rows.Titles[1]);
        Assert.Equal(21, rows.RowBreaks[0].Id); Assert.Equal(original.ColumnBreaks[0], rows.ColumnBreaks[0]);
        session.Execute(new ShiftAxisCommand(id, SheetAxis.Column, 0, 1)); var columns = session.Snapshot.Sheets[0].Printing;
        Assert.Equal(2, columns.Areas[0].Left); Assert.Equal(2, columns.Titles[1].Right); Assert.Equal(7, columns.ColumnBreaks[0].Id);
        Assert.Empty(session.Snapshot.Sheets[0].Cells); session.Undo(); session.Undo();
        Assert.True(original.ContentEquals(session.Snapshot.Sheets[0].Printing)); session.Redo();
        var path = Path.Combine(_folder, "shifted.xlsx"); _storage.Save(path, session.Snapshot); Assert.True(rows.ContentEquals(_storage.Open(path).Sheets[0].Printing));
        var beforeDelete = session.Snapshot;
        Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 0, 3, delete: true)));
        Assert.Same(beforeDelete, session.Snapshot);
        var noBreakBook = Book();
        using var noBreaks = new DocumentSession(noBreakBook.WithSheets(noBreakBook.Sheets.Select(s => s.WithPrinting(s.Printing with { RowBreaks = [], ColumnBreaks = [] }))));
        noBreaks.Execute(new ShiftAxisCommand(noBreaks.Snapshot.Sheets[0].Id, SheetAxis.Row, 0, 2, delete: true));
        Assert.DoesNotContain(noBreaks.Snapshot.Sheets[0].Printing.Titles, t => t.Kind == PrintRangeKind.Rows);
    }

    [Fact]
    public void UnsafeMovesAreAtomicAndDeletingASheetRemovesItsPrintNames()
    {
        using var session = new DocumentSession(Book()); var before = session.Snapshot; var id = before.Sheets[0].Id;
        Assert.Throws<NotSupportedException>(() => session.Execute(new MoveCellsCommand(id, new(1, 1), new(2, 2), id, new(5, 5))));
        Assert.Throws<NotSupportedException>(() => session.Execute(new MoveAxisCommand(id, SheetAxis.Row, 4, 2, 7)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
        session.Execute(new DeleteSheetCommand(id)); var path = Path.Combine(_folder, "delete.xlsx"); _storage.Save(path, session.Snapshot);
        Assert.DoesNotContain("Print_Area", Encoding.UTF8.GetString(Part(path, "xl/workbook.xml")));
        session.Undo(); Assert.True(before.Sheets[0].Printing.ContentEquals(session.Snapshot.Sheets[0].Printing));
    }

    [Fact]
    public void PrinterAndDrawingRelationshipsCoexistAndFullColumnPrintAreasStaySparse()
    {
        var source = Book(); var first = source.Sheets[0];
        var image = new SheetImage(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/mWQAAAAASUVORK5CYII="), "png", 2, 3, 1, 1);
        var printing = first.Printing with { Areas = [new(PrintRangeKind.Columns, 0, 0, PrintRange.MaxRows - 1, 2)] };
        var sheet = new WorksheetSnapshot(first.Id, first.Name, images: [image], printing: printing);
        var path = Path.Combine(_folder, "drawing.xlsx"); _storage.Save(path, source.WithSheets([sheet, source.Sheets[1]]));
        var opened = _storage.Open(path); Assert.True(printing.ContentEquals(opened.Sheets[0].Printing));
        Assert.Single(opened.Sheets[0].Images); Assert.Empty(opened.Sheets[0].Cells); Assert.Empty(opened.FileState!.CompatibilityIssues);
        using var session = new DocumentSession(opened); session.Execute(new ShiftAxisCommand(opened.Sheets[0].Id, SheetAxis.Row, 0, 1));
        Assert.Equal(printing.Areas.ToArray(), session.Snapshot.Sheets[0].Printing.Areas.ToArray());
    }

    [Theory]
    [InlineData("name")]
    [InlineData("attribute")]
    public void UnsupportedPrintingRemainsProtectedUntilExplicitCopy(string kind)
    {
        var path = Path.Combine(_folder, "unknown.xlsx"); _storage.Save(path, Book());
        if (kind == "name") Mutate(path, "xl/workbook.xml", root => root.Element(Main + "definedNames")!.Elements().First().Value = "OFFSET('报价,''版本'!$A$1,0,0,5,5)");
        else Mutate(path, "xl/worksheets/sheet1.xml", root => root.Element(Main + "pageSetup")!.SetAttributeValue("futureOption", "1"));
        var bytes = File.ReadAllBytes(path); var opened = _storage.Open(path); Assert.NotEmpty(opened.FileState!.CompatibilityIssues);
        _storage.Save(Path.Combine(_folder, "unchanged.xlsx"), opened); Assert.Equal(bytes, File.ReadAllBytes(Path.Combine(_folder, "unchanged.xlsx")));
        using var session = new DocumentSession(opened); session.Execute(new SetCellsCommand([new(new(opened.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(2)))]));
        Assert.Throws<WorkbookCompatibilityException>(() => _storage.Save(path, session.Snapshot)); Assert.Equal(bytes, File.ReadAllBytes(path));
    }

    [Fact]
    public void InvalidMarginsCountsAndDuplicateNamesAreRejected()
    {
        foreach (var kind in new[] { "margin", "count", "duplicate" })
        {
            var path = Path.Combine(_folder, kind + ".xlsx"); _storage.Save(path, Book());
            if (kind == "duplicate") Mutate(path, "xl/workbook.xml", root => root.Element(Main + "definedNames")!.Add(new XElement(root.Element(Main + "definedNames")!.Elements().First())));
            else Mutate(path, "xl/worksheets/sheet1.xml", root => root.Element(Main + (kind == "margin" ? "pageMargins" : "rowBreaks"))!.SetAttributeValue(kind == "margin" ? "left" : "count", kind == "margin" ? "NaN" : "9"));
            Assert.Throws<InvalidDataException>(() => _storage.Open(path));
        }
    }
}
