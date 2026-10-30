using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class CellTextTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-cell-text-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private string FilePath(string name) => Path.Combine(_folder, name);
    public CellTextTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private static RichCellText Rich() => new([new("中文😀", new("Arial", 18, true, false, "FFFF0000")), new(" blue\n", new("Calibri", 14, false, true, "FF0000FF", Underline: "single")), new("tail")]);
    private static WorkbookSnapshot Book(bool annotations = true)
    {
        var notes = annotations ? new WorksheetAnnotations(
            [new(new(0, 1), new("https://example.invalid/a?q=1&b=2", Tooltip: "说明")), new(new(0, 2), new(null, "'Other'!$A$1", "跳转"))],
            [new(new(2, 0), new("测试作者", new([new("批注\n完整内容", new("Arial", 9, IndexedColor: 81))])))]) : new();
        var sheet = new WorksheetSnapshot(SheetId.New(), "Data", [new(new(0, 0), new(CellContent.FromRichText(Rich()))),
            new(new(0, 1), new(CellContent.FromLiteral(CellValue.FromText("链接"))))], annotations: notes);
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet, new(SheetId.New(), "Other")]);
    }
    private static void EditPart(string file, string name, Action<XDocument> edit)
    {
        using var zip = ZipFile.Open(file, ZipArchiveMode.Update); var entry = zip.GetEntry(name)!; XDocument xml;
        using (var stream = entry.Open()) xml = XDocument.Load(stream);
        edit(xml); entry.Delete(); using var output = zip.CreateEntry(name).Open(); xml.Save(output);
    }
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void RichTextLinksAndNotesSurviveNormalValueStyleEditAndReopen(string extension)
    {
        var path = FilePath("source." + extension); var original = Book(); _storage.Save(path, original); var book = _storage.Open(path);
        Assert.Empty(book.FileState!.CompatibilityIssues); Assert.Equal(Rich(), book.Sheets[0].GetCell(0, 0)!.Content.RichText);
        Assert.True(original.Sheets[0].Annotations.ContentEquals(book.Sheets[0].Annotations));
        Assert.Equal(3, book.Sheets[0].UsedRowCount);
        using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new SetCellsCommand([new(new(id, 4, 4), CellContent.FromLiteral(CellValue.FromNumber(7)))]));
        session.Execute(new SetCellStyleCommand([new(id, 0, 0)], new(FillColor: "FF00FF00")));
        Assert.Equal(Rich(), session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText);
        _storage.Save(path, session.Snapshot); var reopened = _storage.Open(path);
        Assert.Equal(Rich(), reopened.Sheets[0].GetCell(0, 0)!.Content.RichText);
        Assert.True(book.Sheets[0].Annotations.ContentEquals(reopened.Sheets[0].Annotations));
        Assert.Empty(reopened.FileState!.CompatibilityIssues);
        if (extension == "xlsq") { using var zip = ZipFile.OpenRead(path); using var reader = new StreamReader(zip.GetEntry("manifest.json")!.Open()); Assert.Contains("cell-text-annotations", reader.ReadToEnd()); }
        _storage.Export(FilePath("values.xlsx"), session.Snapshot, true, true);
        Assert.Equal(Rich(), _storage.Open(FilePath("values.xlsx")).Sheets[0].GetCell(0, 0)!.Content.RichText);
    }
    [Fact]
    public void EditingRichTextRequiresExplicitFlattenAndPreservesAtomicUndoAndClear()
    {
        var book = Book(false); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetCellsCommand([
            new(new(id, 2, 2), CellContent.FromLiteral(CellValue.FromText("first"))), new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromText("changed")))])));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
        session.Execute(new FlattenRichTextCommand([new(id, 0, 0)])); Assert.True(session.IsDirty);
        Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText); Assert.Equal(Rich().Text, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        session.Undo(); Assert.Equal(Rich(), session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText);
        session.Redo(); Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText);
        session.Undo(); session.Execute(new ClearCellsCommand([new(id, 0, 0)])); Assert.Null(session.Snapshot.Sheets[0].GetCell(0, 0));
        session.Undo(); Assert.Equal(Rich(), session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText);
    }
    [Fact]
    public void WholeCellFontCommandsOverrideEveryRunWhileUndoRestoresMixedFormatting()
    {
        var book = Book(false); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new SetCellStyleCommand([new(id, 0, 0)], new(Bold: true, Italic: false, FontName: "Arial", FontColor: "FF00FF00")));
        Assert.All(session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText!.Runs, r => { Assert.True(r.Font!.Bold); Assert.False(r.Font.Italic); Assert.Equal("Arial", r.Font.Name); Assert.Equal("FF00FF00", r.Font.Color); });
        _storage.Save(FilePath("fonts.xlsx"), session.Snapshot);
        Assert.Equal(session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText, _storage.Open(FilePath("fonts.xlsx")).Sheets[0].GetCell(0, 0)!.Content.RichText);
        session.Undo(); Assert.Equal(Rich(), session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText);
        session.Execute(new SetCellStyleCommand([new(id, 0, 0)], new(Reset: true)));
        Assert.All(session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText!.Runs, r => Assert.Null(r.Font));
    }
    [Fact]
    public void FormattingAnInheritedRunKeepsItsOriginalCellFontSizeAndColor()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Font", [new(new(0, 0), new(CellContent.FromRichText(new([new("继承字体")]))))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new() { FontName = "Arial", FontSize = 14, FontColor = "FFFF0000" }]);
        using var session = new DocumentSession(book); session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new(Bold: true)));
        var font = session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.RichText!.Runs[0].Font!;
        Assert.True(font.Bold); Assert.Equal("Arial", font.Name); Assert.Equal(14, font.Size); Assert.Equal("FFFF0000", font.Color);
        _storage.Save(FilePath("inherited.xlsx"), session.Snapshot);
        Assert.Equal(font, _storage.Open(FilePath("inherited.xlsx")).Sheets[0].GetCell(0, 0)!.Content.RichText!.Runs[0].Font);
    }
    [Fact]
    public void NoteAndLinkCoordinatesCannotBeSilentlyBrokenByStructureSortOrMerge()
    {
        var book = Book(); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 0, 1)));
        Assert.Throws<NotSupportedException>(() => session.Execute(new RenameSheetCommand(book.Sheets[1].Id, "Renamed")));
        Assert.Throws<NotSupportedException>(() => session.Execute(new SortRangeCommand(id, new(0, 0), new(3, 2), 0)));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new MergeCellsCommand(id, new(1, 0, 2, 0))));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
        session.Execute(new ClearCellsCommand([new(id, 0, 1)])); Assert.Equal(2, session.Snapshot.Sheets[0].Annotations.Links.Count);
    }
    [Fact]
    public void SharedAndInlineRichStringsUseTheSameTypedRunsWithoutPhoneticText()
    {
        var path = FilePath("shared.xlsx"); _storage.Save(path, Book(false));
        XDocument? strings = null;
        EditPart(path, "xl/worksheets/sheet1.xml", xml =>
        {
            var cell = xml.Descendants(Main + "c").First(); var inline = cell.Element(Main + "is")!;
            strings = new(new XElement(Main + "sst", new XElement(Main + "si", inline.Elements())));
            inline.Remove(); cell.SetAttributeValue("t", "s"); cell.Add(new XElement(Main + "v", 0));
        });
        EditPart(path, "xl/_rels/workbook.xml.rels", xml => { XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships"; xml.Root!.Add(new XElement(rel + "Relationship", new XAttribute("Id", "strings"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"), new XAttribute("Target", "sharedStrings.xml"))); });
        EditPart(path, "[Content_Types].xml", xml => xml.Root!.Add(new XElement(xml.Root.Name.Namespace + "Override", new XAttribute("PartName", "/xl/sharedStrings.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"))));
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update)) { using var output = zip.CreateEntry("xl/sharedStrings.xml").Open(); strings!.Save(output); }
        Assert.Equal(Rich(), _storage.Open(path).Sheets[0].GetCell(0, 0)!.Content.RichText);
        EditPart(path, "xl/sharedStrings.xml", xml => xml.Descendants(Main + "si").Single().Add(new XElement(Main + "rPh", new XAttribute("sb", 0), new XAttribute("eb", 1), new XElement(Main + "t", "PINYIN"))));
        var withPhonetic = _storage.Open(path); Assert.Equal(Rich(), withPhonetic.Sheets[0].GetCell(0, 0)!.Content.RichText);
        Assert.Contains(withPhonetic.FileState!.CompatibilityDiagnostics, d => d.Feature == "text/phonetic");
    }
    [Fact]
    public void InvalidRunAndHyperlinkDataAreRejectedWithoutChangingSource()
    {
        Assert.Throws<ArgumentException>(() => new RichCellText([new(new string('x', 32767)), new("y")]));
        Assert.Throws<ArgumentException>(() => new RichCellText([new("x", new(Size: double.NaN))]));
        var path = FilePath("bad.xlsx"); _storage.Save(path, Book());
        EditPart(path, "xl/worksheets/sheet1.xml", xml => xml.Descendants(Main + "hyperlink").First().SetAttributeValue(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"), "missing"));
        var bytes = File.ReadAllBytes(path); Assert.Throws<InvalidDataException>(() => _storage.Open(path)); Assert.Equal(bytes, File.ReadAllBytes(path));
    }

    [Fact]
    public void RgbTintIsNotAppliedTwiceWhenRunFormattingIsRewritten()
    {
        var path = FilePath("tint.xlsx"); _storage.Save(path, Book(false));
        EditPart(path, "xl/worksheets/sheet1.xml", xml => xml.Descendants(Main + "color").First().SetAttributeValue("tint", 0.5));
        var book = _storage.Open(path); var rich = book.Sheets[0].GetCell(0, 0)!.Content.RichText!;
        Assert.Equal("FFFF0000", rich.Runs[0].Font!.RawRgb); Assert.NotEqual(rich.Runs[0].Font!.RawRgb, rich.Runs[0].Font!.Color);
        using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 5, 5), CellContent.FromLiteral(CellValue.FromNumber(2)))]));
        _storage.Save(path, session.Snapshot); Assert.Equal(rich, _storage.Open(path).Sheets[0].GetCell(0, 0)!.Content.RichText);
    }

    [Fact]
    public void NoteOnlyVmlIsKeptExactlyAndUnmodeledControlsKeepTheSaveGuard()
    {
        var book = Book(); var sheet = book.Sheets[0];
        var drawing = Encoding.UTF8.GetBytes("""
            <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="note" style="visibility:hidden"><x:ClientData ObjectType="Note"><x:Row>2</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>
            """);
        var annotated = new WorksheetSnapshot(sheet.Id, sheet.Name, sheet.Cells, annotations: new(sheet.Annotations.Links, sheet.Annotations.Notes, drawing));
        book = book.WithSheets([annotated, book.Sheets[1]]); var path = FilePath("vml.xlsx"); _storage.Save(path, book);
        var opened = _storage.Open(path); Assert.Equal(drawing, opened.Sheets[0].Annotations.NoteDrawing);
        using (var session = new DocumentSession(opened))
        {
            session.Execute(new SetCellsCommand([new(new(opened.Sheets[0].Id, 5, 5), CellContent.FromLiteral(CellValue.FromNumber(2)))]));
            _storage.Save(path, session.Snapshot);
        }
        Assert.Equal(drawing, _storage.Open(path).Sheets[0].Annotations.NoteDrawing);
        EditPart(path, "xl/drawings/notes1.vml", xml => xml.Root!.Add(new XElement(XName.Get("rect", "urn:schemas-microsoft-com:vml"))));
        var unsupported = _storage.Open(path); Assert.Contains(unsupported.FileState!.CompatibilityDiagnostics, d => d.Feature == "comments/vml");
        var original = File.ReadAllBytes(path);
        using var blocked = new DocumentSession(unsupported);
        blocked.Execute(new SetCellsCommand([new(new(unsupported.Sheets[0].Id, 5, 5), CellContent.FromLiteral(CellValue.FromNumber(3)))]));
        Assert.Throws<WorkbookCompatibilityException>(() => _storage.Save(FilePath("blocked.xlsx"), blocked.Snapshot));
        Assert.Equal(original, File.ReadAllBytes(path));
    }
}
