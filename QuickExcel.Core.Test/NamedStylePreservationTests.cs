using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class NamedStylePreservationTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "QuickExcel-namedstyles-" + Guid.NewGuid().ToString("N"));
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public NamedStylePreservationTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);

    private string Source()
    {
        var path = Path.Combine(_directory, "source.xlsx");
        new TypedWorkbookStorage().Export(path, new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(SheetId.New(), "Styles", cells: [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(12.3), 1)))])],
            styles: [new(), new() { FontName = "Arial", FontColor = "FFFF0000", NumberFormatCode = "0.000", TextRotation = 45, Indent = 2, ShrinkToFit = true }]));
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = zip.GetEntry("xl/styles.xml")!;
        XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
        var root = xml.Root!;
        var parents = root.Element(Ns + "cellStyleXfs")!;
        parents.Add(new XElement(Ns + "xf", new XAttribute("fontId", 1), new XAttribute("numFmtId", 165), new XAttribute("fillId", 0), new XAttribute("borderId", 0)));
        parents.SetAttributeValue("count", 2);
        root.Add(new XElement(Ns + "cellStyles", new XAttribute("count", 2),
            new XElement(Ns + "cellStyle", new XAttribute("name", "Normal"), new XAttribute("xfId", 0), new XAttribute("builtinId", 0)),
            new XElement(Ns + "cellStyle", new XAttribute("name", "报价金额"), new XAttribute("xfId", 1))));
        var child = root.Element(Ns + "cellXfs")!.Elements().ElementAt(1);
        child.SetAttributeValue("xfId", 1); child.SetAttributeValue("applyFont", 0);
        entry.Delete(); using (var stream = zip.CreateEntry("xl/styles.xml").Open()) xml.Save(stream);
        return path;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ValueAndSingleCellStyleEditsKeepOriginalNamedTablesAndApplyFlags(bool changeStyle)
    {
        var source = Source(); var storage = new TypedWorkbookStorage(); var book = storage.Open(source);
        Assert.Empty(book.FileState!.CompatibilityIssues);
        Assert.Equal(1, book.Styles[1].NamedStyleId); Assert.Equal("FFFF0000", book.Styles[1].FontColor);
        Assert.Equal(45, book.Styles[1].TextRotation); Assert.Equal(2, book.Styles[1].Indent); Assert.True(book.Styles[1].ShrinkToFit);
        using var session = new DocumentSession(book);
        if (changeStyle) session.Execute(new SetCellStyleCommand([new(book.Sheets[0].Id, 0, 0)], new(Bold: true)));
        else session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(42), 1))]));
        var output = Path.Combine(_directory, "edited.xlsq"); storage.Save(output, session.Snapshot);
        Assert.Equal(session.Snapshot.Styles.ToArray(), storage.Open(output).Styles.ToArray());
        using var sourceZip = ZipFile.OpenRead(source);
        XDocument original; using (var stream = sourceZip.GetEntry("xl/styles.xml")!.Open()) original = XDocument.Load(stream);
        using var outer = ZipFile.OpenRead(output); using var embedded = outer.GetEntry("workbook.xlsx")!.Open();
        using var memory = new MemoryStream(); embedded.CopyTo(memory); memory.Position = 0;
        using var zip = new ZipArchive(memory, ZipArchiveMode.Read); using var stylesStream = zip.GetEntry("xl/styles.xml")!.Open(); var saved = XDocument.Load(stylesStream);
        foreach (var name in new[] { "cellStyleXfs", "cellStyles" }) Assert.True(XNode.DeepEquals(original.Root!.Element(Ns + name), saved.Root!.Element(Ns + name)));
        Assert.True(XNode.DeepEquals(original.Root!.Element(Ns + "cellXfs")!.Elements().ElementAt(1), saved.Root!.Element(Ns + "cellXfs")!.Elements().ElementAt(1)));
        Assert.True(session.Undo()); Assert.Equal(book.Styles.ToArray(), session.Snapshot.Styles.ToArray());
        Assert.True(session.Redo()); Assert.Equal(storage.Open(output).Styles.ToArray(), session.Snapshot.Styles.ToArray());
    }

    [Fact]
    public void DanglingNamedParentIsRejected()
    {
        var path = Source();
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/styles.xml")!; XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
            xml.Descendants(Ns + "cellStyle").Last().SetAttributeValue("xfId", 20);
            entry.Delete(); using var target = zip.CreateEntry("xl/styles.xml").Open(); xml.Save(target);
        }
        Assert.Throws<InvalidDataException>(() => new TypedWorkbookStorage().Open(path));
    }

    [Fact]
    public void UnknownStyleExtensionCannotBeCopiedAcrossAnEditedNamedWorkbook()
    {
        var path = Source();
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/styles.xml")!; XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
            xml.Root!.Add(new XElement(Ns + "extLst", new XElement(Ns + "ext", new XAttribute("uri", "unknown"),
                new XElement(XName.Get("data", "urn:unknown"), new XAttribute("sheetRange", "Styles!A1")))));
            entry.Delete(); using var target = zip.CreateEntry("xl/styles.xml").Open(); xml.Save(target);
        }
        var storage = new TypedWorkbookStorage(); var book = storage.Open(path);
        Assert.Contains(book.FileState!.CompatibilityIssues, s => s.Contains("extLst"));
        using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(9)))]));
        var output = Path.Combine(_directory, "keep.xlsq"); File.WriteAllText(output, "existing output");
        Assert.Throws<WorkbookCompatibilityException>(() => storage.Save(output, session.Snapshot));
        Assert.Throws<InvalidOperationException>(() => storage.Save(output, session.Snapshot, true));
        Assert.Equal("existing output", File.ReadAllText(output));
    }
}
