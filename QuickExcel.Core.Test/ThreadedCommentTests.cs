using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Core.Test.Fixtures;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class ThreadedCommentTests
{
    private const string Person = "{11111111-1111-4111-8111-111111111111}", Root = "{22222222-2222-4222-8222-222222222222}", Reply = "{33333333-3333-4333-8333-333333333333}";
    private const string ThreadNs = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
    private static byte[] Fixture(string? person = null, string? parent = null, string extra = "", bool externalRelation = false)
    {
        var types = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>";
        foreach (var (path, type) in new[] { ("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"), ("xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"), ("xl/threadedComments/threadedComment1.xml", "application/vnd.ms-excel.threadedcomments+xml"), ("xl/persons/person.xml", "application/vnd.ms-excel.person+xml") })
            types += $"<Override PartName=\"/{path}\" ContentType=\"{type}\"/>";
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = types + "</Types>",
            ["xl/_rels/workbook.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"people\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/person\" Target=\"persons/person.xml\"/></Relationships>",
            ["xl/worksheets/_rels/sheet1.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"threads\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/threadedComment\" Target=\"../threadedComments/threadedComment1.xml\"/></Relationships>",
            ["xl/persons/person.xml"] = $"<personList xmlns=\"{ThreadNs}\"><person id=\"{Person}\" displayName=\"Fixture Author\"/></personList>",
            ["xl/threadedComments/threadedComment1.xml"] = $"<ThreadedComments xmlns=\"{ThreadNs}\"><threadedComment ref=\"A1\" dT=\"2026-11-03T00:00:00Z\" personId=\"{person ?? Person}\" id=\"{Root}\"><text>Fixture root</text>{extra}</threadedComment><threadedComment ref=\"A1\" dT=\"2026-11-03T00:01:00Z\" personId=\"{Person}\" id=\"{Reply}\" parentId=\"{parent ?? Root}\"><text>Fixture reply</text></threadedComment></ThreadedComments>"
        };
        if (externalRelation) parts["xl/threadedComments/_rels/threadedComment1.xml.rels"] = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"unknown\" Type=\"urn:future:attachment\" Target=\"https://invalid.invalid/do-not-fetch\" TargetMode=\"External\"/></Relationships>";
        return WorkbookPackageFixture.Xlsx("<row r=\"1\"><c r=\"A1\"><v>10</v></c></row>", parts);
    }
    private static string Folder()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../e5/threads", Guid.NewGuid().ToString("N"))); Directory.CreateDirectory(path); return path;
    }
    private static byte[] Part(byte[] bytes, string path)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes)); using var source = zip.GetEntry(path)!.Open(); using var result = new MemoryStream(); source.CopyTo(result); return result.ToArray();
    }
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void ValuesCanChangeWhileThreadsAndPeopleRemainByteIdentical(string extension)
    {
        var folder = Folder(); var input = Path.Combine(folder, "source.xlsx"); var bytes = Fixture(); File.WriteAllBytes(input, bytes);
        var storage = new TypedWorkbookStorage(); using var session = new DocumentSession(storage.Open(input)); var sheet = session.Snapshot.Sheets[0];
        Assert.Empty(session.Snapshot.FileState!.CompatibilityIssues); Assert.Single(session.Snapshot.FileState.PreservedAttachments);
        Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 0, 1)));
        Assert.Throws<NotSupportedException>(() => WorkbookClipboard.Capture(session.Snapshot, sheet.Id, new(0, 0, 0, 0)));
        Assert.Throws<NotSupportedException>(() => session.Execute(new SetNoteCommand(sheet.Id, new(0, 0), new("Fixture", new RichCellText([new TextRun("changed")])))));
        Assert.False(session.CanUndo);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(20)))]));
        Assert.True(session.Undo()); Assert.True(session.Redo());
        Assert.Contains(TypedWorkbookStorage.InspectCompatibility(session.Snapshot).Diagnostics,
            d => d.Code == "xlsx.threaded-comments-preserved" && d.Kind == CompatibilityKind.PreservedOnly);
        var output = Path.Combine(folder, "edited." + extension); storage.Save(output, session.Snapshot); var opened = storage.Open(output);
        Assert.Equal(CellValue.FromNumber(20), opened.Sheets[0].GetCell(0, 0)!.Value); Assert.Single(opened.FileState!.PreservedAttachments);
        foreach (var path in new[] { "xl/threadedComments/threadedComment1.xml", "xl/persons/person.xml" }) Assert.Equal(Part(bytes, path), Part(opened.FileState.OriginalWorkbook.ToArray(), path));
        Assert.Equal(bytes, File.ReadAllBytes(input));
        foreach (var portable in new[] { false, true })
        {
            var exported = Path.Combine(folder, portable ? "portable.xlsx" : "values.xlsx");
            if (portable) storage.ExportPortable(exported, session.Snapshot); else storage.Export(exported, session.Snapshot);
            var exportedBook = storage.Open(exported); Assert.Single(exportedBook.FileState!.PreservedAttachments);
            foreach (var part in new[] { "xl/threadedComments/threadedComment1.xml", "xl/persons/person.xml" })
                Assert.Equal(Part(bytes, part), Part(File.ReadAllBytes(exported), part));
        }
    }
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void BrokenPersonOrReplyIdentityIsRejected(bool person)
    {
        var path = Path.Combine(Folder(), "broken.xlsx"); File.WriteAllBytes(path, Fixture(person: person ? Reply : null, parent: person ? null : Reply));
        Assert.Throws<InvalidDataException>(() => new TypedWorkbookStorage().Open(path));
    }
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void UnknownThreadStructureOrExternalDependencyKeepsOriginalAndBlocksSilentLoss(bool external)
    {
        var folder = Folder(); var path = Path.Combine(folder, "source.xlsx"); var bytes = Fixture(extra: external ? "" : "<future/>", externalRelation: external); File.WriteAllBytes(path, bytes);
        var storage = new TypedWorkbookStorage(); using var session = new DocumentSession(storage.Open(path)); Assert.NotEmpty(session.Snapshot.FileState!.CompatibilityIssues);
        Assert.Empty(session.Snapshot.FileState.PreservedAttachments); var copy = Path.Combine(folder, "copy.xlsx"); storage.Save(copy, session.Snapshot); Assert.Equal(bytes, File.ReadAllBytes(copy));
        session.Execute(new SetCellsCommand([new(new(session.Snapshot.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(2)))]));
        Assert.Throws<WorkbookCompatibilityException>(() => storage.Save(Path.Combine(folder, "blocked.xlsx"), session.Snapshot));
    }
}
