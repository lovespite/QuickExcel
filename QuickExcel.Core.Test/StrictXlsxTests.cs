using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using QuickExcel.Core.Test.Fixtures;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class StrictXlsxTests
{
    private const string Main = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    private const string Rel = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    private static string DirectoryForTest()
    {
        var directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../e5/strict", Guid.NewGuid().ToString("N")));
        Directory.CreateDirectory(directory); return directory;
    }
    private static byte[] Fixture(string rows, bool system1904 = false, string extra = "") => WorkbookPackageFixture.Xlsx("", new Dictionary<string, string>
    {
        ["_rels/.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"book\" Type=\"{Rel}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>",
        ["xl/workbook.xml"] = $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\" conformance=\"strict\"><workbookPr date1904=\"{(system1904 ? 1 : 0)}\"/><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"s\"/></sheets></workbook>",
        ["xl/_rels/workbook.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"s\" Type=\"{Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>",
        ["xl/worksheets/sheet1.xml"] = $"<worksheet xmlns=\"{Main}\"><sheetData>{rows}</sheetData>{extra}</worksheet>"
    });

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void StrictInputEditsAndReopensWithoutChangingVariantOrText(string extension)
    {
        var directory = DirectoryForTest(); var path = Path.Combine(directory, "source.xlsx");
        var bytes = Fixture($"<row r=\"1\"><c r=\"A1\" t=\"d\"><v>2024-04-18T12:00:00.000</v></c><c r=\"B1\" t=\"inlineStr\"><is><t>{Main}</t></is></c><c r=\"C1\"><f>A1+1</f><v>45401.5</v></c></row>");
        File.WriteAllBytes(path, bytes); var storage = new TypedWorkbookStorage();
        using var session = new DocumentSession(storage.Open(path));
        Assert.Empty(session.Snapshot.FileState!.CompatibilityIssues);
        Assert.Equal(CellValue.FromNumber(45400.5), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        session.Execute(new SetCellsCommand([new(new(session.Snapshot.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(45402.25)))]));
        ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts()));
        var output = Path.Combine(directory, "edited." + extension); storage.Save(output, session.Snapshot);
        var reopened = storage.Open(output);
        Assert.Equal(CellValue.FromNumber(45403.25), reopened.Sheets[0].GetCell(0, 2)!.Value);
        Assert.Equal(CellValue.FromText(Main), reopened.Sheets[0].GetCell(0, 1)!.Value);
        Assert.Equal(bytes, File.ReadAllBytes(path));
        using var zip = new ZipArchive(new MemoryStream(reopened.FileState!.OriginalWorkbook.ToArray()));
        using var xml = zip.GetEntry("xl/workbook.xml")!.Open();
        var root = XDocument.Load(xml).Root!;
        Assert.Equal(XName.Get("workbook", Main), root.Name); Assert.Equal("strict", (string?)root.Attribute("conformance"));
    }

    [Theory]
    [InlineData("1900-02-28T00:00:00", false, 59)]
    [InlineData("1900-02-29T12:00:00", false, 60.5)]
    [InlineData("1900-03-01", false, 61)]
    [InlineData("1904-01-01T06:00:00", true, 0.25)]
    public void DateCellsUseWorkbookCalendar(string text, bool system1904, double expected)
    {
        var path = Path.Combine(DirectoryForTest(), "dates.xlsx");
        File.WriteAllBytes(path, Fixture($"<row r=\"1\"><c r=\"A1\" t=\"d\"><v>{text}</v></c></row>", system1904));
        Assert.Equal(CellValue.FromNumber(expected), new TypedWorkbookStorage().Open(path).Sheets[0].GetCell(0, 0)!.Value);
    }

    [Fact]
    public void UnknownStrictObjectsRetainOriginalPackageAndBlockLossyNormalSave()
    {
        var directory = DirectoryForTest(); var source = Path.Combine(directory, "source.xlsx");
        var bytes = Fixture("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>", extra: "<futureObject/>");
        File.WriteAllBytes(source, bytes); var storage = new TypedWorkbookStorage();
        using var session = new DocumentSession(storage.Open(source));
        Assert.NotEmpty(session.Snapshot.FileState!.CompatibilityIssues);
        var copy = Path.Combine(directory, "copy.xlsx"); storage.Save(copy, session.Snapshot);
        Assert.Equal(bytes, File.ReadAllBytes(copy));
        session.Execute(new SetCellsCommand([new(new(session.Snapshot.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(2)))]));
        Assert.Throws<WorkbookCompatibilityException>(() => storage.Save(Path.Combine(directory, "blocked.xlsx"), session.Snapshot));
    }
}
