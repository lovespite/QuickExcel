using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Core.Test.Fixtures;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class OfficeMetadataPreservationTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-office-metadata-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace Ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    private static readonly XNamespace Xr = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private static readonly XNamespace Xr2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
    public OfficeMetadataPreservationTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);

    private string Source()
    {
        var path = Path.Combine(_folder, "source.xlsx");
        File.WriteAllBytes(path, WorkbookPackageFixture.Xlsx("", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                </Types>
                """,
            ["xl/_rels/workbook.xml.rels"] = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """,
            ["xl/workbook.xml"] = """
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                    xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
                    xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" mc:Ignorable="x15 xr xr2 xr6 xr10">
                  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="30527"/>
                  <workbookPr codeName="ThisWorkbook"/>
                  <mc:AlternateContent><mc:Choice Requires="x15"><absPath xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" url="D:\Downloads\"/></mc:Choice></mc:AlternateContent>
                  <xr:revisionPtr revIDLastSave="0" documentId="original-document" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>
                  <bookViews><workbookView xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/></bookViews>
                  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """,
            ["xl/styles.xml"] = """
                <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="9"/><name val="宋体"/><family val="3"/><charset val="134"/></font></fonts>
                  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
                  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
                  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
                  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
                  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
                </styleSheet>
                """,
            ["xl/worksheets/sheet1.xml"] = """
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:d="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                    xmlns:x14ac="urn:other-prefix-owner" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" mc:Ignorable="d xr" xr:uid="{00000000-0001-0000-0000-000000000000}">
                  <sheetFormatPr defaultRowHeight="14.4" d:dyDescent="0.3"/>
                  <sheetData><row r="1" d:dyDescent="0.3"><c r="A1" t="inlineStr"><is><t>123</t></is></c></row>
                    <row r="2" d:dyDescent="0.4"><c r="A2"><v>42</v></c></row></sheetData>
                  <phoneticPr fontId="1" type="Hiragana" alignment="left"/>
                  <ignoredErrors><ignoredError sqref="A1 A2:B3" numberStoredAsText="1" formula="true"/></ignoredErrors>
                </worksheet>
                """
        }));
        return path;
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void AllTwelveKindsSurviveOrdinaryEditedSaveAndReopen(string extension)
    {
        var source = Source(); var original = File.ReadAllBytes(source); var book = _storage.Open(source);
        Assert.Empty(book.FileState!.CompatibilityIssues);
        using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("456")))]));
        var path = Path.Combine(_folder, "edited" + extension); _storage.Save(path, session.Snapshot);
        var reopened = _storage.Open(path);
        Assert.Empty(reopened.FileState!.CompatibilityIssues);
        Assert.Equal("456", reopened.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        Assert.True(book.Sheets[0].Data.ContentEquals(reopened.Sheets[0].Data));
        Assert.Equal(.3, reopened.Sheets[0].Rows[0].Descent); Assert.Equal(.4, reopened.Sheets[0].Rows[1].Descent);
        var workbook = Part(path, "xl/workbook.xml"); var sheet = Part(path, "xl/worksheets/sheet1.xml");
        Assert.Equal("30527", (string?)workbook.Element(Main + "fileVersion")!.Attribute("rupBuild"));
        Assert.Equal("ThisWorkbook", (string?)workbook.Element(Main + "workbookPr")!.Attribute("codeName"));
        var choice = workbook.Element(Mc + "AlternateContent")!.Element(Mc + "Choice")!;
        Assert.Equal("http://schemas.microsoft.com/office/spreadsheetml/2010/11/main", choice.GetNamespaceOfPrefix((string)choice.Attribute("Requires")!)!.NamespaceName);
        Assert.Equal("D:\\Downloads\\", (string?)choice.Elements().Single().Attribute("url"));
        Assert.Equal("original-document", (string?)workbook.Element(Xr + "revisionPtr")!.Attribute("documentId"));
        Assert.Equal("{00000000-000D-0000-FFFF-FFFF00000000}", (string?)workbook.Element(Main + "bookViews")!.Element(Main + "workbookView")!.Attribute(Xr2 + "uid"));
        Assert.Equal("{00000000-0001-0000-0000-000000000000}", (string?)sheet.Attribute(Xr + "uid"));
        Assert.Contains(((string)sheet.Attribute(Mc + "Ignorable")!).Split(' '), prefix => sheet.GetNamespaceOfPrefix(prefix) == Ac);
        Assert.Equal(original, File.ReadAllBytes(source));
        Assert.DoesNotContain(TypedWorkbookStorage.InspectCompatibility(reopened).Diagnostics, d => d.Kind is CompatibilityKind.PotentialLoss or CompatibilityKind.CannotCalculate);
    }

    [Fact]
    public void SheetCopyRenameAndReorderKeepSettingsWithoutDuplicatingIdentity()
    {
        var book = _storage.Open(Source()); var id = book.Sheets[0].Id;
        using var session = new DocumentSession(book);
        var copy = new CopySheetCommand(id, "Copy", 0); session.Execute(copy);
        session.Execute(new RenameSheetCommand(id, "Renamed"));
        session.Execute(new ShiftAxisCommand(copy.Id, SheetAxis.Row, 0, 1));
        var path = Path.Combine(_folder, "copied.xlsx"); _storage.Save(path, session.Snapshot);
        var reopened = _storage.Open(path);
        Assert.Empty(reopened.FileState!.CompatibilityIssues);
        Assert.Equal("Copy", reopened.Sheets[0].Name); Assert.Equal("Renamed", reopened.Sheets[1].Name);
        Assert.Equal(book.Sheets[0].Data.PhoneticProperties, reopened.Sheets[0].Data.PhoneticProperties);
        Assert.Equal(.3, reopened.Sheets[0].Rows[1].Descent);
        Assert.Null(Part(path, "xl/worksheets/sheet1.xml").Attribute(Xr + "uid"));
        Assert.NotNull(Part(path, "xl/worksheets/sheet2.xml").Attribute(Xr + "uid"));
        Assert.Equal(new CellRectangle(1, 0, 1, 0), reopened.Sheets[0].Data.IgnoredErrors[0].Range);
    }

    [Fact]
    public void PhoneticFontIsSelfContainedWhenOriginalStyleTableIsUnavailable()
    {
        var source = _storage.Open(Source()); var sheet = source.Sheets[0];
        var detached = new WorkbookSnapshot(Guid.NewGuid(), 0, source.DateSystem, [sheet], styles: source.Styles);
        var path = Path.Combine(_folder, "detached.xlsx"); _storage.Save(path, detached);
        var reopened = _storage.Open(path);
        Assert.Equal(sheet.Data.PhoneticProperties, reopened.Sheets[0].Data.PhoneticProperties);
        var properties = Part(path, "xl/worksheets/sheet1.xml").Element(Main + "phoneticPr")!;
        var font = Part(path, "xl/styles.xml").Element(Main + "fonts")!.Elements().ElementAt((int)properties.Attribute("fontId")!);
        Assert.Equal("宋体", (string?)font.Element(Main + "name")!.Attribute("val"));
    }

    [Theory]
    [InlineData("choice")]
    [InlineData("version")]
    [InlineData("revision")]
    [InlineData("ignored")]
    [InlineData("phonetic")]
    [InlineData("descent")]
    public void UnrecognizedMetadataStillBlocksEditedSave(string variant)
    {
        var source = Source();
        Edit(source, variant is "choice" or "version" or "revision" ? "xl/workbook.xml" : "xl/worksheets/sheet1.xml", root =>
        {
            if (variant == "choice") root.Element(Mc + "AlternateContent")!.Add(new XElement(Mc + "Fallback", new XElement(Main + "unknown")));
            if (variant == "version") root.Element(Main + "fileVersion")!.SetAttributeValue("unknown", "keep");
            if (variant == "revision") root.Element(Xr + "revisionPtr")!.Attribute("documentId")!.Remove();
            if (variant == "ignored") root.Element(Main + "ignoredErrors")!.Element(Main + "ignoredError")!.SetAttributeValue("futureFlag", "1");
            if (variant == "phonetic") root.Element(Main + "phoneticPr")!.SetAttributeValue("futureFlag", "1");
            if (variant == "descent") root.Element(Main + "sheetFormatPr")!.SetAttributeValue(Ac + "dyDescent", "NaN");
        });
        var book = _storage.Open(source); Assert.NotEmpty(book.FileState!.CompatibilityIssues);
        using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("changed")))]));
        var output = Path.Combine(_folder, "blocked.xlsx");
        Assert.Throws<WorkbookCompatibilityException>(() => _storage.Save(output, session.Snapshot));
        Assert.False(File.Exists(output));
        _storage.Save(output, session.Snapshot, discardUnsupported: true);
        Assert.Empty(_storage.Open(output).FileState!.CompatibilityIssues);
    }

    [Fact]
    public void UnknownAttributesWithSameLocalNameInDifferentNamespacesRemainSeparateTypes()
    {
        var source = Source();
        Edit(source, "xl/worksheets/sheet1.xml", root =>
        {
            var row = root.Element(Main + "sheetData")!.Elements().First();
            row.Add(new XAttribute(XName.Get("setting", "urn:first"), "1"), new XAttribute(XName.Get("setting", "urn:second"), "1"));
        });
        var report = TypedWorkbookStorage.InspectCompatibility(_storage.Open(source));
        var types = report.Diagnostics.Where(d => d.Code == "xlsx.unmodeled").GroupBy(d => (d.TypeFeature, d.TypeNamespace)).ToArray();
        Assert.Equal(2, types.Length);
        Assert.All(types, type => Assert.Equal(1, CompatibilityDiagnosticCollector.Summarize(type).Rows.Count));
        Assert.Equal(1, CompatibilityDiagnosticCollector.Summarize(report.Diagnostics).Rows.Count);
    }

    [Theory]
    [InlineData("xl/workbook.xml", "workbookPr")]
    [InlineData("xl/worksheets/sheet1.xml", "sheetFormatPr")]
    public void DuplicateSingletonMetadataCannotDisappearOnSave(string part, string element)
    {
        var source = Source();
        Edit(source, part, root => root.Add(new XElement(root.Element(Main + element)!)));
        Assert.Throws<InvalidDataException>(() => _storage.Open(source));
    }

    private static XElement Part(string path, string part)
    {
        using var zip = ZipFile.OpenRead(path);
        if (Path.GetExtension(path) == ".xlsq")
        {
            using var entry = zip.GetEntry("workbook.xlsx")!.Open(); using var bytes = new MemoryStream(); entry.CopyTo(bytes); bytes.Position = 0;
            using var nested = new ZipArchive(bytes, ZipArchiveMode.Read); using var stream = nested.GetEntry(part)!.Open(); return XDocument.Load(stream).Root!;
        }
        using var source = zip.GetEntry(part)!.Open(); return XDocument.Load(source).Root!;
    }
    private static void Edit(string path, string part, Action<XElement> edit)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update); var entry = zip.GetEntry(part)!;
        XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
        edit(xml.Root!); entry.Delete(); using var output = zip.CreateEntry(part).Open(); xml.Save(output);
    }
}
