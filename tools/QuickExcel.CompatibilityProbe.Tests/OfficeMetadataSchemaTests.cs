using System.Text;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.CompatibilityProbe.OpenXmlAudit;

namespace QuickExcel.CompatibilityProbe.Tests;

public sealed class OfficeMetadataSchemaTests
{
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace Ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    private static readonly XNamespace Revision = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
    private static readonly XNamespace Revision2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
    private const string SheetUid = "{00000000-0001-0000-0000-000000000000}";
    private const string EmptyUid = "{00000000-0000-0000-0000-000000000000}";
    private const string ViewUid = "{00000000-000D-0000-FFFF-FFFF00000000}";

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void EditedOfficeMetadataAndDrawingRemainSchemaValid(string extension)
    {
        var bytes = Fixture();
        AssertValid(bytes);
        var directory = NewDirectory();
        var source = Path.Combine(directory, "source.xlsx");
        File.WriteAllBytes(source, bytes);
        var storage = new TypedWorkbookStorage();
        var book = storage.Open(source);
        Assert.Empty(book.FileState!.CompatibilityIssues);
        using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("edited")))]));
        var output = Path.Combine(directory, "edited" + extension);
        storage.Save(output, session.Snapshot);
        var saved = File.ReadAllBytes(output);
        if (extension == ".xlsq") saved = ReadParts(saved)["workbook.xlsx"];
        AssertValid(saved);
        var parts = ReadParts(saved);
        var workbook = Xml(parts["xl/workbook.xml"]).Root!;
        var sheet = Xml(parts["xl/worksheets/sheet1.xml"]).Root!;
        var styles = Xml(parts["xl/styles.xml"]).Root!;
        Assert.Equal("xl", (string?)workbook.Element(Main + "fileVersion")?.Attribute("appName"));
        Assert.Equal("ThisWorkbook", (string?)workbook.Element(Main + "workbookPr")?.Attribute("codeName"));
        Assert.NotNull(workbook.Element(Mc + "AlternateContent"));
        Assert.NotNull(workbook.Element(Revision + "revisionPtr"));
        Assert.NotNull(workbook.Element(Main + "bookViews")?.Element(Main + "workbookView")?.Attribute(Revision2 + "uid"));
        Assert.Equal(SheetUid, (string?)sheet.Attribute(Revision + "uid"));
        Assert.Equal("0.3", (string?)sheet.Element(Main + "sheetFormatPr")?.Attribute(Ac + "dyDescent"));
        Assert.Equal("0.3", (string?)sheet.Descendants(Main + "row").Single().Attribute(Ac + "dyDescent"));
        Assert.True(sheet.Elements().ToList().IndexOf(sheet.Element(Main + "ignoredErrors")!) <
            sheet.Elements().ToList().IndexOf(sheet.Element(Main + "drawing")!));
        var phonetic = sheet.Element(Main + "phoneticPr")!;
        Assert.Equal("MS Gothic", (string?)styles.Element(Main + "fonts")!.Elements().ElementAt((int)phonetic.Attribute("fontId")!).Element(Main + "name")?.Attribute("val"));
        Assert.Empty(storage.Open(output).FileState!.CompatibilityIssues);
        Assert.Equal(bytes, File.ReadAllBytes(source));
    }

    [Fact]
    public void CopiedAndRenamedSheetsKeepMetadataWithoutAliasingSourceIdentity()
    {
        var directory = NewDirectory();
        var source = Path.Combine(directory, "source.xlsx");
        File.WriteAllBytes(source, Fixture());
        var storage = new TypedWorkbookStorage();
        var book = storage.Open(source);
        using var session = new DocumentSession(book);
        session.Execute(new RenameSheetCommand(book.Sheets[0].Id, "Renamed"));
        session.Execute(new CopySheetCommand(book.Sheets[0].Id, "Copy", 0));
        var output = Path.Combine(directory, "copy.xlsx");
        storage.Save(output, session.Snapshot);
        var bytes = File.ReadAllBytes(output);
        AssertValid(bytes);
        var parts = ReadParts(bytes);
        var copy = Xml(parts["xl/worksheets/sheet1.xml"]).Root!;
        var original = Xml(parts["xl/worksheets/sheet2.xml"]).Root!;
        Assert.NotEqual(SheetUid, (string?)copy.Attribute(Revision + "uid"));
        Assert.Equal(SheetUid, (string?)original.Attribute(Revision + "uid"));
        Assert.Equal((string?)original.Element(Main + "phoneticPr")?.Attribute("fontId"),
            (string?)copy.Element(Main + "phoneticPr")?.Attribute("fontId"));
        Assert.NotNull(copy.Element(Main + "ignoredErrors"));
        var descentPrefix = copy.GetPrefixOfNamespace(Ac);
        Assert.NotNull(descentPrefix);
        Assert.Contains(descentPrefix, ((string?)copy.Attribute(Mc + "Ignorable"))!.Split(' '));
        Assert.Empty(storage.Open(output).FileState!.CompatibilityIssues);
    }

    [Theory]
    [InlineData("revision-required", "workbook/revisionPtr")]
    [InlineData("alternate-extra", "workbook/AlternateContent")]
    [InlineData("alternate-prefix", "workbook/AlternateContent")]
    public void UnknownOrMalformedMetadataKeepsTheLossGuard(string change, string feature)
    {
        var parts = ReadParts(Fixture());
        FixtureCatalog.Edit(parts, "xl/workbook.xml", workbook =>
        {
            if (change == "revision-required") workbook.Element(Revision + "revisionPtr")!.Attribute("documentId")!.Remove();
            else if (change == "alternate-extra") workbook.Element(Mc + "AlternateContent")!.Element(Mc + "Choice")!
                .Add(new XElement(XName.Get("feature", "urn:unknown-payload")));
            else workbook.SetAttributeValue(XNamespace.Xmlns + "modern", "urn:unsupported-version");
        });
        var directory = NewDirectory();
        var source = Path.Combine(directory, "unknown.xlsx");
        File.WriteAllBytes(source, Zip(parts));
        var storage = new TypedWorkbookStorage();
        var book = storage.Open(source);
        Assert.Contains(book.FileState!.CompatibilityDiagnostics, d => d.Feature == feature);
        using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("edited")))]));
        Assert.Throws<WorkbookCompatibilityException>(() => storage.Save(Path.Combine(directory, "blocked.xlsx"), session.Snapshot));
    }

    private static void AssertValid(byte[] bytes)
    {
        var issues = Inspect(bytes).Issues;
        Assert.True(issues.Count == 0, string.Join(Environment.NewLine, issues.Select(i => i.Part + " " + i.Location + " " + i.Message)));
    }

    private static string NewDirectory()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root.FullName, "QuickExcel.slnx")))
            root = root.Parent ?? throw new InvalidOperationException("Run with --artifacts-path inside QuickExcel.");
        return Directory.CreateDirectory(Path.Combine(root.FullName, "artifacts", "compatibility", "office-metadata", Guid.NewGuid().ToString("N"))).FullName;
    }

    // Authored as OPC/XML rather than through the production writer. All twelve metadata classes
    // coexist with a drawing, and nonstandard prefixes exercise lexical MC namespace bindings.
    private static byte[] Fixture()
    {
        const string revision6 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6";
        const string revision10 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";
        const string modern = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
        const string path = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac";
        var parts = new Dictionary<string, byte[]>();
        void Add(string name, string xml) => parts.Add(name, Encoding.UTF8.GetBytes(xml));
        Add("[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
              <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
              <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
              <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
            </Types>
            """);
        Add("_rels/.rels", $"""<Relationships xmlns="{PackageRel}"><Relationship Id="book" Type="{Rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
        Add("xl/_rels/workbook.xml.rels", $"""<Relationships xmlns="{PackageRel}"><Relationship Id="sheet" Type="{Rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="{Rel}/styles" Target="styles.xml"/></Relationships>""");
        Add("xl/workbook.xml", $"""
            <workbook xmlns="{Main}" xmlns:r="{Rel}" xmlns:compat="{Mc}" xmlns:modern="{modern}"
              xmlns:rev="{Revision}" xmlns:rev2="{Revision2}" xmlns:edit6="{revision6}" xmlns:edit10="{revision10}"
              compat:Ignorable="modern rev rev2 edit6 edit10">
              <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="30527"/>
              <workbookPr codeName="ThisWorkbook"/>
              <compat:AlternateContent><compat:Choice Requires="modern"><p:absPath xmlns:p="{path}" url="D:\source\"/></compat:Choice></compat:AlternateContent>
              <rev:revisionPtr revIDLastSave="0" documentId="fixture" edit6:coauthVersionLast="47" edit6:coauthVersionMax="47" edit10:uidLastSave="{EmptyUid}"/>
              <bookViews><workbookView rev2:uid="{ViewUid}"/></bookViews>
              <sheets><sheet name="Data" sheetId="1" r:id="sheet"/></sheets>
            </workbook>
            """);
        Add("xl/styles.xml", $"""
            <styleSheet xmlns="{Main}">
              <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="9"/><name val="MS Gothic"/><charset val="128"/></font></fonts>
              <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
              <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
              <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
              <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
              <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
            </styleSheet>
            """);
        Add("xl/worksheets/sheet1.xml", $"""
            <worksheet xmlns="{Main}" xmlns:r="{Rel}" xmlns:compat="{Mc}" xmlns:metric="{Ac}" xmlns:rev="{Revision}"
              xmlns:x14ac="urn:prefix-conflict" compat:Ignorable="metric rev x14ac" rev:uid="{SheetUid}">
              <sheetFormatPr defaultRowHeight="14.4" metric:dyDescent="0.3"/>
              <sheetData><row r="1" metric:dyDescent="0.3"><c r="A1" t="inlineStr"><is><t>00123</t></is></c></row></sheetData>
              <phoneticPr fontId="1" type="Hiragana" alignment="center"/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
              <ignoredErrors><ignoredError sqref="A1:A2" numberStoredAsText="1"/></ignoredErrors>
              <drawing r:id="drawing"/>
            </worksheet>
            """);
        Add("xl/worksheets/_rels/sheet1.xml.rels", $"""<Relationships xmlns="{PackageRel}"><Relationship Id="drawing" Type="{Rel}/drawing" Target="../drawings/drawing1.xml"/></Relationships>""");
        Add("xl/drawings/drawing1.xml", $"""
            <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{Rel}">
              <xdr:oneCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                <xdr:ext cx="9525" cy="9525"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Pixel"/><xdr:cNvPicPr/></xdr:nvPicPr>
                  <xdr:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
                  <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
                </xdr:pic><xdr:clientData/>
              </xdr:oneCellAnchor>
            </xdr:wsDr>
            """);
        Add("xl/drawings/_rels/drawing1.xml.rels", $"""<Relationships xmlns="{PackageRel}"><Relationship Id="image" Type="{Rel}/image" Target="../media/pixel.png"/></Relationships>""");
        parts.Add("xl/media/pixel.png", Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a5YkAAAAASUVORK5CYII="));
        return Zip(parts);
    }
}
