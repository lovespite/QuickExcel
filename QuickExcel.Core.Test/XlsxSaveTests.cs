using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class XlsxSaveTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-xlsxsave-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace App = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace Vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    public XlsxSaveTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private string PathFor(string name) => Path.Combine(_folder, name);
    private static WorkbookSnapshot Model() => new(Guid.NewGuid(), 0, ExcelDateSystem.Mac1904,
        [new(SheetId.New(), "原表", cells: [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(42)))),
            new(new(0, 1), new(CellContent.FromFormula("=_xlfn.NOTYET(A1)"), new(CellValue.FromNumber(84), CalculationState.Stale, 0, 0, 0)))],
            showGridLines: false, defaultRowHeight: 30, tabColor: "FF123456")]);
    private static byte[] Part(string path, string name)
    { using var zip = ZipFile.OpenRead(path); using var stream = zip.GetEntry(name)!.Open(); using var memory = new MemoryStream(); stream.CopyTo(memory); return memory.ToArray(); }
    private static void Edit(string path, string name, Func<XDocument?, XDocument> update)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update); var entry = zip.GetEntry(name); XDocument? xml = null;
        if (entry is not null) { using (var stream = entry.Open()) xml = XDocument.Load(stream); entry.Delete(); }
        using var target = zip.CreateEntry(name).Open(); update(xml).Save(target);
    }
    private string Source(bool metadata = true)
    {
        var path = PathFor("source.xlsx"); _storage.Save(path, Model());
        if (!metadata) return path;
        var properties = new[]
        {
            ("core", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "application/vnd.openxmlformats-package.core-properties+xml", """
                <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>报价 &amp; 预算</dc:title><dc:creator>作者</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-10-30T01:00:00Z</dcterms:created></cp:coreProperties>
                """),
            ("app", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "application/vnd.openxmlformats-officedocument.extended-properties+xml", """
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><Company>测试公司</Company><DocSecurity>0</DocSecurity><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>原表</vt:lpstr></vt:vector></TitlesOfParts></Properties>
                """),
            ("custom", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", "application/vnd.openxmlformats-officedocument.custom-properties+xml", """
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="报价号"><vt:lpwstr>00123</vt:lpwstr></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="已批准"><vt:bool>true</vt:bool></property></Properties>
                """)
        };
        foreach (var (name, relationship, type, xml) in properties)
        {
            Edit(path, "docProps/" + name + ".xml", _ => XDocument.Parse(xml));
            Edit(path, "_rels/.rels", doc => { doc!.Root!.Add(new XElement(doc.Root.Name.Namespace + "Relationship", new XAttribute("Id", name), new XAttribute("Type", relationship), new XAttribute("Target", "docProps/" + name + ".xml"))); return doc; });
            Edit(path, "[Content_Types].xml", doc => { doc!.Root!.Add(new XElement(doc.Root.Name.Namespace + "Override", new XAttribute("PartName", "/docProps/" + name + ".xml"), new XAttribute("ContentType", type))); return doc; });
        }
        return path;
    }
    private static WorkbookSnapshot Edited(WorkbookSnapshot snapshot)
    {
        using var session = new DocumentSession(snapshot);
        session.Execute(new SetCellsCommand([new(new(snapshot.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(100)))]));
        return session.Snapshot;
    }

    [Fact]
    public void InPlaceSaveRetainsUnknownFormulaMetadataAndModeledConfiguration()
    {
        var path = Source(); var original = _storage.Open(path); Assert.Empty(original.FileState!.CompatibilityIssues);
        Assert.Equal("", original.ScriptSource); Assert.False(TypedWorkbookStorage.RequiresScriptContainer(original));
        var properties = new[] { "core", "app", "custom" }.ToDictionary(n => n, n => Part(path, "docProps/" + n + ".xml"));
        var edited = Edited(original);
        var plan = TypedWorkbookStorage.InspectSavePlan(edited);
        Assert.Equal(3, plan.Count(p => p.Action == "preserve"));
        var version = _storage.SaveChecked(path, edited, false, FileVersions.Read(path)); Assert.Equal(FileVersions.Read(path), version);
        var saved = _storage.Open(path); Assert.Empty(saved.FileState!.CompatibilityIssues);
        Assert.Equal(100, saved.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal("=_xlfn.NOTYET(A1)", saved.Sheets[0].GetCell(0, 1)!.Content.FormulaText);
        Assert.Equal(CalculationState.Stale, saved.Sheets[0].GetCell(0, 1)!.Calculation!.State);
        Assert.Equal(ExcelDateSystem.Mac1904, saved.DateSystem); Assert.False(saved.Sheets[0].ShowGridLines);
        Assert.Equal(30, saved.Sheets[0].DefaultRowHeight); Assert.Equal("FF123456", saved.Sheets[0].TabColor);
        foreach (var (name, bytes) in properties) Assert.Equal(bytes, Part(path, "docProps/" + name + ".xml"));
        Assert.Equal("1", (string?)XDocument.Parse(Encoding.UTF8.GetString(Part(path, "xl/workbook.xml"))).Root!.Element(Main + "calcPr")!.Attribute("fullCalcOnLoad"));
        var noEdit = PathFor("unchanged.xlsx"); _storage.Save(noEdit, saved); Assert.Equal(File.ReadAllBytes(path), File.ReadAllBytes(noEdit));
    }

    [Fact]
    public void TitleIndexTracksRenameAddDeleteAndOrderWhileOtherPropertiesRemainExact()
    {
        var path = Source(); var original = _storage.Open(path);
        var first = original.Sheets[0];
        var sheets = new[] { new WorksheetSnapshot(SheetId.New(), "新增"), new WorksheetSnapshot(first.Id, "改名", first.Cells) };
        var edited = original.WithSheets(sheets);
        Assert.Contains(TypedWorkbookStorage.InspectSavePlan(edited), p => p.Action == "update-sheet-titles");
        var output = PathFor("renamed.xlsx"); _storage.Save(output, edited);
        var app = XDocument.Parse(Encoding.UTF8.GetString(Part(output, "docProps/app.xml")));
        Assert.Equal(["新增", "改名"], app.Root!.Element(App + "TitlesOfParts")!.Descendants(Vt + "lpstr").Select(e => e.Value));
        Assert.Equal("2", app.Descendants(Vt + "i4").Single().Value);
        Assert.Equal(Part(path, "docProps/core.xml"), Part(output, "docProps/core.xml"));
        _storage.Save(output, edited.WithSheets([sheets[1]]));
        app = XDocument.Parse(Encoding.UTF8.GetString(Part(output, "docProps/app.xml")));
        Assert.Equal("1", app.Descendants(Vt + "i4").Single().Value);
        Assert.Equal("改名", app.Root!.Element(App + "TitlesOfParts")!.Descendants(Vt + "lpstr").Single().Value);
    }

    [Theory]
    [InlineData("custom", "linked")]
    [InlineData("custom", "unknown-type")]
    [InlineData("core", "unknown")]
    [InlineData("app", "unknown-group")]
    [InlineData("app", "unknown-label")]
    [InlineData("app", "unknown")]
    public void UnsafePropertiesBlockEditsAndOnlyExplicitCopyDropsThatPart(string part, string change)
    {
        var path = Source();
        Edit(path, "docProps/" + part + ".xml", doc =>
        {
            var root = doc!.Root!;
            if (change == "linked") root.Elements().First().SetAttributeValue("linkTarget", "原表!A1");
            else if (change == "unknown-type") root.Descendants(Vt + "lpwstr").Single().Name = Vt + "blob";
            else if (change == "unknown-group") root.Element(App + "TitlesOfParts")!.Descendants(Vt + "lpstr").Single().Value = "dangling-name";
            else if (change == "unknown-label") root.Element(App + "HeadingPairs")!.Descendants(Vt + "lpstr").Single().Value = "Named Ranges";
            else root.Add(new XElement("unknown", "data"));
            return doc;
        });
        var bytes = File.ReadAllBytes(path); var book = _storage.Open(path); var changed = Edited(book);
        Assert.Contains(book.FileState!.CompatibilityDiagnostics, d => d.Feature == "metadata/" + part);
        Assert.Throws<WorkbookCompatibilityException>(() => _storage.Save(path, changed));
        Assert.Throws<InvalidOperationException>(() => _storage.Save(path, changed, true));
        Assert.Equal(bytes, File.ReadAllBytes(path));
        var unchanged = PathFor("preserve.xlsx"); _storage.Save(unchanged, book); Assert.Equal(bytes, File.ReadAllBytes(unchanged));
        var copy = PathFor("copy.xlsx"); _storage.Save(copy, changed, true);
        using var zip = ZipFile.OpenRead(copy); Assert.Null(zip.GetEntry("docProps/" + part + ".xml"));
        Assert.Empty(_storage.Open(copy).FileState!.CompatibilityIssues);
    }

    [Theory]
    [InlineData("modified")]
    [InlineData("deleted")]
    [InlineData("created")]
    public void ExpectedVersionGuardsXlsxAndNeverClearsDirtyOnConflict(string change)
    {
        var path = Source(false); var book = _storage.Open(path); var version = FileVersions.Read(path);
        if (change == "deleted") File.Delete(path);
        else if (change == "created") version = new FileVersion(null);
        else File.WriteAllText(path, "external write");
        var bytes = File.Exists(path) ? File.ReadAllBytes(path) : null;
        using var doc = new WorkbookDocument(book, new(_storage, new UnusedCalculator())); doc.SetCell(0, 0, 0, "99");
        Assert.Throws<WorkbookFileConflictException>(() => doc.SaveChecked(path, version)); Assert.True(doc.IsDirty);
        Assert.Equal(bytes, File.Exists(path) ? File.ReadAllBytes(path) : null);
    }

    private string LossySource(string extension)
    {
        var xlsx = Source(false);
        Edit(xlsx, "xl/worksheets/sheet1.xml", doc => { doc!.Root!.Add(new XElement(Main + "customSheetViews")); return doc; });
        if (extension == ".xlsx") return xlsx;
        var container = PathFor("source.xlsq"); _storage.Save(container, _storage.Open(xlsx)); return container;
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void ExplicitCheckedOverwriteAcceptsLossesAndSavesCurrentEdits(string extension)
    {
        var path = LossySource(extension); var original = File.ReadAllBytes(path); var version = FileVersions.Read(path);
        using var doc = new WorkbookDocument(_storage.Open(path), new(_storage, new UnusedCalculator()));
        doc.SetCell(0, 0, 0, "99");
        Assert.NotEmpty(TypedWorkbookStorage.SaveIssues(doc.Snapshot));
        Assert.Throws<InvalidOperationException>(() => doc.SaveChecked(path, version, discardUnsupported: true));
        Assert.Throws<ArgumentException>(() => doc.SaveChecked(path, version, overwriteOriginal: true));
        Assert.True(doc.IsDirty); Assert.Equal(original, File.ReadAllBytes(path));
        var saved = doc.SaveChecked(path, version, discardUnsupported: true, overwriteOriginal: true);
        Assert.Equal(FileVersions.Read(path), saved); Assert.NotEqual(version, saved); Assert.False(doc.IsDirty);
        var reopened = _storage.Open(path);
        Assert.Equal(99, reopened.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal("=_xlfn.NOTYET(A1)", reopened.Sheets[0].GetCell(0, 1)!.Content.FormulaText);
        Assert.Empty(reopened.FileState!.CompatibilityIssues);
    }

    [Theory]
    [InlineData(".xlsx", false)]
    [InlineData(".xlsx", true)]
    [InlineData(".xlsq", false)]
    [InlineData(".xlsq", true)]
    public void ExplicitOverwriteStillRejectsExternalChangesAndDeletion(string extension, bool deleted)
    {
        var path = LossySource(extension); var version = FileVersions.Read(path);
        using var doc = new WorkbookDocument(_storage.Open(path), new(_storage, new UnusedCalculator()));
        doc.SetCell(0, 0, 0, "99");
        if (deleted) File.Delete(path); else File.WriteAllText(path, "external change");
        Assert.Throws<WorkbookFileConflictException>(() => doc.SaveChecked(path, version, discardUnsupported: true, overwriteOriginal: true));
        Assert.True(doc.IsDirty);
        if (deleted) Assert.False(File.Exists(path)); else Assert.Equal("external change", File.ReadAllText(path));
    }

    [Fact]
    public void ExplicitOverwriteRequiresKnownVersionAndStillValidatesBeforeReplacement()
    {
        var path = LossySource(".xlsx"); var original = File.ReadAllBytes(path); var book = _storage.Open(path);
        Assert.Throws<InvalidOperationException>(() => _storage.OverwriteChecked(path, Edited(book), new(null)));
        var badStyle = new WorkbookSnapshot(book.DocumentId, 0, book.DateSystem, book.Sheets,
            styles: [new() { NumberFormatId = 14, NumberFormatCode = "0.000" }]);
        Assert.Throws<InvalidDataException>(() => _storage.OverwriteChecked(path, badStyle, FileVersions.Read(path)));
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Fact]
    public void ExplicitOverwriteDoesNotBypassV1MigrationCopyRequirement()
    {
        var xlsx = LossySource(".xlsx"); var path = PathFor("legacy.xlsq");
        Fixtures.WorkbookPackageFixture.WriteV1(path, File.ReadAllBytes(xlsx));
        var original = File.ReadAllBytes(path); var book = _storage.Open(path);
        Assert.True(book.FileState!.RequiresMigration);
        Assert.Throws<InvalidOperationException>(() => _storage.OverwriteChecked(path, Edited(book), FileVersions.Read(path)));
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Theory]
    [InlineData("// comment")]
    [InlineData(WorkbookDocument.DefaultScript)]
    [InlineData(" ")]
    public void ScriptsNeedXlsqEvenWhenDisabledAndExplicitDiscardDoesNotBypassThis(string script)
    {
        var path = PathFor("target.xlsx"); File.WriteAllText(path, "original");
        var snapshot = Model().WithScriptSource(script); Assert.True(TypedWorkbookStorage.RequiresScriptContainer(snapshot));
        Assert.Throws<InvalidOperationException>(() => _storage.Save(path, snapshot, true)); Assert.Equal("original", File.ReadAllText(path));
        var container = PathFor("main.xlsq"); _storage.Save(container, snapshot); Assert.Equal(script, _storage.Open(container).ScriptSource);
    }

    [Fact]
    public void EmptyAdvancedModuleAndSignatureStillNeedContainer()
    {
        var original = Model();
        Assert.True(TypedWorkbookStorage.RequiresScriptContainer(original.WithScriptModules(new(
            [new("entry", new("scripts/empty.js", ""))], ["entry"], scriptApiVersion: 2))));
        Assert.True(TypedWorkbookStorage.RequiresScriptContainer(original.WithScriptModules(original.ScriptModules.WithSignature(new(1, "test", "certificate", "signature")))));
        using var blank = WorkbookDocument.CreateBlank(new(_storage, new UnusedCalculator()));
        Assert.False(TypedWorkbookStorage.RequiresScriptContainer(blank.Snapshot));
    }

    [Fact]
    public void CalculationConfigurationOutsideKnownFlagsIsReported()
    {
        var path = Source(false);
        Edit(path, "xl/workbook.xml", doc => { doc!.Root!.Element(Main + "calcPr")!.SetAttributeValue("iterate", "1"); return doc; });
        var book = _storage.Open(path);
        Assert.Contains(book.FileState!.CompatibilityIssues, i => i.Contains("iterate"));
        Assert.Throws<WorkbookCompatibilityException>(() => _storage.Save(path, Edited(book)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteValidationFailureNeverReplacesDestination(bool existing)
    {
        var path = PathFor("validate.xlsx"); if (existing) File.WriteAllText(path, "previous file");
        var book = Model();
        // A custom code serializes as a custom format; a simultaneous nonzero built-in ID cannot round-trip.
        var conflictingStyle = new WorkbookSnapshot(book.DocumentId, 0, book.DateSystem, book.Sheets,
            styles: [new() { NumberFormatId = 14, NumberFormatCode = "0.000" }]);
        Assert.Throws<InvalidDataException>(() => _storage.SaveChecked(path, conflictingStyle, false, FileVersions.Read(path)));
        if (existing) Assert.Equal("previous file", File.ReadAllText(path)); else Assert.False(File.Exists(path));
        Assert.Equal(existing ? 1 : 0, Directory.GetFiles(_folder).Length);
    }

    private sealed class UnusedCalculator : IWorkbookCalculator
    { public CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default) => throw new NotSupportedException(); }
}
