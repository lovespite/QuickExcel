using System.Text;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.CompatibilityProbe.OpenXmlAudit;

namespace QuickExcel.CompatibilityProbe.Tests;

public sealed partial class CompatibilityProbeTests
{
    private static byte[] SharedDiagnosticFixture(int count = 2)
    {
        var parts = ReadParts(Fixture("cross-sheet").Bytes);
        parts["xl/sharedStrings.xml"] = Encoding.UTF8.GetBytes("""
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>共享</t><rPh sb="0" eb="1"><t>gong</t></rPh></si></sst>
            """);
        FixtureCatalog.Edit(parts, "[Content_Types].xml", root => root.Add(new XElement(XName.Get("Override", "http://schemas.openxmlformats.org/package/2006/content-types"), new XAttribute("PartName", "/xl/sharedStrings.xml"),
            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"))));
        FixtureCatalog.Edit(parts, "xl/_rels/workbook.xml.rels", root => root.Add(new XElement(PackageRel + "Relationship", new XAttribute("Id", "shared"),
            new XAttribute("Type", Rel.NamespaceName + "/sharedStrings"), new XAttribute("Target", "sharedStrings.xml"))));
        FixtureCatalog.Edit(parts, "xl/worksheets/sheet1.xml", root =>
        {
            root.AddFirst(new XElement(Main + "cols", new XElement(Main + "col", new XAttribute("min", 3), new XAttribute("max", 4), new XAttribute("style", 0))));
            root.Element(Main + "sheetData")!.ReplaceNodes(Enumerable.Range(1, count).Select(i => new XElement(Main + "row", new XAttribute("r", i),
                new XAttribute("s", 0), new XAttribute("customFormat", 1), new XElement(Main + "c", new XAttribute("r", "A" + i), new XAttribute("t", "s"), new XElement(Main + "v", 0)))));
        });
        FixtureCatalog.Edit(parts, "xl/styles.xml", root => root.Element(Main + "fonts")!.Elements().First().Add(new XElement(Main + "outline")));
        return Zip(parts);
    }

    [Fact]
    public void SharedTextAndFontWarningsExpandToAffectedCellsAndRowColumnRangesWithoutWeakeningSaveGuard()
    {
        var folder = NewEvidenceDirectory(); Directory.CreateDirectory(folder); var source = Path.Combine(folder, "shared.xlsx");
        var bytes = SharedDiagnosticFixture(); File.WriteAllBytes(source, bytes); var storage = new TypedWorkbookStorage(); var model = storage.Open(source);
        var diagnostics = model.FileState!.CompatibilityDiagnostics;
        Assert.Contains(diagnostics, d => d.Feature == "text/phonetic" && d.Range == "sharedStrings[0]");
        foreach (var cell in new[] { "A1", "A2" })
            Assert.Contains(diagnostics, d => d.Code == "xlsx.affected-location" && d.Feature == "text/phonetic · sharedStrings[0]" && d.Range == cell && d.SheetName == "数据");
        foreach (var range in new[] { "A1", "A2", "row:1", "row:2", "columns:3:4" })
            Assert.Contains(diagnostics, d => d.Code == "xlsx.affected-location" && d.Feature.Contains("font/outline") && d.Range == range && d.SheetName == "数据");
        Assert.Contains(diagnostics, d => d.Feature.Contains("font/outline") && d.SheetName == "汇总" && d.Range == "A2");
        using var session = new DocumentSession(model); session.Execute(new RenameSheetCommand(model.Sheets[0].Id, "新名称"));
        Assert.Contains(TypedWorkbookStorage.InspectCompatibility(session.Snapshot).Diagnostics, d => d.Code == "xlsx.affected-location" && d.Kind == CompatibilityKind.PotentialLoss && d.SheetName == "数据");
        Assert.Throws<WorkbookCompatibilityException>(() => storage.Save(Path.Combine(folder, "blocked.xlsx"), session.Snapshot));
        Assert.True(session.Undo()); var saved = Path.Combine(folder, "preserved.xlsq"); storage.Save(saved, session.Snapshot);
        Assert.Equal(diagnostics.ToArray(), storage.Open(saved).FileState!.CompatibilityDiagnostics.ToArray()); Assert.Equal(bytes, File.ReadAllBytes(source));
    }
    [Fact]
    public void SharedFeatureExpansionRemainsBoundedAndExplicit()
    {
        var folder = NewEvidenceDirectory(); Directory.CreateDirectory(folder); var source = Path.Combine(folder, "shared.xlsx");
        File.WriteAllBytes(source, SharedDiagnosticFixture(1030)); var model = new TypedWorkbookStorage().Open(source);
        Assert.Equal(1025, model.FileState!.CompatibilityDiagnostics.Length);
        Assert.Contains(model.FileState.CompatibilityDiagnostics, d => d.Code == "report.locations-truncated");
        Assert.True(model.FileState.CompatibilityIssues.Length < 10);
    }
    [Fact]
    public void NamedBaseStyleWarningsFollowCellXfAssociations()
    {
        var parts = ReadParts(SharedDiagnosticFixture());
        FixtureCatalog.Edit(parts, "xl/styles.xml", root => root.Element(Main + "cellStyleXfs")!.Elements().First().SetAttributeValue("quotePrefix", 1));
        var folder = NewEvidenceDirectory(); Directory.CreateDirectory(folder); var path = Path.Combine(folder, "named.xlsx"); File.WriteAllBytes(path, Zip(parts));
        var model = new TypedWorkbookStorage().Open(path);
        Assert.Contains(model.FileState!.CompatibilityDiagnostics, d => d.Range == "A1" && d.SheetName == "数据" && d.Feature.Contains("cellStyleXfs[0]") && d.Feature.Contains("cellXfs[0]"));
    }
    [Theory]
    [InlineData("style")]
    [InlineData("shared-string")]
    public void InvalidCellFieldRejectionCarriesPartSheetAndAddress(string fault)
    {
        var parts = ReadParts(SharedDiagnosticFixture());
        FixtureCatalog.Edit(parts, "xl/worksheets/sheet1.xml", root =>
        {
            var cell = root.Descendants(Main + "c").Last();
            if (fault == "style") cell.SetAttributeValue("s", 999); else cell.Element(Main + "v")!.Value = "999";
        });
        var folder = NewEvidenceDirectory(); Directory.CreateDirectory(folder); var path = Path.Combine(folder, "invalid.xlsx"); File.WriteAllBytes(path, Zip(parts));
        var exception = Assert.Throws<InvalidDataException>(() => new TypedWorkbookStorage().Open(path));
        var diagnostic = Assert.IsType<CompatibilityDiagnostic>(TypedWorkbookStorage.DescribeRejection(exception));
        Assert.Equal(CompatibilityKind.Rejected, diagnostic.Kind); Assert.Equal("数据", diagnostic.SheetName); Assert.Equal("A2", diagnostic.Range);
        Assert.Equal("xl/worksheets/sheet1.xml", diagnostic.Part);
    }
    [Fact]
    public void WorkbookLevelRejectionDoesNotBorrowTheLastVisitedCellLocation()
    {
        var parts = ReadParts(Fixture("cross-sheet").Bytes);
        FixtureCatalog.Edit(parts, "xl/workbook.xml", root => root.Element(Main + "sheets")!.Elements().Last().SetAttributeValue("name", "数据"));
        var folder = NewEvidenceDirectory(); Directory.CreateDirectory(folder); var path = Path.Combine(folder, "invalid.xlsx"); File.WriteAllBytes(path, Zip(parts));
        var exception = Assert.Throws<InvalidDataException>(() => new TypedWorkbookStorage().Open(path));
        var diagnostic = Assert.IsType<CompatibilityDiagnostic>(TypedWorkbookStorage.DescribeRejection(exception));
        Assert.Equal("xl/workbook.xml", diagnostic.Part); Assert.Null(diagnostic.SheetName); Assert.Null(diagnostic.Range);
    }
}
