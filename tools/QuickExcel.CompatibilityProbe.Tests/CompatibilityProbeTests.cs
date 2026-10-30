using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using QuickExcel.CompatibilityProbe;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.CompatibilityProbe.OpenXmlAudit;

namespace QuickExcel.CompatibilityProbe.Tests;

public sealed partial class CompatibilityProbeTests
{
    [Theory]
    [InlineData("roster")]
    [InlineData("quotation")]
    [InlineData("date-ledger-1900")]
    [InlineData("date-ledger-1904")]
    [InlineData("cross-sheet")]
    [InlineData("print-template")]
    [InlineData("protected-print-template")]
    [InlineData("unknown-function")]
    [InlineData("array-formula")]
    [InlineData("rich-links-notes")]
    [InlineData("phonetic-rich-text")]
    [InlineData("row-column-styles")]
    public void IndependentlyAuthoredInputsPassSdkAndRelationshipValidation(string id)
    {
        var audit = Inspect(Fixture(id).Bytes);
        Assert.True(audit.Issues.Count == 0, JsonSerializer.Serialize(audit.Issues, ProbeRunner.JsonOptions));
        Assert.Empty(Compare(audit, Inspect(Fixture(id).Bytes)));
    }

    [Theory]
    [InlineData("relationship", "relationships")]
    [InlineData("style", "styles")]
    [InlineData("formula", "formulas")]
    [InlineData("metadata", "metadata")]
    public void CorruptionProducesLocatedDifferences(string fault, string category)
    {
        var bytes = Fixture(fault == "metadata" ? "print-template" : "cross-sheet").Bytes;
        var parts = ReadParts(bytes);
        switch (fault)
        {
            case "relationship":
                FixtureCatalog.Edit(parts, "xl/_rels/workbook.xml.rels", root => root.Elements().Single(e => (string?)e.Attribute("Id") == "sheet2").Remove());
                break;
            case "style":
                FixtureCatalog.Edit(parts, "xl/styles.xml", root => root.Element(Main + "cellXfs")!.Elements().First().Remove());
                FixtureCatalog.Edit(parts, "xl/worksheets/sheet1.xml", root => root.Descendants(Main + "c").First().SetAttributeValue("s", 1));
                break;
            case "formula":
                FixtureCatalog.Edit(parts, "xl/worksheets/sheet2.xml", root => root.Descendants(Main + "f").Single().Value = "'数据'!A2*3");
                break;
            case "metadata":
                FixtureCatalog.Edit(parts, "xl/worksheets/sheet1.xml", root => root.Element(Main + "pageSetup")!.Remove());
                break;
        }
        var changed = Inspect(Zip(parts));
        Assert.Contains(Compare(Inspect(bytes), changed), d => d.Category == category && d.Before != d.After && d.Location.Contains("xl/"));
        if (fault is "relationship" or "style") Assert.NotEmpty(changed.Issues);
        if (fault == "formula")
        {
            Assert.Empty(changed.Issues); // A structurally valid wrong formula still needs the semantic diff.
            Assert.Contains(Compare(Inspect(bytes), changed), d => d.Category == "cells" && d.Location.EndsWith("/formula") && d.Before == "'数据'!A2*2" && d.After == "'数据'!A2*3");
        }
    }

    [Fact]
    public void MissingInternalTargetIsDetectedAndExternalLinksAreNotFollowed()
    {
        var parts = ReadParts(Fixture("rich-links-notes").Bytes);
        Assert.Empty(Inspect(Zip(parts)).Issues);
        parts.Remove("xl/comments1.xml");
        Assert.Contains(Inspect(Zip(parts)).Issues, e => e.Message.Contains("Relationship target is missing: xl/comments1.xml"));
    }

    [Fact]
    public void CellFactsKeepLeadingZerosEmptyTextAndDateSerials()
    {
        var roster = Inspect(Fixture("roster").Bytes).Facts;
        Assert.Contains("00123", roster["cells|xl/worksheets/sheet1.xml!A2/value"]);
        Assert.Equal("{\"type\":\"text\",\"value\":\"\"}", roster["cells|xl/worksheets/sheet1.xml!D2/value"]);
        Assert.Equal("{\"type\":\"n\",\"value\":\"60\"}", Inspect(Fixture("date-ledger-1900").Bytes).Facts["cells|xl/worksheets/sheet1.xml!A2/value"]);
    }

    [Fact]
    public void ZipAmbiguityAndXmlEntitiesAreRejected()
    {
        Assert.Throws<InvalidDataException>(() => ReadParts(Zip(new Dictionary<string, byte[]> { ["../outside.xml"] = [] })));
        Assert.Throws<System.Xml.XmlException>(() => Xml(Encoding.UTF8.GetBytes("<!DOCTYPE x [<!ENTITY e SYSTEM 'file:///not-read'>]><x>&e;</x>")));
    }

    [Fact]
    public void EveryEditPathKeepsOriginalAndLabelsLossyCopiesAndPendingExcel()
    {
        var directory = NewEvidenceDirectory();
        var report = ProbeRunner.Run(Fixture("protected-print-template"), directory);
        Assert.True(report.SourceUnchanged);
        Assert.Equal("pending-real-excel", report.ExcelVerification);
        Assert.Null(report.ExcelBuild);
        Assert.Equal(ProbeRunner.Modes, report.Paths.Select(p => p.Mode));
        var unchanged = Assert.Single(report.Paths, p => p.Mode == "unchanged");
        Assert.Equal("preserved-original", unchanged.Status);
        Assert.Equal(report.Source.Sha256, unchanged.Output!.Sha256);
        Assert.Empty(unchanged.Differences);
        var value = Assert.Single(report.Paths, p => p.Mode == "value");
        Assert.Equal("normal-save-blocked;diagnostic-lossy-copy", value.Status);
        Assert.NotEmpty(value.CompatibilityIssues);
        Assert.NotNull(value.Error);
        Assert.Contains("diagnostic-lossy-copy", value.OutputFile!);
        Assert.Contains(value.Differences, d => d.Category == "cells" && d.Location.EndsWith("A2/value") && d.After!.Contains("99"));
        Assert.Throws<IOException>(() => ProbeRunner.Run(Fixture("protected-print-template"), directory));
    }

    [Fact]
    public void SupportedPrintTemplateEditsSaveNormallyAndPreservePageSemantics()
    {
        var directory = NewEvidenceDirectory(); var report = ProbeRunner.Run(Fixture("print-template"), directory);
        Assert.True(report.SourceUnchanged);
        Assert.All(report.Paths, path =>
        {
            Assert.Equal(path.Mode == "unchanged" ? "preserved-original" : "saved-modeled-subset", path.Status);
            Assert.Empty(path.CompatibilityIssues); Assert.NotNull(path.Output); Assert.Empty(path.Output.Issues);
        });
        var storage = new TypedWorkbookStorage();
        var source = storage.Open(Path.Combine(directory, "source.xlsx"));
        foreach (var mode in new[] { "value", "style", "rename" })
        {
            var path = report.Paths.Single(p => p.Mode == mode);
            var output = storage.Open(Path.Combine(directory, path.OutputFile!));
            Assert.True(source.Sheets[0].Printing.ContentEquals(output.Sheets[0].Printing));
        }
    }

    [Fact]
    public void CrossSheetFormulaTransformsAreAuditedWithoutTurningCachesIntoResults()
    {
        var report = ProbeRunner.Run(Fixture("cross-sheet"), NewEvidenceDirectory());
        string Formula(string mode) => Assert.Single(report.Paths, p => p.Mode == mode).Output!.Facts["cells|xl/worksheets/sheet2.xml!A2/formula"];
        Assert.Equal("'数据'!A2*2", Formula("value"));
        Assert.Contains("A3", Formula("insert-row"));
        Assert.Contains("改名后", Formula("rename"));
        Assert.Contains("#REF!", Formula("delete-row"));
    }

    [Theory]
    [InlineData("save")]
    [InlineData("snapshot-export")]
    [InlineData("portable-export")]
    public void WrittenFontComponentsPassIndependentSchemaValidation(string route)
    {
        var directory = NewEvidenceDirectory();
        Directory.CreateDirectory(directory);
        var style = new StyleDefinition { FontName = "Calibri", FontSize = 12, Bold = true, Italic = true, Strike = true, Underline = "double", FontColor = "FF123456" };
        var model = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new WorksheetSnapshot(SheetId.New(), "字体", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("测试"))))])], styles: [style]);
        var storage = new TypedWorkbookStorage();
        var path = Path.Combine(directory, route == "save" ? "font.xlsq" : "font.xlsx");
        if (route == "save") storage.Save(path, model);
        else if (route == "snapshot-export") storage.Export(path, model);
        else storage.ExportPortable(path, model);
        var bytes = File.ReadAllBytes(path);
        if (route == "save") bytes = ReadParts(bytes)["workbook.xlsx"];
        var audit = Inspect(bytes);
        Assert.True(audit.Issues.Count == 0, JsonSerializer.Serialize(audit.Issues, ProbeRunner.JsonOptions));
        Assert.Equal(style, storage.Open(path).Styles[0]);
    }

    private static CompatibilityFixture Fixture(string id) => FixtureCatalog.Create().Single(f => f.Id == id);

    private static string NewEvidenceDirectory()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root.FullName, "QuickExcel.slnx")))
            root = root.Parent ?? throw new InvalidOperationException("Tests must run with --artifacts-path under the QuickExcel repository.");
        return Path.Combine(root.FullName, "artifacts", "compatibility", "tests", Guid.NewGuid().ToString("N"));
    }
}
