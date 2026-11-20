using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.CompatibilityProbe.OpenXmlAudit;

namespace QuickExcel.CompatibilityProbe.Tests;

public sealed partial class CompatibilityProbeTests
{
    [Fact]
    public void RichTextAndAttachmentsUseNormalStyleSaveWhileUnsafeEditsAreRejected()
    {
        var report = ProbeRunner.Run(Fixture("rich-links-notes"), NewEvidenceDirectory());
        var saved = Assert.Single(report.Paths, p => p.Mode == "style");
        Assert.Equal("saved-modeled-subset", saved.Status); Assert.Empty(saved.CompatibilityIssues); Assert.Empty(saved.Output!.Issues);
        Assert.Contains(saved.Diagnostics, d => d.Code == "text.rich-preserved");
        Assert.Contains(saved.Diagnostics, d => d.Code == "text.attachments-preserved");
        Assert.All(report.Paths.Where(p => p.Mode is not ("style" or "unchanged")), p => Assert.Equal("edit-rejected", p.Status));
    }
    [Theory]
    [InlineData("protected-print-template", "worksheet/sheetProtection", "数据")]
    [InlineData("phonetic-rich-text", "text/phonetic", "数据")]
    public void ImportedFeatureLocationsChangeToLossOnEditAndUndoRestoresPreservation(string fixture, string feature, string? sheetName)
    {
        var folder = NewEvidenceDirectory(); Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "source.xlsx");
        var bytes = Fixture(fixture).Bytes; File.WriteAllBytes(path, bytes);
        var storage = new TypedWorkbookStorage(); var snapshot = storage.Open(path);
        var sourceReport = TypedWorkbookStorage.InspectCompatibility(snapshot);
        Assert.True(sourceReport.OriginalPackageUnchanged);
        Assert.Contains(sourceReport.Diagnostics, d => d.Feature == feature && d.SheetName == sheetName && d.Kind == CompatibilityKind.PreservedOnly && d.Part is not null);
        if (fixture == "phonetic-rich-text")
        {
            Assert.Contains(sourceReport.Diagnostics, d => d.Feature == feature && d.Range == "A2");
            Assert.DoesNotContain(sourceReport.Diagnostics, d => d.Feature == "relationship/hyperlink");
        }
        using var session = new DocumentSession(snapshot);
        session.Execute(new SetCellsCommand([new(new(snapshot.Sheets[0].Id, 0, 5), CellContent.FromLiteral(CellValue.FromNumber(99)))]));
        Assert.Contains(TypedWorkbookStorage.InspectCompatibility(session.Snapshot).Diagnostics, d => d.Feature == feature && d.Kind == CompatibilityKind.PotentialLoss);
        var rejectedPath = Path.Combine(folder, "blocked.xlsq");
        Assert.Throws<WorkbookCompatibilityException>(() => storage.Save(rejectedPath, session.Snapshot));
        Assert.False(File.Exists(rejectedPath));
        session.Undo();
        Assert.True(TypedWorkbookStorage.InspectCompatibility(session.Snapshot).OriginalPackageUnchanged);
        var preservedPath = Path.Combine(folder, "preserved.xlsq"); storage.Save(preservedPath, session.Snapshot);
        Assert.Equal(bytes, ReadParts(File.ReadAllBytes(preservedPath))["workbook.xlsx"]);
        Assert.Equal(snapshot.FileState!.CompatibilityDiagnostics.ToArray(), storage.Open(preservedPath).FileState!.CompatibilityDiagnostics.ToArray());
        Assert.Equal(bytes, File.ReadAllBytes(path));
    }

    [Fact]
    public void ProbeReportsUnknownFunctionsAndReadOnlyArrayPreservationSeparately()
    {
        var report = ProbeRunner.Run(Fixture("unknown-function"), NewEvidenceDirectory());
        var unchanged = Assert.Single(report.Paths, p => p.Mode == "unchanged");
        Assert.Contains(unchanged.Diagnostics, d => d.Code == "formula.unsupported-function" && d.SheetName == "数据" && d.Range == "A2");
        Assert.Empty(unchanged.CompatibilityIssues); // Computability does not itself imply format loss.
        Assert.DoesNotContain(Assert.Single(report.Paths, p => p.Mode == "value").Diagnostics, d => d.Kind == CompatibilityKind.CannotCalculate);
        var readOnly = ProbeRunner.Run(Fixture("array-formula"), NewEvidenceDirectory());
        Assert.All(readOnly.Paths, p => {
            var diagnostic = Assert.Single(p.Diagnostics, d => d.Code == "formula.preserved-read-only");
            Assert.Equal(CompatibilityKind.PreservedOnly, diagnostic.Kind);
            Assert.Equal("数据", diagnostic.SheetName);
            Assert.Equal("B2:B3", diagnostic.Range);
            Assert.Equal("formula/Array", diagnostic.Feature);
            if (p.Mode == "unchanged") { Assert.Equal("preserved-original", p.Status); Assert.Equal(readOnly.Source.Sha256, p.Output!.Sha256); }
            else { Assert.Equal("edit-rejected", p.Status); Assert.Null(p.Output); }
        });
    }

    [Fact]
    public void RepeatedUnsupportedCellFeaturesCapLocationsWithoutRejectingOrWeakeningSaveGuard()
    {
        var parts = ReadParts(Fixture("roster").Bytes);
        FixtureCatalog.Edit(parts, "xl/worksheets/sheet1.xml", root => root.Element(Main + "sheetData")!.ReplaceNodes(
            Enumerable.Range(1, 1030).Select(row => new XElement(Main + "row", new XAttribute("r", row),
                new XElement(Main + "c", new XAttribute("r", "A" + row), new XAttribute("t", "inlineStr"),
                    new XElement(Main + "is", new XElement(Main + "t", "plain"), new XElement(Main + "rPh", new XAttribute("sb", 0), new XAttribute("eb", 1), new XElement(Main + "t", "unsupported phonetic"))))))));
        var directory = NewEvidenceDirectory(); Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "source.xlsx"); File.WriteAllBytes(path, Zip(parts));
        var snapshot = new TypedWorkbookStorage().Open(path);
        Assert.Equal(1030, snapshot.Sheets[0].Cells.Count);
        Assert.Contains(snapshot.FileState!.CompatibilityDiagnostics, d => d.Code == "report.locations-truncated");
        Assert.Equal(1025, snapshot.FileState.CompatibilityDiagnostics.Length);
        Assert.Equal(1030, CompatibilityDiagnosticCollector.Summarize(snapshot.FileState.CompatibilityDiagnostics
            .Where(d => d.TypeFeature == "text/phonetic")).Rows.Count);
        using var session = new DocumentSession(snapshot);
        session.Execute(new SetCellsCommand([new(new(snapshot.Sheets[0].Id, 1029, 0), CellContent.FromLiteral(CellValue.FromText("changed")))]));
        Assert.Throws<WorkbookCompatibilityException>(() => new TypedWorkbookStorage().Save(Path.Combine(directory, "blocked.xlsq"), session.Snapshot));
    }
    [Fact]
    public void UnsupportedFeatureFirstSeenAfterLocationCapRemainsInReportAndSaveGuard()
    {
        var parts = ReadParts(Fixture("roster").Bytes);
        FixtureCatalog.Edit(parts, "xl/worksheets/sheet1.xml", root => root.Element(Main + "sheetData")!.ReplaceNodes(
            Enumerable.Range(1, 1030).Select(row => new XElement(Main + "row", new XAttribute("r", row),
                row == 1030 ? new XAttribute("futureRowSetting", "1") : null,
                new XElement(Main + "c", new XAttribute("r", "A" + row), new XAttribute("t", "inlineStr"),
                    new XElement(Main + "is", new XElement(Main + "t", "plain"), new XElement(Main + "rPh", new XAttribute("sb", 0), new XAttribute("eb", 1), new XElement(Main + "t", "phonetic"))))))));
        var folder = NewEvidenceDirectory(); Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "late-feature.xlsx"); File.WriteAllBytes(path, Zip(parts));
        var storage = new TypedWorkbookStorage(); var model = storage.Open(path);
        var late = Assert.Single(model.FileState!.CompatibilityDiagnostics, d => d.TypeFeature == "row/@futureRowSetting");
        Assert.Equal(1, late.Summary!.Rows.Count);
        Assert.Contains(new CompatibilityRowSpan(model.Sheets[0].Name, 1030, 1030), late.Summary.Rows.Spans);
        using var session = new DocumentSession(model);
        session.Execute(new RenameSheetCommand(model.Sheets[0].Id, "Changed"));
        Assert.Throws<WorkbookCompatibilityException>(() => storage.Save(Path.Combine(folder, "blocked.xlsx"), session.Snapshot));
    }
}
