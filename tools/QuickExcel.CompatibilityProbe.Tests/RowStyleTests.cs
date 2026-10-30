using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.CompatibilityProbe.OpenXmlAudit;

namespace QuickExcel.CompatibilityProbe.Tests;

public sealed partial class CompatibilityProbeTests
{
    [Theory]
    [InlineData("save")]
    [InlineData("snapshot-export")]
    [InlineData("portable-export")]
    public void RowStyleActivationSurvivesEditingAllWritePathsAndIndependentValidation(string route)
    {
        var directory = NewEvidenceDirectory(); Directory.CreateDirectory(directory);
        var source = Path.Combine(directory, "source.xlsx"); var original = Fixture("row-column-styles").Bytes;
        File.WriteAllBytes(source, original);
        var storage = new TypedWorkbookStorage(); var snapshot = storage.Open(source);
        void Verify(WorkbookSnapshot book)
        {
            var sheet = book.Sheets[0];
            int[] expected = [1, 0, 1, 1, 2, 0, 1];
            for (var row = 0; row < expected.Length; row++)
            {
                Assert.Equal(expected[row], sheet.GetStyleId(row, 2)); // No stored cell; grid and future input use this.
                Assert.Equal(expected[row], sheet.GetStyleId(row, 0));
            }
            Assert.Equal(0, sheet.GetStyleId(0, 1)); Assert.Equal(2, sheet.GetStyleId(1, 1)); Assert.Equal(0, sheet.GetStyleId(4, 1));
            Assert.False(sheet.Rows[0].ApplyStyle); Assert.True(sheet.Rows[1].ApplyStyle);
            Assert.False(sheet.Rows[2].ApplyStyle); Assert.Equal(2, sheet.Rows[2].StyleId);
            Assert.False(sheet.Rows[3].ApplyStyle); Assert.True(sheet.Rows[4].ApplyStyle); Assert.True(sheet.Rows[5].ApplyStyle);
            Assert.Equal(32, sheet.Rows[0].Height);
        }
        Verify(snapshot); Assert.Empty(snapshot.FileState!.CompatibilityIssues);
        using var session = new DocumentSession(snapshot);
        session.Execute(new SetCellsCommand([new(new(snapshot.Sheets[0].Id, 9, 6), CellContent.FromLiteral(CellValue.FromText("edit")))]));
        var output = Path.Combine(directory, route == "save" ? "edited.xlsq" : "edited.xlsx");
        if (route == "save") storage.Save(output, session.Snapshot);
        else if (route == "snapshot-export") storage.Export(output, session.Snapshot);
        else storage.ExportPortable(output, session.Snapshot);
        Verify(storage.Open(output));
        var bytes = File.ReadAllBytes(output); if (route == "save") bytes = ReadParts(bytes)["workbook.xlsx"];
        Assert.Empty(Inspect(bytes).Issues);
        var rows = Xml(ReadParts(bytes)["xl/worksheets/sheet1.xml"]).Descendants(Main + "row").ToArray();
        Assert.Equal("0", (string?)rows.Single(r => (string?)r.Attribute("r") == "1").Attribute("customFormat"));
        Assert.Equal("1", (string?)rows.Single(r => (string?)r.Attribute("r") == "2").Attribute("customFormat"));
        Assert.Equal("0", (string?)rows.Single(r => (string?)r.Attribute("r") == "3").Attribute("customFormat"));
        Assert.Equal(original, File.ReadAllBytes(source));
    }

    [Fact]
    public void ChangingOnlyRowActivationIsDetectedAndDoesNotReuseOriginalPackage()
    {
        var directory = NewEvidenceDirectory(); Directory.CreateDirectory(directory);
        var source = Path.Combine(directory, "source.xlsx"); File.WriteAllBytes(source, Fixture("row-column-styles").Bytes);
        var storage = new TypedWorkbookStorage(); var original = storage.Open(source); var sheet = original.Sheets[0];
        var changedSheet = new WorksheetSnapshot(sheet.Id, sheet.Name, sheet.Cells,
            sheet.Rows.SetItem(1, sheet.Rows[1] with { ApplyStyle = false }), sheet.Columns);
        var changed = original.WithSheets([changedSheet]);
        Assert.False(TypedWorkbookStorage.InspectCompatibility(changed).OriginalPackageUnchanged);
        var path = Path.Combine(directory, "activation.xlsq"); storage.Save(path, changed);
        var reopened = storage.Open(path);
        Assert.False(reopened.Sheets[0].Rows[1].ApplyStyle);
        Assert.Equal(1, reopened.Sheets[0].GetStyleId(1, 2));
        Assert.Equal(0, reopened.Sheets[0].GetStyleId(1, 0)); // A stored cell retains its resolved style.
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("2")]
    public void InvalidRowCustomFormatIsRejectedEvenWithoutExplicitStyle(string value)
    {
        var parts = ReadParts(Fixture("row-column-styles").Bytes);
        FixtureCatalog.Edit(parts, "xl/worksheets/sheet1.xml", root => root.Descendants(Main + "row").First().SetAttributeValue("customFormat", value));
        var directory = NewEvidenceDirectory(); Directory.CreateDirectory(directory);
        var source = Path.Combine(directory, "bad.xlsx"); File.WriteAllBytes(source, Zip(parts));
        Assert.Throws<InvalidDataException>(() => new TypedWorkbookStorage().Open(source));
    }

    [Fact]
    public void RowStyleFixtureRunsAllEightPathsWithoutLossAndKeepsOriginalBytes()
    {
        var report = ProbeRunner.Run(Fixture("row-column-styles"), NewEvidenceDirectory());
        Assert.True(report.SourceUnchanged);
        Assert.All(report.Paths, path => { Assert.NotNull(path.Output); Assert.Empty(path.Output.Issues); Assert.Empty(path.CompatibilityIssues); });
        Assert.Equal(report.Source.Sha256, report.Paths.Single(p => p.Mode == "unchanged").Output!.Sha256);
    }
}
