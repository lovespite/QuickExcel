using QuickExcel.Core.Test.Fixtures;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class StorageAdapterTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-storage-" + Guid.NewGuid().ToString("N"));
    private static readonly byte[] Pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aA1sAAAAASUVORK5CYII=");

    public StorageAdapterTests() => Directory.CreateDirectory(_folder);
    private string FilePath(string name) => Path.Combine(_folder, name);

    [Fact]
    public void V1ImportPreservesLiteralTextIdentityAndTypedFormulaCache()
    {
        var path = FilePath("text-v1.xlsq");
        // Shared strings reproduce the old writer's text-only literals, including an explicit empty cell.
        var xlsx = WorkbookPackageFixture.Xlsx("""
            <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c><c r="E1"><f>6*7</f><v>42</v></c></row>
            """, new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = """
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                      <Default Extension="xml" ContentType="application/xml"/>
                      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                    </Types>
                    """,
                ["xl/sharedStrings.xml"] = """
                    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>00123</t></si><si><t>TRUE</t></si><si><t>#VALUE!</t></si><si><t></t></si></sst>
                    """,
                ["xl/_rels/workbook.xml.rels"] = """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>
                    """
            });
        const string script = "throw new Error('must not run');";
        WorkbookPackageFixture.WriteV1(path, xlsx, script);
        var snapshot = new TypedWorkbookStorage().Open(path);
        string[] texts = ["00123", "TRUE", "#VALUE!", ""];
        for (var column = 0; column < texts.Length; column++)
            Assert.Equal(CellValue.FromText(texts[column]), snapshot.Sheets[0].GetCell(0, column)!.Value);
        var formula = snapshot.Sheets[0].GetCell(0, 4)!;
        Assert.Equal("=6*7", formula.Content.FormulaText);
        Assert.Equal(CellValue.FromNumber(42), formula.Value);
        Assert.Equal(CalculationState.Stale, formula.Calculation!.State);
        Assert.Equal(script, snapshot.ScriptSource);
        using var session = new DocumentSession(snapshot);
        Assert.True(session.IsDirty);
        Assert.False(session.ScriptsEnabled);
    }

    [Fact]
    public void V1MigrationKeepsSparseCoordinatesAndDoesNotRewriteTheSource()
    {
        var path = FilePath("sparse-v1.xlsq");
        WorkbookPackageFixture.WriteV1(path, WorkbookPackageFixture.Xlsx("""
            <row r="4"><c r="M4" t="inlineStr"><is><t>123</t></is></c></row>
            <row r="11"><c r="C11"><f>6*7</f><v>42</v></c></row>
            """));
        var original = File.ReadAllBytes(path);
        var storage = new TypedWorkbookStorage();
        var imported = storage.Open(path);
        var upgraded = FilePath("upgraded.xlsq");
        storage.Save(upgraded, imported);
        var loaded = storage.Open(upgraded);
        Assert.Equal(2, loaded.Sheets[0].Cells.Count);
        Assert.Null(loaded.Sheets[0].GetCell(3, 0));
        Assert.Equal(CellValue.FromText("123"), loaded.Sheets[0].GetCell(3, 12)!.Value);
        Assert.Equal("=6*7", loaded.Sheets[0].GetCell(10, 2)!.Content.FormulaText);
        Assert.Equal(CalculationState.Stale, loaded.Sheets[0].GetCell(10, 2)!.Calculation!.State);
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Equal(1, imported.FileState!.SourceFormatVersion);
        Assert.Equal(2, loaded.FileState!.SourceFormatVersion);
    }

    [Fact]
    public void SavePreservesSupportedStylesSizesMergesAndImagesWithoutMutatingSnapshot()
    {
        var style = new StyleDefinition { FontName = "Calibri", FontSize = 14, Bold = true, FillColor = "FFCCFFCC",
            Horizontal = 2, BottomBorder = 1, BottomBorderColor = "FF000000" };
        var bytes = Pixel.ToArray();
        var sheet = new WorksheetSnapshot(SheetId.New(), "Supported",
            [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("header"), 1)))],
            rows: [new(0, new(31, false, 1))], columns: [new(0, new(29, true, 1))],
            mergedCells: ["A1:B1"], images: [new SheetImage(bytes, "png", 4, 3)],
            showGridLines: false, tabColor: "FF123456");
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), style]);
        bytes[0] = 0;
        var path = FilePath("styles.xlsq");
        var storage = new TypedWorkbookStorage();
        storage.Save(path, snapshot);
        var loaded = storage.Open(path);
        var result = loaded.Sheets[0];
        Assert.Equal(style, loaded.Styles[result.GetCell(0, 0)!.Content.StyleId]);
        Assert.Equal(sheet.Rows, result.Rows);
        Assert.Equal(sheet.Columns, result.Columns);
        Assert.Equal(sheet.MergedCells.ToArray(), result.MergedCells.ToArray());
        Assert.False(result.ShowGridLines);
        Assert.Equal("FF123456", result.TabColor);
        Assert.Equal(Pixel, Assert.Single(result.Images).Data.ToArray());
        Assert.Equal(Pixel, Assert.Single(sheet.Images).Data.ToArray());
        Assert.Single(sheet.Cells);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExportPreservesSnapshotAndRejectsUnsupportedInCellImages(bool placeInCell)
    {
        var source = Snapshot(new Dictionary<CellPosition, CellSnapshot>
        {
            [new(0, 0)] = new(CellContent.FromFormula("=6*7"), new CellCalculation(CellValue.FromNumber(42), CalculationState.Fresh, 0, 0, 0))
        }, [new SheetImage(Pixel, "png", 40, 10, placeInCell: placeInCell)]);
        var storage = new TypedWorkbookStorage();
        var path = FilePath("values.xlsx");
        if (placeInCell)
        {
            File.WriteAllText(path, "previous bytes");
            Assert.Throws<NotSupportedException>(() => storage.Export(path, source));
            Assert.Equal("previous bytes", File.ReadAllText(path));
            Assert.Empty(Directory.GetFiles(_folder, "*.tmp"));
        }
        else
        {
            storage.Export(path, source);
            var loaded = storage.Open(path);
            Assert.Equal(CellContentKind.Literal, loaded.Sheets[0].GetCell(0, 0)!.Content.Kind);
            Assert.Equal(CellValue.FromNumber(42), loaded.Sheets[0].GetCell(0, 0)!.Value);
        }
        Assert.Single(source.Sheets[0].Cells);
        Assert.Equal("=6*7", source.Sheets[0].GetCell(0, 0)!.Content.FormulaText);
        Assert.Equal(Pixel, source.Sheets[0].Images[0].Data.ToArray());
    }

    [Fact]
    public void FailedReplacementLeavesPreviousPackageAndSnapshotUntouched()
    {
        if (!OperatingSystem.IsWindows()) return; // The product and its replacement semantics target Windows x64.
        var storage = new TypedWorkbookStorage();
        var path = FilePath("locked.xlsq");
        var snapshot = Snapshot(new Dictionary<CellPosition, CellSnapshot>
        { [new(0, 0)] = new(CellContent.FromLiteral(CellValue.FromText("original"))) });
        storage.Save(path, snapshot);
        var previous = File.ReadAllBytes(path);
        var changed = snapshot.WithScriptSource("// changed");
        using (var held = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            Assert.ThrowsAny<IOException>(() => storage.Save(path, changed));
            Assert.Equal(previous, File.ReadAllBytes(path));
        }
        Assert.Equal("// changed", changed.ScriptSource);
        Assert.Equal("original", changed.Sheets[0].GetCell(0, 0)!.Value.TextValue);
        Assert.Empty(Directory.GetFiles(_folder, "*.tmp"));
    }

    private static WorkbookSnapshot Snapshot(Dictionary<CellPosition, CellSnapshot> cells, IEnumerable<SheetImage>? images = null) =>
        new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new WorksheetSnapshot(SheetId.New(), "Sheet1", cells, images: images)]);

    public void Dispose() => Directory.Delete(_folder, recursive: true);
}
