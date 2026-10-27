using System.Collections.Immutable;
using System.IO.Compression;
using System.Text.Json;
using ExcelIO;
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
    public void ImportPreservesLiteralTextIdentityAndOnlyInterpretsFormulaCache()
    {
        var legacy = new XlWorkbook();
        var sheet = legacy.NewWorksheet("Sheet1");
        sheet.AddRow("00123", "TRUE", "#VALUE!", "", "42");
        sheet[0].Cells[4].SetFormula("6*7", "42");
        var snapshot = LegacyWorkbookAdapter.FromWorkbook(legacy, "throw new Error('must not run');");
        foreach (var column in Enumerable.Range(0, 4))
        {
            var cell = snapshot.Sheets[0].GetCell(0, column)!;
            Assert.Equal(CellValueKind.Text, cell.Content.LiteralValue.Kind);
            Assert.Equal(sheet[0].Cells[column].Value, cell.Content.LiteralValue.TextValue);
        }
        var formula = snapshot.Sheets[0].GetCell(0, 4)!;
        Assert.Equal("=6*7", formula.Content.FormulaText);
        Assert.Equal(42, formula.Calculation!.Value!.Value.NumberValue);
        Assert.Equal(CalculationState.Stale, formula.Calculation.State);
        Assert.Equal("throw new Error('must not run');", snapshot.ScriptSource);
    }

    [Fact]
    public void AdaptersDeepCopyStyleMetadataAndImageBytesInBothDirections()
    {
        var source = new XlWorkbook();
        var sheet = source.NewWorksheet("Styled");
        sheet.AddRow("hello");
        var style = new XlStyle
        {
            Bold = true, FontName = "Calibri", FontSize = 12, FillColor = "FF00FF00",
            Alignment = new XlAlignment { Horizontal = XlHorizontalAlignment.Right, WrapText = true },
            Border = new XlBorder { Left = XlBorderStyle.Thin, LeftColor = "FFFF0000" }
        };
        sheet[0].Cells[0].Style = style;
        sheet[0].Style = style;
        sheet[0].Height = 28;
        sheet.Columns[0] = new XlColumn { Width = 24, Hidden = true, Style = style };
        sheet.MergedCells.Add("A1:B1");
        sheet.Options.ShowGridLines = false;
        sheet.Options.TabColor = "FF123456";
        sheet.AddImage(Pixel, "png", 3, 2, placeInCell: true);
        var snapshot = LegacyWorkbookAdapter.FromWorkbook(source);
        style.Bold = false;
        style.Alignment.Horizontal = XlHorizontalAlignment.Left;
        style.Border.Left = XlBorderStyle.Thick;
        sheet.Images[0].Bytes[0] = 0;
        sheet.MergedCells.Clear();

        var copied = LegacyWorkbookAdapter.ToWorkbook(snapshot);
        Assert.True(copied[0][0].Cells[0].Style!.Bold);
        Assert.Equal(XlHorizontalAlignment.Right, copied[0][0].Cells[0].Style!.Alignment!.Horizontal);
        Assert.Equal(XlBorderStyle.Thin, copied[0][0].Cells[0].Style!.Border!.Left);
        Assert.Equal(28, copied[0][0].Height);
        Assert.True(copied[0].Columns[0].Hidden);
        Assert.Equal(24, copied[0].Columns[0].Width);
        Assert.Equal(new[] { "A1:B1" }, copied[0].MergedCells);
        Assert.False(copied[0].Options.ShowGridLines);
        Assert.Equal("FF123456", copied[0].Options.TabColor);
        Assert.Equal(Pixel, copied[0].Images[0].Bytes);
        Assert.NotSame(copied[0][0].Style, copied[0][0].Cells[0].Style);
        copied[0][0].Cells[0].Style!.Alignment!.WrapText = false;
        copied[0][0].Cells[0].Style!.Border!.LeftColor = "FF000000";
        copied[0].Images[0].Bytes[0] = 1;
        var originalStyle = snapshot.Styles[snapshot.Sheets[0].GetCell(0, 0)!.Content.StyleId];
        Assert.True(originalStyle.WrapText);
        Assert.Equal("FFFF0000", originalStyle.LeftBorderColor);
        Assert.Equal(Pixel[0], snapshot.Sheets[0].Images[0].Data[0]);
    }

    [Fact]
    public void PackageKeepsV1AndSparseCoordinatesWhileReportingLiteralTypeLoss()
    {
        var cells = new Dictionary<CellPosition, CellSnapshot>
        {
            [new(3, 12)] = new(CellContent.FromLiteral(CellValue.FromNumber(123))),
            [new(10, 2)] = new(CellContent.FromFormula("=6*7"), new CellCalculation(CellValue.FromNumber(42), CalculationState.Fresh, 0, 0, 0))
        };
        var snapshot = Snapshot(cells);
        var storage = new LegacyWorkbookStorage();
        var path = FilePath("roundtrip.xlsq");
        storage.Save(path, snapshot);
        var loaded = storage.Open(path);
        Assert.Equal(2, loaded.Sheets[0].Cells.Count);
        Assert.Null(loaded.Sheets[0].GetCell(3, 0));
        Assert.Equal(CellValue.FromText("123"), loaded.Sheets[0].GetCell(3, 12)!.Content.LiteralValue);
        Assert.Equal("=6*7", loaded.Sheets[0].GetCell(10, 2)!.Content.FormulaText);
        Assert.Equal(CalculationState.Stale, loaded.Sheets[0].GetCell(10, 2)!.Calculation!.State);
        Assert.Contains("文本", LegacyWorkbookStorage.LegacyTypeWarning);
        Assert.Equal(CellValue.FromNumber(123), snapshot.Sheets[0].GetCell(3, 12)!.Content.LiteralValue);
        using var archive = ZipFile.OpenRead(path);
        using var manifest = JsonDocument.Parse(archive.GetEntry("manifest.json")!.Open());
        Assert.Equal(1, manifest.RootElement.GetProperty("formatVersion").GetInt32());
        Assert.Equal(3, archive.Entries.Count);
    }

    [Fact]
    public void SavePreservesSupportedStylesSizesMergesAndImages()
    {
        var workbook = new XlWorkbook();
        var sheet = workbook.NewWorksheet("Supported");
        sheet.AddRow("header");
        sheet[0].Height = 31;
        sheet[0].Cells[0].Style = new XlStyle
        {
            FontName = "Calibri", FontSize = 14, Bold = true, FillColor = "FFCCFFCC",
            Alignment = new XlAlignment { Horizontal = XlHorizontalAlignment.Center },
            Border = new XlBorder { Bottom = XlBorderStyle.Thin, BottomColor = "FF000000" }
        };
        sheet.Columns[0] = new XlColumn { Width = 29 };
        sheet.MergedCells.Add("A1:B1");
        sheet.AddImage(Pixel, "png", 4, 3);
        var storage = new LegacyWorkbookStorage();
        var path = FilePath("styles.xlsq");
        storage.Save(path, LegacyWorkbookAdapter.FromWorkbook(workbook));
        var result = LegacyWorkbookAdapter.ToWorkbook(storage.Open(path));
        Assert.Equal(31, result[0][0].Height);
        Assert.Equal(29, result[0].Columns[0].Width);
        Assert.Contains("A1:B1", result[0].MergedCells);
        Assert.True(result[0][0].Cells[0].Style!.Bold);
        Assert.Equal("FFCCFFCC", result[0][0].Cells[0].Style!.FillColor);
        Assert.Equal(XlHorizontalAlignment.Center, result[0][0].Cells[0].Style!.Alignment!.Horizontal);
        Assert.Equal(XlBorderStyle.Thin, result[0][0].Cells[0].Style!.Border!.Bottom);
        Assert.Equal(Pixel, Assert.Single(result[0].Images).Bytes);
    }

    [Fact]
    public void ExportUsesCacheAndLeavesFormulaAndImageSnapshotUntouched()
    {
        var source = Snapshot(new Dictionary<CellPosition, CellSnapshot>
        {
            [new(0, 0)] = new(CellContent.FromFormula("=6*7"), new CellCalculation(CellValue.FromNumber(42), CalculationState.Fresh, 0, 0, 0))
        }, [new SheetImage(Pixel, "png", 40, 10, placeInCell: true)]);
        var storage = new LegacyWorkbookStorage();
        var path = FilePath("values.xlsx");
        storage.Export(path, source);
        var loaded = storage.Open(path);
        Assert.Equal(CellContentKind.Literal, loaded.Sheets[0].GetCell(0, 0)!.Content.Kind);
        Assert.Equal(CellValue.FromText("42"), loaded.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Single(source.Sheets[0].Cells);
        Assert.Equal("=6*7", source.Sheets[0].GetCell(0, 0)!.Content.FormulaText);
        Assert.Equal(Pixel, source.Sheets[0].Images[0].Data.ToArray());
    }

    [Fact]
    public void DenseLegacyConversionBudgetRejectsSparseExpansionBeforeAllocationOrFileReplacement()
    {
        var path = FilePath("budget.xlsq");
        File.WriteAllText(path, "previous bytes");
        var sparse = Enumerable.Range(0, 977).ToDictionary(row => new CellPosition(row, 1023),
            _ => new CellSnapshot(CellContent.FromLiteral(CellValue.FromText("one"))));
        var snapshot = Snapshot(sparse);
        Assert.Equal(977, snapshot.Sheets[0].Cells.Count);
        Assert.Throws<InvalidDataException>(() => new LegacyWorkbookStorage().Save(path, snapshot));
        Assert.Equal("previous bytes", File.ReadAllText(path));
        Assert.Empty(Directory.GetFiles(_folder, "*.tmp"));
    }

    [Fact]
    public void FailedReplacementLeavesPreviousPackageAndSnapshotUntouched()
    {
        if (!OperatingSystem.IsWindows()) return; // The product and its replacement semantics target Windows x64.
        var storage = new LegacyWorkbookStorage();
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
