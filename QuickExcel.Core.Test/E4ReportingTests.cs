using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class E4ReportingTests : IDisposable
{
    private readonly string _folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../validation/e4-unit/" + Guid.NewGuid().ToString("N")));
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/mWQAAAAASUVORK5CYII=");
    public E4ReportingTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private static WorkbookSnapshot Book(WorksheetSnapshot sheet) => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
    private static WorksheetSnapshot Sheet() => new(SheetId.New(), "Report");

    [Theory]
    [InlineData(DrawingAnchorKind.TwoCell, ".xlsx")]
    [InlineData(DrawingAnchorKind.OneCell, ".xlsx")]
    [InlineData(DrawingAnchorKind.Absolute, ".xlsx")]
    [InlineData(DrawingAnchorKind.TwoCell, ".xlsq")]
    [InlineData(DrawingAnchorKind.OneCell, ".xlsq")]
    [InlineData(DrawingAnchorKind.Absolute, ".xlsq")]
    public void ImagesKeepExactOffsetsAndBehaviorAfterEditAndRoundTrip(DrawingAnchorKind kind, string extension)
    {
        var sheet = Sheet(); var anchor = DrawingGeometry.FromBounds(sheet, new(172.25, 36.5, 220.75, 95.125), kind);
        var image = new SheetImage(Png, "png", 0, 0).WithAnchor(anchor);
        using var session = new DocumentSession(Book(sheet)); session.Execute(new SetImageCommand(sheet.Id, 0, image));
        var storage = new TypedWorkbookStorage(); var path = Path.Combine(_folder, "image" + extension); storage.Save(path, session.Snapshot);
        var reopened = storage.Open(path); Assert.Empty(reopened.FileState!.CompatibilityIssues); Assert.Equal(anchor, reopened.Sheets[0].Images[0].Anchor);
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 0, 2));
        var moved = session.Snapshot.Sheets[0].Images[0].Anchor.Bounds(session.Snapshot.Sheets[0]);
        Assert.Equal(kind == DrawingAnchorKind.Absolute ? 36.5 : 98.5, moved.Y, 3); Assert.Equal(95.125, moved.Height, 3);
        session.Undo(); Assert.Equal(anchor, session.Snapshot.Sheets[0].Images[0].Anchor);
        session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 1, 1, size: 62));
        var resized = session.Snapshot.Sheets[0].Images[0].Anchor.Bounds(session.Snapshot.Sheets[0]);
        Assert.Equal(kind == DrawingAnchorKind.TwoCell ? 126.125 : 95.125, resized.Height, 3);
    }

    [Fact]
    public void PaginationHonorsHiddenDimensionsManualBreaksTitlesAndArea()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Report", rows: [new(3, new(Hidden: true))], columns: [new(1, new(Hidden: true))], printing: new()
        {
            Areas = [new(PrintRangeKind.Cells, 0, 0, 9, 2)], Titles = [new(PrintRangeKind.Rows, 0, 0, 0, PrintRange.MaxColumns - 1)],
            RowBreaks = [new(5, 0, PrintRange.MaxColumns - 1, true)], Setup = new() { PaperSize = 9, Orientation = "landscape" }
        });
        var layout = PrintingLayout.Create(sheet); Assert.Equal(2, layout.Pages.Length); Assert.True(layout.PaperWidth > layout.PaperHeight);
        Assert.Equal(new[] { 0, 1, 2, 4 }, layout.Pages[0].Rows); Assert.Equal(new[] { 0, 5, 6, 7, 8, 9 }, layout.Pages[1].Rows);
        Assert.All(layout.Pages, p => Assert.Equal(new[] { 0, 2 }, p.Columns));
    }
    [Fact]
    public void FitAndBudgetAndCancellationAreExplicit()
    {
        var sheet = Sheet().WithPrinting(new() { Areas = [new(PrintRangeKind.Cells, 0, 0, 99, 9)], Properties = new(FitToPage: true), Setup = new() { FitToWidth = 1, FitToHeight = 1 } });
        var layout = PrintingLayout.Create(sheet); Assert.Single(layout.Pages); Assert.InRange(layout.Scale, .1, 1);
        Assert.Equal(1, PrintingLayout.Create(Sheet().WithPrinting(new() { Properties = new(FitToPage: true) })).Scale);
        Assert.Throws<OperationCanceledException>(() => PrintingLayout.Create(sheet, new CancellationToken(true)));
        Assert.Throws<NotSupportedException>(() => PrintingLayout.Create(sheet.WithPrinting(new() { Areas = [new(PrintRangeKind.Cells, 0, 0, 99999, 100)] })));
        Assert.Equal(("Report", "2/8", "file.xlsx"), PrintingLayout.Header("&L&A&C&P/&N&R&F", "Report", "file.xlsx", 2, 8, DateTime.UnixEpoch));
        Assert.Throws<NotSupportedException>(() => PrintingLayout.Header("&G", "", "", 1, 1, DateTime.UnixEpoch));
    }
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void ProtectionSurvivesRoundTripAndBlocksAtomicMixedEdits(string extension)
    {
        var sheet = Sheet(); using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellProtectionCommand(sheet.Id, new(0, 1, 0, 1), false));
        session.Execute(new SetProtectionCommand(sheet.Id, new() { FormatCells = false }));
        var storage = new TypedWorkbookStorage(); var path = Path.Combine(_folder, "protected" + extension); storage.Save(path, session.Snapshot);
        using var reopened = new DocumentSession(storage.Open(path)); var id = reopened.Snapshot.Sheets[0].Id;
        Assert.Empty(reopened.Snapshot.FileState!.CompatibilityIssues); Assert.False(reopened.Snapshot.Styles[reopened.Snapshot.Sheets[0].GetStyleId(0, 1)].Locked);
        var before = reopened.Snapshot;
        Assert.Throws<InvalidOperationException>(() => reopened.Execute(new SetCellsCommand([new(new(id, 0, 1), CellContent.FromLiteral(CellValue.FromNumber(1))), new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(2)))])));
        Assert.Same(before, reopened.Snapshot); Assert.False(reopened.CanUndo);
        reopened.Execute(new SetCellsCommand([new(new(id, 0, 1), CellContent.FromLiteral(CellValue.FromNumber(3), reopened.Snapshot.Sheets[0].GetStyleId(0, 1)))]));
        Assert.Equal(3, reopened.Snapshot.Sheets[0].GetCell(0, 1)!.Value.NumberValue);
        reopened.Execute(new SetCellStyleCommand([new(id, 0, 0)], new(Bold: true)));
        Assert.Throws<InvalidOperationException>(() => reopened.Execute(new SetAxisMetadataCommand(id, SheetAxis.Row, 0, 1, size: 80)));
    }
    [Fact]
    public void NotesCreateEditRemoveAndUndoWithoutDroppingOtherShapes()
    {
        var sheet = Sheet(); using var session = new DocumentSession(Book(sheet));
        var first = new CellNote("作者", new([new TextRun("第一条")])); var second = first with { Text = new([new TextRun("第二条")]) };
        session.Execute(new SetNoteCommand(sheet.Id, new(1, 1), first)); session.Execute(new SetNoteCommand(sheet.Id, new(2, 2), second));
        var before = session.Snapshot.Sheets[0].Annotations.NoteDrawing;
        session.Execute(new SetNoteCommand(sheet.Id, new(1, 1), second)); Assert.True(before.SequenceEqual(session.Snapshot.Sheets[0].Annotations.NoteDrawing));
        session.Execute(new SetNoteCommand(sheet.Id, new(2, 2), null)); Assert.Single(session.Snapshot.Sheets[0].Annotations.Notes);
        session.Undo(); Assert.Equal(2, session.Snapshot.Sheets[0].Annotations.Notes.Count);
        var storage = new TypedWorkbookStorage(); var path = Path.Combine(_folder, "notes.xlsx"); storage.Save(path, session.Snapshot);
        var reopened = storage.Open(path); Assert.Empty(reopened.FileState!.CompatibilityIssues); Assert.True(session.Snapshot.Sheets[0].Annotations.ContentEquals(reopened.Sheets[0].Annotations));
    }

    [Theory]
    [InlineData(WorksheetChartKind.Column, ".xlsx")]
    [InlineData(WorksheetChartKind.Line, ".xlsx")]
    [InlineData(WorksheetChartKind.Pie, ".xlsx")]
    [InlineData(WorksheetChartKind.Column, ".xlsq")]
    public void ChartReferencesCachesAndSourceEditsRoundTrip(WorksheetChartKind kind, string extension)
    {
        var sheet = Sheet(); using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("Month"))), new(new(sheet.Id, 0, 1), CellContent.FromLiteral(CellValue.FromText("Sales"))),
            new(new(sheet.Id, 1, 0), CellContent.FromLiteral(CellValue.FromText("Jan"))), new(new(sheet.Id, 2, 0), CellContent.FromLiteral(CellValue.FromText("Feb"))),
            new(new(sheet.Id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(10))), new(new(sheet.Id, 2, 1), CellContent.FromLiteral(CellValue.FromNumber(-20)))]));
        sheet = session.Snapshot.Sheets[0];
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 2, 1), kind, DrawingGeometry.FromBounds(sheet, new(12, 150, 560, 320), DrawingAnchorKind.OneCell), "Sales");
        session.Execute(new SetChartCommand(sheet.Id, 0, chart));
        session.Execute(new RenameSheetCommand(sheet.Id, "O'Brien"));
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(123.5)))]));
        var storage = new TypedWorkbookStorage(); var path = Path.Combine(_folder, "chart" + extension); storage.Save(path, session.Snapshot);
        var reopened = storage.Open(path); Assert.Empty(reopened.FileState!.CompatibilityIssues);
        var actual = Assert.Single(reopened.Sheets[0].Data.Charts); Assert.True(session.Snapshot.Sheets[0].Data.Charts[0].ContentEquals(actual));
        Assert.Equal(new double?[] { 123.5, -20 }, ChartData.Read(reopened, actual)[0].Values); Assert.Equal("Sales", ChartData.Read(reopened, actual)[0].Name);
        var before = session.Snapshot; Assert.Throws<NotSupportedException>(() => session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 1))); Assert.Same(before, session.Snapshot);
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 1, 1), CellContent.FromFormula("=1+2"))]));
        Assert.Throws<InvalidOperationException>(() => ChartData.Read(session.Snapshot, session.Snapshot.Sheets[0].Data.Charts[0]));
        Assert.Throws<InvalidOperationException>(() => storage.Save(Path.Combine(_folder, "stale.xlsx"), session.Snapshot));
    }

    [Fact]
    public void ImageAndNotePayloadsCountAgainstHistoryBeforeCommit()
    {
        var sheet = Sheet(); using var session = new DocumentSession(Book(sheet), historyBudgetBytes: 4096);
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetImageCommand(sheet.Id, 0, new SheetImage(new byte[8192], "png", 0, 0))));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetNoteCommand(sheet.Id, new(0, 0), new("author", new([new TextRun(new string('x', 4000))])))));
        Assert.Same(before, session.Snapshot);
    }

    [Fact]
    public void NewSparseAndStyleBoundariesAreEnforced()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Sparse", [new(new(WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1), new(CellContent.FromLiteral(CellValue.FromNumber(42))))]);
        var book = Book(sheet); var path = Path.Combine(_folder, "sparse.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, book);
        Assert.Single(storage.Open(path).Sheets[0].Cells);
        Assert.Throws<ArgumentException>(() => new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: Enumerable.Repeat(new StyleDefinition(), WorkbookLimits.MaxStyles + 1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => sheet.GetCell(WorkbookLimits.MaxRows, 0)); Assert.Throws<ArgumentOutOfRangeException>(() => sheet.GetCell(0, WorkbookLimits.MaxColumns));
        using var session = new DocumentSession(book);
        Assert.Throws<ArgumentException>(() => session.Execute(new SetImageCommand(sheet.Id, 0, new SheetImage(Png, "png", 0, 0).WithAnchor(
            new() { Kind = DrawingAnchorKind.Absolute, X = 10_000_000_000_000L, Width = 9525, Height = 9525 }))));
        Assert.Same(book, session.Snapshot);
    }

    [Fact]
    public void AxisEditsRetainDefaultsAndAllowedProtection()
    {
        var sheet = Sheet().WithData(new() { DefaultColumnWidth = 8.43, Protection = new() { InsertRows = false } });
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 0, 1));
        Assert.Equal(sheet.Data.Protection, session.Snapshot.Sheets[0].Data.Protection);
        Assert.Equal(8.43, session.Snapshot.Sheets[0].Data.DefaultColumnWidth);
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Column, 0, 1)));
        session.Undo(); Assert.Equal(sheet.Data.Protection, session.Snapshot.Sheets[0].Data.Protection);
    }

    [Fact]
    public void PrintingAndChartsUseFilteredVisibilityWithoutMutatingTheWorkbook()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Report", [
            new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("Category")))),
            new(new(0, 1), new(CellContent.FromLiteral(CellValue.FromText("Value")))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromText("keep")))),
            new(new(1, 1), new(CellContent.FromLiteral(CellValue.FromNumber(12)))),
            new(new(2, 0), new(CellContent.FromLiteral(CellValue.FromText("hide")))),
            new(new(2, 1), new(CellContent.FromLiteral(CellValue.FromNumber(99))))],
            data: new(new(new(0, 0, 2, 1), [new(0, new(values: ["keep"]))])));
        var book = Book(sheet); var visible = PrintingLayout.VisibleSheet(book, sheet);
        Assert.Equal(new[] { 0, 1 }, Assert.Single(PrintingLayout.Create(visible).Pages).Rows);
        Assert.Empty(sheet.Rows);
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 2, 1), WorksheetChartKind.Column, new() { Kind = DrawingAnchorKind.Absolute, Width = 1000000, Height = 1000000 });
        Assert.Equal(new double?[] { 12 }, ChartData.Read(book, chart)[0].Values);
        Assert.Equal(new double?[] { 12, 99 }, ChartData.Read(book, chart with { VisibleOnly = false })[0].Values);
    }

    [Fact]
    public void PasswordProtectionAndLockedOverridesCannotBeBypassed()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Protected", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(5))))],
            rows: [new(0, new(StyleId: 1))], data: new() { Protection = new() { DeleteRows = false, Password = "AB12" } });
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { Locked = false }]);
        using var session = new DocumentSession(book);
        Assert.Throws<NotSupportedException>(() => session.Execute(new SetProtectionCommand(sheet.Id, null)));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 0, 1, delete: true)));
        Assert.Same(book, session.Snapshot);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void FullStyleTierAndColumnDefaultRoundTrip(string extension)
    {
        var styles = Enumerable.Range(0, WorkbookLimits.MaxStyles).Select(i => new StyleDefinition { FontColor = "FF" + i.ToString("X6") }).ToArray();
        var sheet = new WorksheetSnapshot(SheetId.New(), "Wide", [new(new(199999, 2047), new(CellContent.FromLiteral(CellValue.FromNumber(42), 8191)))], data: new() { DefaultColumnWidth = 8.43 });
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: styles);
        var storage = new TypedWorkbookStorage(); var path = Path.Combine(_folder, "tier" + extension); storage.Save(path, book); var reopened = storage.Open(path);
        Assert.Equal(8192, reopened.Styles.Length); Assert.Equal(styles[8191], reopened.Styles[8191]);
        Assert.Equal(64, DrawingGeometry.ColumnSize(reopened.Sheets[0], 0));
        if (extension == ".xlsq")
        {
            using var zip = System.IO.Compression.ZipFile.OpenRead(path); using var reader = new StreamReader(zip.GetEntry("manifest.json")!.Open());
            var manifest = reader.ReadToEnd(); Assert.Contains("extended-capacity", manifest); Assert.Contains("worksheet-default-width", manifest);
        }
    }

    [Fact]
    public void UnsupportedChartSemanticsRemainOriginalOnly()
    {
        var sheet = Sheet(); var chart = SetChartCommand.FromRange(sheet, new(0, 0, 2, 1), WorksheetChartKind.Column, new() { Kind = DrawingAnchorKind.Absolute, Width = 1000000, Height = 1000000 });
        var storage = new TypedWorkbookStorage(); var path = Path.Combine(_folder, "unsupported.xlsx"); storage.Save(path, Book(sheet.WithData(new() { Charts = [chart] })));
        using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/charts/chart1-1.xml")!; string xml;
            using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
            entry.Delete(); using var writer = new StreamWriter(zip.CreateEntry("xl/charts/chart1-1.xml").Open()); writer.Write(xml.Replace("clustered", "stacked"));
        }
        var source = File.ReadAllBytes(path); var opened = storage.Open(path); Assert.Empty(opened.Sheets[0].Data.Charts);
        Assert.Contains(opened.FileState!.CompatibilityDiagnostics, d => d.Feature == "chart/unsupported");
        var copy = Path.Combine(_folder, "preserved.xlsx"); storage.Save(copy, opened); Assert.Equal(source, File.ReadAllBytes(copy));
        using var session = new DocumentSession(opened); session.Execute(new SetCellsCommand([new(new(opened.Sheets[0].Id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(42)))]));
        Assert.Throws<WorkbookCompatibilityException>(() => storage.Save(Path.Combine(_folder, "changed.xlsx"), session.Snapshot));
    }
}
