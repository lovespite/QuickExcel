using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class StructuredCutTests
{
    [Theory]
    [InlineData(false, "xlsx")]
    [InlineData(true, "xlsx")]
    [InlineData(false, "xlsq")]
    [InlineData(true, "xlsq")]
    public void CutMovesMergeAndTableTogetherAndRetainsTheirReferences(bool crossSheet, string extension)
    {
        var fixture = StructuredReferenceTests.Fixture();
        var data = fixture.Sheets[0];
        var originalTable = data.Data.Tables[0];
        const string total = "=SUM(Sales[Amount])+$K$1";
        var columns = originalTable.Columns.SetItem(2, originalTable.Columns[2] with { TotalsFunction = "custom", TotalsFormula = total });
        var table = new WorksheetTable(originalTable.Id, originalTable.Name, originalTable.Range, columns, totalsRow: true, totalsRowShown: true,
            style: TableStyles.Resolve(fixture, "TableStyleMedium2"), filter: new(new(0, 0, 2, 2), [new(0, new(values: ["2", "3"]))]),
            sort: new(new(1, 0, 2, 2), [new(1, descending: true)]));
        var cells = data.Cells.SetItem(new(3, 2), F(total)).Add(new(0, 10), N(100))
            .Add(new(0, 3), T("Title")).Add(new(0, 4), F("=A2+$K$1"));
        var source = new WorksheetSnapshot(data.Id, data.Name, cells, mergedCells: ["D1:E2", "M1:N1"],
            data: new() { Tables = [table], DefaultColumnWidth = 12 });
        var destination = new WorksheetSnapshot(SheetId.New(), "New Sheet", [new(new(1, 1), T("Old"))], mergedCells: ["B2:C2"]);
        var summary = Sheet("Summary", (0, 0, F("=SUM('Data'!Sales[Amount])")), (1, 0, F("=SUM('Data'!Sales)")),
            (2, 0, F("=FirstQty")), (3, 0, F("=Data!E1")));
        var book = Book(source, destination, summary).WithDefinedNames([new("FirstQty", "=Data!$A$2")]);
        using var session = new DocumentSession(book);
        var target = crossSheet ? destination.Id : source.Id;
        var clipboard = WorkbookClipboard.Capture(book, source.Id, new(0, 0, 3, 4), cut: true);
        Assert.False(session.CanUndo);
        session.Execute(new PasteCellsCommand(clipboard, target, crossSheet ? new(1, 1, 1, 2) : new(1, 1, 1, 1), 0));
        var result = session.Snapshot.GetSheet(target);
        Assert.Null(session.Snapshot.GetSheet(source.Id).GetCell(0, 0));
        Assert.DoesNotContain("D1:E2", session.Snapshot.GetSheet(source.Id).MergedCells);
        Assert.Contains("M1:N1", session.Snapshot.GetSheet(source.Id).MergedCells);
        Assert.Contains("E2:F3", result.MergedCells);
        Assert.DoesNotContain("B2:C2", result.MergedCells);
        Assert.Equal("Title", result.GetCell(1, 4)!.Value.TextValue);
        var relocated = Assert.Single(session.Snapshot.Sheets.SelectMany(s => s.Data.Tables));
        Assert.Same(relocated, Assert.Single(result.Data.Tables));
        Assert.Equal(table.Id, relocated.Id);
        Assert.Equal(table.Name, relocated.Name);
        Assert.True(table.Style.ContentEquals(relocated.Style));
        Assert.Equal(table.Columns.Select(c => (c.Id, c.Name)), relocated.Columns.Select(c => (c.Id, c.Name)));
        Assert.Equal(new(1, 1, 4, 3), relocated.Range);
        Assert.Equal(new(1, 1, 3, 3), relocated.Filter!.Range);
        Assert.True(relocated.Filter.Columns[1].ContentEquals(table.Filter!.Columns[0]));
        Assert.Equal(new(2, 1, 3, 3), relocated.Sort!.Range);
        Assert.Equal(2, Assert.Single(relocated.Sort.Keys).Column);
        Assert.True(relocated.Sort.Keys[0].Descending);
        if (crossSheet) Assert.Contains("'Data'!$K$1", relocated.Columns[2].TotalsFormula);
        void CheckValues(WorkbookSnapshot snapshot)
        {
            var calc = new WorkbookCalculator(new NoScripts()).Calculate(Request(snapshot));
            var report = snapshot.Sheets.Single(s => s.Name == "Summary");
            var owner = snapshot.Sheets.Single(s => s.Name == result.Name);
            Assert.Equal(CellValue.FromNumber(80), C(calc, report, 0).Value);
            Assert.Equal(CellValue.FromNumber(115), C(calc, report, 1).Value);
            Assert.Equal(CellValue.FromNumber(2), C(calc, report, 2).Value);
            Assert.Equal(CellValue.FromNumber(102), C(calc, report, 3).Value);
            Assert.Equal(CellValue.FromNumber(180), C(calc, owner, 4, 3).Value);
        }
        CheckValues(session.Snapshot);
        Assert.True(session.Undo()); Assert.False(session.CanUndo);
        Assert.Equal(source.MergedCells, session.Snapshot.GetSheet(source.Id).MergedCells);
        Assert.Same(table, Assert.Single(session.Snapshot.GetSheet(source.Id).Data.Tables));
        Assert.Equal("Old", session.Snapshot.GetSheet(destination.Id).GetCell(1, 1)!.Value.TextValue);
        Assert.True(session.Redo()); CheckValues(session.Snapshot);

        var storage = new TypedWorkbookStorage();
        var folder = Path.Combine(Directory.GetCurrentDirectory(), "artifacts", "structured-cut", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "moved." + extension);
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        storage.Save(path, session.Snapshot);
        if (extension == "xlsx")
        {
            using var package = ZipFile.OpenRead(path);
            using var part = package.GetEntry($"xl/tables/table{table.Id}.xml")!.Open();
            var xml = XDocument.Load(part).Root!;
            Assert.Equal("B3:D4", (string?)xml.Element(xml.Name.Namespace + "sortState")?.Attribute("ref"));
        }
        var reopened = storage.Open(path);
        Assert.Empty(reopened.FileState!.CompatibilityIssues);
        var saved = reopened.Sheets.Single(s => s.Name == result.Name);
        Assert.Contains("E2:F3", saved.MergedCells);
        Assert.True(relocated.ContentEquals(Assert.Single(saved.Data.Tables), equivalentFormulas: true));
        CheckValues(reopened);
        using var imported = new DocumentSession(reopened);
        var nextCut = WorkbookClipboard.Capture(reopened, saved.Id, new(1, 1, 4, 5), true);
        imported.Execute(new PasteCellsCommand(nextCut, saved.Id, new(10, 5, 10, 5), reopened.Revision));
        Assert.Equal(new(10, 5, 13, 7), Assert.Single(imported.Snapshot.GetSheet(saved.Id).Data.Tables).Range);
        Assert.Contains("I11:J12", imported.Snapshot.GetSheet(saved.Id).MergedCells);
    }

    [Fact]
    public void FailedTargetsPreserveTheCutForAnotherAttempt()
    {
        var book = StructuredReferenceTests.Fixture();
        var source = book.Sheets[0];
        var target = new WorksheetSnapshot(book.Sheets[1].Id, "Target", mergedCells: ["A4:B5"],
            data: new() { Tables = [new(2, "Existing", new(0, 6, 2, 6), [new(1, "Header")])] });
        book = book.WithSheets([source, target]);
        using var session = new DocumentSession(book);
        Assert.Throws<InvalidOperationException>(() => WorkbookClipboard.Capture(book, source.Id, new(1, 0, 3, 2), true));
        var cut = WorkbookClipboard.Capture(book, source.Id, source.Data.Tables[0].Range, true);
        foreach (var position in new[] { new CellPosition(0, 0), new(0, 5) })
        {
            Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteCellsCommand(cut, target.Id,
                new(position.Row, position.Column, position.Row, position.Column), 0)));
            Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
        }
        session.Execute(new PasteCellsCommand(cut, target.Id, new(6, 0, 6, 0), 0));
        Assert.Equal(2, session.Snapshot.GetSheet(target.Id).Data.Tables.Length);
    }

    [Fact]
    public void InheritedBlankStylesTravelWithTheMergeWithoutChangingRowOrColumnDefaults()
    {
        var source = new WorksheetSnapshot(SheetId.New(), "S", [new(new(0, 0), T("Title"))], mergedCells: ["A1:B2"],
            rows: [new(1, new(StyleId: 1))], columns: [new(4, new(StyleId: 2))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [source], styles: [new(), new() { Bold = true }, new() { Italic = true }]);
        using var session = new DocumentSession(book);
        session.Execute(new PasteCellsCommand(WorkbookClipboard.Capture(book, source.Id, new(0, 0, 1, 1), true),
            source.Id, new(3, 3, 3, 3), 0));
        var moved = session.Snapshot.Sheets[0];
        Assert.Equal(1, moved.GetStyleId(4, 4)); Assert.Equal(0, moved.GetStyleId(3, 4));
        Assert.Equal(2, moved.GetStyleId(5, 4)); Assert.Equal(1, moved.GetStyleId(1, 1));
        Assert.Null(moved.GetCell(1, 1)); Assert.Contains("D4:E5", moved.MergedCells);
        Assert.True(session.Undo()); Assert.Null(session.Snapshot.Sheets[0].GetCell(4, 4));
    }

    [Fact]
    public void FilteredTableRowsPreventMoveAtSourceOrDestination()
    {
        var fixture = StructuredReferenceTests.Fixture();
        var source = fixture.Sheets[0];
        var table = source.Data.Tables[0].WithFilter(new(new(0, 0, 2, 2), [new(0, new(values: ["2"]))]));
        var filtered = source.WithData(source.Data with { Tables = [table] });
        var book = fixture.WithSheets([filtered, fixture.Sheets[1]]);
        using var session = new DocumentSession(book);
        Assert.Throws<InvalidOperationException>(() => WorkbookClipboard.Capture(book, source.Id, table.Range, true));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new MoveCellsCommand(fixture.Sheets[1].Id, new(0, 0), new(0, 0), source.Id, new(2, 5))));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }
}
