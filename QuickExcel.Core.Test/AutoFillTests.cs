using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class AutoFillTests
{
    private static WorkbookSnapshot Book(string[] seeds, bool horizontal = false, int style = 0, ExcelDateSystem system = ExcelDateSystem.Windows1900) =>
        new(Guid.NewGuid(), 0, system, [new(SheetId.New(), "Fill", seeds.Select((text, i) =>
            new KeyValuePair<CellPosition, CellSnapshot>(new(3 + (horizontal ? 0 : i), 3 + (horizontal ? i : 0)), new(CellInput.Parse(text).WithStyle(style)))))],
            styles: [new(), new() { NumberFormatId = 14 }]);

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void NumericSeriesExtendsInAllDirectionsAndHasOneUndo(bool horizontal, bool reverse)
    {
        var book = Book(["2", "5"], horizontal); var sheet = book.Sheets[0];
        var source = new CellRectangle(3, 3, horizontal ? 3 : 4, horizontal ? 4 : 3);
        var extent = horizontal ? source with { Left = reverse ? 0 : 3, Right = reverse ? 4 : 7 }
            : source with { Top = reverse ? 0 : 3, Bottom = reverse ? 4 : 7 };
        using var session = new DocumentSession(book);
        session.Execute(new FillCellsCommand(sheet.Id, source, extent, CellFillMode.Auto));
        Assert.Equal(reverse ? -7 : 14, session.Snapshot.Sheets[0].GetCell(horizontal ? 3 : reverse ? 0 : 7, horizontal ? reverse ? 0 : 7 : 3)!.Value.NumberValue);
        Assert.Equal(sheet.GetCell(3, 3), session.Snapshot.Sheets[0].GetCell(3, 3));
        var result = session.Snapshot.Sheets[0].Cells;
        session.Undo(); Assert.False(session.CanUndo); Assert.Equal(sheet.Cells, session.Snapshot.Sheets[0].Cells);
        session.Redo(); Assert.Equal(result, session.Snapshot.Sheets[0].Cells);
    }

    [Theory]
    [InlineData("7", null, "7", "7")]
    [InlineData("0.1", "0.2", "0.3", "0.4")]
    [InlineData("10", "7", "4", "1")]
    [InlineData("ID007", null, "ID008", "ID009")]
    [InlineData("001", null, "002", "003")]
    [InlineData("编号10", "编号8", "编号6", "编号4")]
    [InlineData("A9", "A10", "A11", "A12")]
    [InlineData("A01", "B02", "A01", "B02")]
    [InlineData("word", null, "word", "word")]
    [InlineData("TRUE", null, "TRUE", "TRUE")]
    public void AutomaticPatternsPreserveTypes(string first, string? second, string expected1, string expected2)
    {
        // Explicit text seeds ensure numeric-looking IDs are never converted to numbers.
        var text = first is "001";
        var book = Book(second is null ? [first] : [first, second]); var id = book.Sheets[0].Id;
        using var session = new DocumentSession(book);
        if (text) session.Execute(new SetCellsCommand([new(new(id, 3, 3), CellContent.FromLiteral(CellValue.FromText(first)))]));
        var bottom = second is null ? 3 : 4;
        session.Execute(new FillCellsCommand(id, new(3, 3, bottom, 3), new(3, 3, bottom + 2, 3), CellFillMode.Auto));
        var actual1 = session.Snapshot.Sheets[0].GetCell(bottom + 1, 3)!.Value;
        var actual2 = session.Snapshot.Sheets[0].GetCell(bottom + 2, 3)!.Value;
        if (text) { Assert.Equal(expected1, actual1.TextValue); Assert.Equal(expected2, actual2.TextValue); }
        else if (CellInput.Parse(expected1).LiteralValue.Kind == CellValueKind.Number)
        {
            Assert.Equal(CellInput.Parse(expected1).LiteralValue.NumberValue, actual1.NumberValue, 12);
            Assert.Equal(CellInput.Parse(expected2).LiteralValue.NumberValue, actual2.NumberValue, 12);
        }
        else { Assert.Equal(CellInput.Parse(expected1).LiteralValue, actual1); Assert.Equal(CellInput.Parse(expected2).LiteralValue, actual2); }
    }

    [Theory]
    [InlineData("1", "2", "4")]
    [InlineData("1", "", "3")]
    [InlineData("1000000000000", "1000000000001", "1000000000003")]
    public void AmbiguousLanesRepeatTheWholePattern(string a, string b, string c)
    {
        var book = Book([a, b, c]); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new FillCellsCommand(id, new(3, 3, 5, 3), new(3, 3, 8, 3), CellFillMode.Auto));
        for (var i = 0; i < 3; i++) Assert.Equal(book.Sheets[0].GetCell(3 + i, 3)!.Value, session.Snapshot.Sheets[0].GetCell(6 + i, 3)?.Value ?? CellValue.Blank);
    }

    [Theory]
    [InlineData(ExcelDateSystem.Windows1900, 59, null, 60, 61)]
    [InlineData(ExcelDateSystem.Mac1904, 0, null, 1, 2)]
    [InlineData(ExcelDateSystem.Windows1900, 31, 60d, 91, 121)]
    [InlineData(ExcelDateSystem.Windows1900, 45351.5, 45716.5, 46081.5, 46446.5)]
    [InlineData(ExcelDateSystem.Windows1900, 45322, 45291d, 45260, 45230)]
    public void DatePatternsRespectCalendarAndDateSystem(ExcelDateSystem system, double start, double? second, double third, double fourth)
    {
        string Text(double v) => v.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        var book = Book(second is { } n ? [Text(start), Text(n)] : [Text(start)], style: 1, system: system);
        using var session = new DocumentSession(book); var bottom = second is null ? 3 : 4;
        session.Execute(new FillCellsCommand(book.Sheets[0].Id, new(3, 3, bottom, 3), new(3, 3, bottom + 2, 3), CellFillMode.Auto));
        Assert.Equal(third, session.Snapshot.Sheets[0].GetCell(bottom + 1, 3)!.Value.NumberValue);
        Assert.Equal(fourth, session.Snapshot.Sheets[0].GetCell(bottom + 2, 3)!.Value.NumberValue);
        Assert.Equal(1, session.Snapshot.Sheets[0].GetStyleId(bottom + 2, 3));
    }

    [Theory]
    [InlineData("ID9223372036854775807", 0)]
    [InlineData("2958465", 1)]
    public void OverflowRejectsTheWholeFill(string seed, int style)
    {
        var book = Book([seed], style: style); using var session = new DocumentSession(book);
        Assert.Throws<ArgumentException>(() => session.Execute(new FillCellsCommand(book.Sheets[0].Id, new(3, 3, 3, 3), new(3, 3, 5, 3), CellFillMode.Auto)));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void ReverseTextUnderflowIsAtomicAndCopyModeAlwaysRepeats()
    {
        var book = Book(["A001"]); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        Assert.Throws<ArgumentException>(() => session.Execute(new FillCellsCommand(id, new(3, 3, 3, 3), new(0, 3, 3, 3), CellFillMode.Auto)));
        Assert.Same(book, session.Snapshot);
        session.Execute(new FillCellsCommand(id, new(3, 3, 3, 3), new(0, 3, 3, 3), CellFillMode.Copy));
        Assert.Equal("A001", session.Snapshot.Sheets[0].GetCell(0, 3)!.Value.TextValue);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void IndependentLanesKeepFormulaReferencesAndAlternatingStylesOnSave(string extension)
    {
        var book = Book(["2", "5"]); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new SetCellsCommand([new(new(id, 3, 4), CellContent.FromFormula("=D4+$A$1")),
            new(new(id, 4, 4), CellContent.FromFormula("=D5+$A$1")), new(new(id, 4, 3), CellContent.FromLiteral(CellValue.FromNumber(5), 1))]));
        // Both numeric cells use non-date styles for numeric lane inference.
        var styles = session.Snapshot.Styles.SetItem(1, new() { Bold = true });
        book = new WorkbookSnapshot(book.DocumentId, session.Snapshot.Revision, book.DateSystem, session.Snapshot.Sheets, styles: styles);
        using var filled = new DocumentSession(book);
        filled.Execute(new FillCellsCommand(id, new(3, 3, 4, 4), new(3, 3, 7, 4), CellFillMode.Auto));
        Assert.Equal(14, filled.Snapshot.Sheets[0].GetCell(7, 3)!.Value.NumberValue);
        Assert.Equal("=(D8+$A$1)", filled.Snapshot.Sheets[0].GetCell(7, 4)!.Content.FormulaText);
        Assert.Equal(0, filled.Snapshot.Sheets[0].GetStyleId(5, 3)); Assert.Equal(1, filled.Snapshot.Sheets[0].GetStyleId(6, 3));
        var path = Path.Combine(AppContext.BaseDirectory, "artifacts", $"autofill-{Guid.NewGuid():N}.{extension}");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!); var storage = new TypedWorkbookStorage(); storage.Save(path, filled.Snapshot);
        Assert.Equal(14, storage.Open(path).Sheets[0].GetCell(7, 3)!.Value.NumberValue);
    }

    [Fact]
    public void RichTextIsCopiedAndHiddenTargetsStillRejectAtomically()
    {
        var book = Book(["ID1"]); using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        var rich = new RichCellText([new("ID", new(Bold: true)), new("1")]);
        session.Execute(new SetCellsCommand([new(new(id, 3, 3), CellContent.FromRichText(rich))]));
        session.Execute(new FillCellsCommand(id, new(3, 3, 3, 3), new(3, 3, 4, 3), CellFillMode.Auto));
        Assert.Equal(rich, session.Snapshot.Sheets[0].GetCell(4, 3)!.Content.RichText);
        session.Execute(new SetAxisMetadataCommand(id, SheetAxis.Row, 5, 1, hidden: true)); var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new FillCellsCommand(id, new(3, 3, 3, 3), new(3, 3, 6, 3), CellFillMode.Auto)));
        Assert.Same(before, session.Snapshot);
    }
}
