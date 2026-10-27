using System.Collections.Immutable;
using System.Globalization;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class WorkbookSnapshotTests
{
    [Fact]
    public void RemoteCellAllocatesOnlyOneSparseEntry()
    {
        var sheet = Sheet("Sparse").WithCells(ImmutableDictionary<CellPosition, CellSnapshot>.Empty
            .Add(new(99_999, 1_023), Literal("remote")));

        Assert.Single(sheet.Cells);
        Assert.Empty(sheet.Rows);
        Assert.Empty(sheet.Columns);
        Assert.Equal(100_000, sheet.UsedRowCount);
        Assert.Equal(1_024, sheet.UsedColumnCount);
        Assert.Null(sheet.GetCell(0, 0));
        Assert.Equal(CellValue.FromText("remote"), sheet.GetCell(99_999, 1_023)!.Value);
    }

    [Fact]
    public void ChangingAndDeletingCellsDoesNotMutateOlderSnapshots()
    {
        var source = new Dictionary<CellPosition, CellSnapshot> { [new(4, 8)] = Literal("before") };
        var initial = new WorksheetSnapshot(SheetId.New(), "Sheet1", source);
        source.Clear();
        var edited = initial.WithCells(initial.Cells.SetItem(new(4, 8), Literal("after")));
        var cleared = edited.WithCells(edited.Cells.Remove(new(4, 8)));

        Assert.Equal(CellValue.FromText("before"), initial.GetCell(4, 8)!.Value);
        Assert.Equal(CellValue.FromText("after"), edited.GetCell(4, 8)!.Value);
        Assert.Null(cleared.GetCell(4, 8));
        Assert.Equal(0, cleared.UsedRowCount);
        Assert.Equal(0, cleared.UsedColumnCount);
    }

    [Fact]
    public void SeededSparseChangesAgreeWithAReferenceMap()
    {
        var random = new Random(20261022);
        var reference = new Dictionary<CellPosition, string>();
        var sheet = Sheet("Random");
        var positions = Enumerable.Range(0, 50)
            .Select(_ => new CellPosition(random.Next(WorkbookLimits.MaxRows), random.Next(WorkbookLimits.MaxColumns)))
            .ToArray();
        for (var index = 0; index < 500; index++)
        {
            var previous = sheet;
            var position = positions[random.Next(positions.Length)];
            var previousCell = previous.GetCell(position.Row, position.Column);
            if (random.Next(3) == 0)
            {
                reference.Remove(position);
                sheet = sheet.WithCells(sheet.Cells.Remove(position));
            }
            else
            {
                var text = index.ToString(CultureInfo.InvariantCulture);
                reference[position] = text;
                sheet = sheet.WithCells(sheet.Cells.SetItem(position, Literal(text)));
            }
            Assert.Same(previousCell, previous.GetCell(position.Row, position.Column));
            Assert.Equal(reference.Count, sheet.Cells.Count);
            Assert.Equal(reference.Count == 0 ? 0 : reference.Keys.Max(key => key.Row) + 1, sheet.UsedRowCount);
            Assert.Equal(reference.Count == 0 ? 0 : reference.Keys.Max(key => key.Column) + 1, sheet.UsedColumnCount);
            foreach (var (key, text) in reference)
                Assert.Equal(CellValue.FromText(text), sheet.GetCell(key.Row, key.Column)!.Value);
        }
    }

    [Fact]
    public void RenamingAndReorderingPreservesSheetIdentityAndPreviousState()
    {
        var first = Sheet("First");
        var second = Sheet("Second");
        var source = new List<WorksheetSnapshot> { first, second };
        var original = Book(source);
        source.Reverse();
        var updated = original.WithSheets([second, first.WithName("Renamed")]).WithRevision(1).WithScriptSource("// script");

        Assert.Equal(first.Id, original.Sheets[0].Id);
        Assert.Equal("First", original.GetSheet(first.Id).Name);
        Assert.Equal("Renamed", updated.GetSheet(first.Id).Name);
        Assert.Equal(second.Id, updated.Sheets[0].Id);
        Assert.Equal(original.DocumentId, updated.DocumentId);
        Assert.Equal(0, original.Revision);
        Assert.Equal(1, updated.Revision);
        Assert.Equal("", original.ScriptSource);
        Assert.Equal("// script", updated.ScriptSource);
        Assert.Throws<KeyNotFoundException>(() => updated.GetSheet(SheetId.New()));
    }

    [Fact]
    public void ImagesStylesAndMetadataAreCopiedAndPreservedByEdits()
    {
        byte[] bytes = [1, 2, 3];
        var image = new SheetImage(bytes, ".PNG", 2, 3, 4, 5, true);
        var images = new List<SheetImage> { image };
        var styles = new List<StyleDefinition> { new(), new() { Bold = true, FillColor = "FFFF0000" } };
        var rows = new Dictionary<int, RowMetadata> { [2] = new(25, true, 1) };
        var columns = new Dictionary<int, ColumnMetadata> { [3] = new(20, true, 1) };
        var merges = new List<string> { "A1:B2" };
        var sheet = new WorksheetSnapshot(SheetId.New(), "Source", rows: rows, columns: columns,
            mergedCells: merges, images: images, showGridLines: false, defaultRowHeight: 18, tabColor: "FF00FF00");
        var original = Book([sheet], styles);
        bytes[0] = 99;
        styles[1] = new();
        rows.Clear();
        columns.Clear();
        images.Clear();
        merges.Clear();
        var changed = sheet.WithName("Changed").WithCells(sheet.Cells.Add(new(0, 0), Literal("value")));

        Assert.Equal(new byte[] { 1, 2, 3 }, image.Data);
        Assert.Equal("png", image.Extension);
        Assert.True(image.PlaceInCell);
        Assert.True(original.Styles[1].Bold);
        Assert.Equal(new RowMetadata(25, true, 1), changed.Rows[2]);
        Assert.Equal(new ColumnMetadata(20, true, 1), changed.Columns[3]);
        Assert.Equal("A1:B2", Assert.Single(changed.MergedCells));
        Assert.Same(image, Assert.Single(changed.Images));
        Assert.False(changed.ShowGridLines);
        Assert.Equal(18, changed.DefaultRowHeight);
        Assert.Equal("FF00FF00", changed.TabColor);
    }

    [Fact]
    public void FormulaInputRemainsIndependentOfCachedResults()
    {
        var content = CellContent.FromFormula("=1+2");
        var missing = new CellSnapshot(content);
        var calculated = new CellSnapshot(content, new(CellValue.FromNumber(3), CalculationState.Fresh, 0, 0, 0));

        Assert.Equal(CellValue.Blank, missing.Value);
        Assert.Equal(CellValue.FromNumber(3), calculated.Value);
        Assert.Same(content, missing.Content);
        Assert.Same(content, calculated.Content);
        Assert.Null(missing.Calculation);
        Assert.Equal("=1+2", CellInput.EditText(calculated.Content));
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    [InlineData(100_000, 0)]
    [InlineData(0, 1_024)]
    public void InvalidCoordinatesAreRejectedAtSnapshotBoundary(int row, int column)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new WorksheetSnapshot(SheetId.New(), "Sheet1",
            new Dictionary<CellPosition, CellSnapshot> { [new(row, column)] = Literal("x") }));
        Assert.Throws<ArgumentOutOfRangeException>(() => Sheet("Sheet1").GetCell(row, column));
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("Bad/Name")]
    [InlineData("Bad[Name]")]
    [InlineData("'Name")]
    [InlineData("Name'")]
    [InlineData("Bad\nName")]
    [InlineData("12345678901234567890123456789012")]
    public void InvalidSheetNamesAreRejected(string name) =>
        Assert.ThrowsAny<ArgumentException>(() => Sheet(name));

    [Fact]
    public void WorkbookRejectsDuplicateIdentitiesNamesAndInvalidStyleReferences()
    {
        var first = Sheet("First");
        Assert.Throws<ArgumentException>(() => Book([first, first.WithName("Second")]));
        Assert.Throws<ArgumentException>(() => Book([first, Sheet("FIRST")]));
        Assert.Throws<ArgumentException>(() => Book([]));
        Assert.Throws<ArgumentException>(() => Book(Enumerable.Range(0, 65).Select(i => Sheet($"Sheet{i}"))));
        Assert.Throws<ArgumentException>(() => Book([first], []));
        var invalidStyle = first.WithCells(first.Cells.Add(new(0, 0), new(CellContent.FromLiteral(CellValue.Blank, 1))));
        Assert.Throws<ArgumentException>(() => Book([invalidStyle]));
        Assert.Throws<ArgumentException>(() => new WorkbookSnapshot(Guid.Empty, 0, ExcelDateSystem.Windows1900, [first]));
        Assert.Throws<ArgumentException>(() => new WorksheetSnapshot(default, "Sheet1"));
        Assert.Throws<ArgumentOutOfRangeException>(() => Book([first]).WithRevision(-1));
    }

    [Fact]
    public void SnapshotRejectsInvalidMetadataAndTextSizes()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new WorksheetSnapshot(SheetId.New(), "Sheet1",
            rows: new Dictionary<int, RowMetadata> { [0] = new(double.NaN) }));
        Assert.Throws<ArgumentOutOfRangeException>(() => new SheetImage([1], "png", 99_999, 0, 2));
        Assert.Throws<ArgumentException>(() => new CellSnapshot(CellContent.FromLiteral(CellValue.FromText(new('x', 32_768)))));
        Assert.Throws<ArgumentException>(() => CellInput.Parse(new('x', 32_768)));
        Assert.Equal(32_767, CellInput.Parse("'" + new string('x', 32_767)).LiteralValue.TextValue.Length);
    }

    [Theory]
    [InlineData("00123")]
    [InlineData("+00123")]
    [InlineData("-00123")]
    [InlineData("00.5")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("1e999")]
    [InlineData("#VALUE!")]
    [InlineData("=")]
    [InlineData("=   ")]
    [InlineData("1,234")]
    public void UserInputDoesNotMisinterpretIdentifiersOrUnsupportedValues(string text)
    {
        var content = CellInput.Parse(text);
        Assert.Equal(CellContentKind.Literal, content.Kind);
        Assert.Equal(CellValue.FromText(text), content.LiteralValue);
    }

    [Theory]
    [InlineData("0", 0)]
    [InlineData("0.5", 0.5)]
    [InlineData("-12.25", -12.25)]
    [InlineData("1e3", 1000)]
    [InlineData(" 42 ", 42)]
    public void NumericInputIsInvariantAndFinite(string text, double expected)
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            Assert.Equal(CellValue.FromNumber(expected), CellInput.Parse(text).LiteralValue);
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    [Theory]
    [InlineData("")]
    [InlineData("text")]
    [InlineData("00123")]
    [InlineData("123")]
    [InlineData("TRUE")]
    [InlineData("false")]
    [InlineData("=1+2")]
    [InlineData("'")]
    [InlineData("'hello")]
    [InlineData(" 42 ")]
    [InlineData("#DIV/0!")]
    [InlineData(" ")]
    public void TypedTextSurvivesAnEditRoundTrip(string text)
    {
        var original = CellContent.FromLiteral(CellValue.FromText(text), 4);
        Assert.Equal(original, CellInput.Parse(CellInput.EditText(original), original.StyleId));
        Assert.Equal(text, CellInput.DisplayText(original.LiteralValue));
    }

    [Fact]
    public void InputDistinguishesBlankExplicitTextBooleanAndFormula()
    {
        Assert.Equal(CellValue.Blank, CellInput.Parse("").LiteralValue);
        Assert.Equal(CellValue.FromText(""), CellInput.Parse("'").LiteralValue);
        Assert.Equal(CellValue.FromText("TRUE"), CellInput.Parse("'TRUE").LiteralValue);
        Assert.Equal(CellValue.FromBoolean(true), CellInput.Parse("true").LiteralValue);
        Assert.Equal(CellValue.FromBoolean(false), CellInput.Parse("FALSE").LiteralValue);
        Assert.Equal("=A1+2", CellInput.Parse("=A1+2").FormulaText);
        Assert.Equal(7, CellInput.Parse("=A1+2", 7).StyleId);
    }

    [Fact]
    public void ErrorDisplayMapsEverySemanticCodeAndUsesExplicitStorageParsing()
    {
        foreach (var code in Enum.GetValues<CellErrorCode>())
        {
            var value = CellValue.FromError(code);
            var displayed = CellInput.DisplayText(value);
            Assert.True(CellInput.TryFromErrorText(displayed, out var parsed));
            Assert.Equal(value, parsed);
            Assert.Equal(value, CellInput.FromErrorText(displayed));
        }
        Assert.False(CellInput.TryFromErrorText("custom error", out _));
        Assert.Equal(CellValue.FromText("custom error"), CellInput.FromErrorText("custom error"));
    }

    private static WorksheetSnapshot Sheet(string name) => new(SheetId.New(), name);
    private static CellSnapshot Literal(string text) => new(CellContent.FromLiteral(CellValue.FromText(text)));
    private static WorkbookSnapshot Book(IEnumerable<WorksheetSnapshot> sheets, IEnumerable<StyleDefinition>? styles = null) =>
        new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, sheets, styles: styles);
}
