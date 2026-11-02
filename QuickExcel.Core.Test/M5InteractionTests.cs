using QuickExcel.Desktop;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public class M5InteractionTests
{
    [Fact]
    public void SparseRowsHaveStableIdentityAndDoNotAllocateAnEmptyRectangle()
    {
        var rows = new VirtualRowList<Row>(100_000, i => new(i), row => row.Index);
        Assert.Equal(0, rows.CreatedCount);
        var last = rows[99_999];
        Assert.Same(last, rows[99_999]);
        Assert.Equal(1, rows.CreatedCount);
        Assert.Equal(99_999, rows.IndexOf(last));
        Assert.Equal(-1, rows.IndexOf(new Row(99_999)));
        Assert.Throws<ArgumentOutOfRangeException>(() => rows[100_000]);
    }

    private sealed record Row(int Index);

    [Fact]
    public void NonRectangularSelectionCannotLeakUnselectedCells()
    {
        Assert.Throws<ArgumentException>(() => GridRange.Rectangle([new(0, 0), new(1, 1)]));
        Assert.Equal(new GridRange(2, 3, 3, 4), GridRange.Rectangle([new(2, 3), new(3, 4), new(2, 4), new(3, 3)]));
    }

    [Theory]
    [InlineData("a\tb\r\nc\td\r\n", 2, 2)]
    [InlineData("a\t", 1, 2)]
    [InlineData("", 1, 1)]
    [InlineData("a\n\n", 2, 1)]
    [InlineData("\"a\tb\"\t\"c\nd\"", 1, 2)]
    [InlineData("\"\"", 1, 1)]
    public void TsvPreservesEmptyFieldsAndOneTrailingRowTerminator(string text, int rows, int columns)
    {
        var result = GridClipboard.Read(text);
        Assert.Equal(rows, result.Length);
        Assert.All(result, row => Assert.Equal(columns, row.Length));
    }

    [Fact]
    public void QuotedMultilineClipboardRoundTripsWithoutDataLoss()
    {
        string[][] data = [["a\tb", "\"quoted\""], ["line1\r\nline2", ""]];
        var result = GridClipboard.Read(GridClipboard.Write(new(0, 0, 1, 1), (r, c) => data[r][c]));
        Assert.Equal(data[0], result[0]);
        Assert.Equal(data[1], result[1]);
    }

    [Theory]
    [InlineData("a\tb\nc")]
    [InlineData("\"unterminated")]
    [InlineData("\"a\"x")]
    public void InvalidTsvIsRejected(string text) => Assert.Throws<ArgumentException>(() => GridClipboard.Read(text));

    [Theory]
    [InlineData("=1+2", false, CellContentKind.Literal, "=1+2")]
    [InlineData("=1+2", true, CellContentKind.Formula, "=1+2")]
    [InlineData("'abc", false, CellContentKind.Literal, "'abc")]
    [InlineData("00123", false, CellContentKind.Literal, "00123")]
    public void FormulaInterpretationIsExplicit(string input, bool allowFormulas, CellContentKind kind, string expected)
    {
        var edit = Assert.Single(GridClipboard.Prepare(input, new(0, 0, 0, 0), allowFormulas));
        var content = CellInput.Parse(edit.Value);
        Assert.Equal(kind, content.Kind);
        Assert.Equal(expected, kind == CellContentKind.Formula ? content.FormulaText : content.LiteralValue.TextValue);
    }

    [Fact]
    public void ScalarFillsRectangleAndMatrixRequiresSingleAnchorOrMatchingSelection()
    {
        var fill = GridClipboard.Prepare("7", new(2, 4, 3, 5), false);
        Assert.Equal(4, fill.Length);
        Assert.Equal((3, 5, "7"), fill[^1]);
        Assert.Equal(4, GridClipboard.Prepare("1\t2\n3\t4", new(2, 4, 2, 4), false).Length);
        Assert.Throws<ArgumentException>(() => GridClipboard.Prepare("1\t2\n3\t4", new(0, 0, 0, 1), false));
    }

    [Fact]
    public void AllPasteBudgetsAreValidatedBeforeCommit()
    {
        Assert.Throws<ArgumentException>(() => GridClipboard.Prepare("1\t2", new(0, WorkbookLimits.MaxColumns - 1, 0, WorkbookLimits.MaxColumns - 1), false));
        Assert.Throws<ArgumentException>(() => GridClipboard.Prepare("1\n2", new(WorkbookLimits.MaxRows - 1, 0, WorkbookLimits.MaxRows - 1, 0), false));
        Assert.Throws<ArgumentException>(() => GridClipboard.Prepare("1", new(0, 0, 200, 100), false));
        Assert.Throws<ArgumentException>(() => GridClipboard.Read(new string('x', GridClipboard.MaxCharacters + 1)));
        Assert.Throws<ArgumentException>(() => GridClipboard.Read(string.Join('\n', Enumerable.Repeat("1", 20_001))));
        Assert.Throws<ArgumentException>(() => GridClipboard.Prepare("ok\t" + new string('x', 32768), new(0, 0, 0, 0), false));
    }

    [Fact]
    public void FormulaDraftRetainsItsOriginalAddressUntilAcceptedOrCanceled()
    {
        var state = new CellEditState();
        state.BeginFormula(1, 2, 3);
        state.BeginFormula(2, 5, 6);
        Assert.Equal(new CellEditState.FormulaDraft(1, 2, 3), state.Formula);
        state.StartComposition();
        Assert.True(state.IsComposing);
        state.EndComposition();
        Assert.NotNull(state.Formula);
        state.EndFormula();
        Assert.Null(state.Formula);
        state.BeginGrid();
        state.Reset();
        Assert.False(state.GridEditing);
        Assert.False(state.IsComposing);
    }
}
