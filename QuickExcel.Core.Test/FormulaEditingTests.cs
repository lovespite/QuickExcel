using QuickExcel.Desktop;
using QuickExcel.Domain;
using System.Text.Json;

namespace QuickExcel.Core.Test;

public class FormulaEditingTests
{
    [Fact]
    public void ExcelCapturedNavigationAndReferenceModesMatchWithinDeclaredCapacity()
    {
        using var fixture = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-formula-entry.json")));
        Assert.Equal("16.0", fixture.RootElement.GetProperty("Version").GetString());
        static CellPosition Position(string address)
        {
            var split = 0; var col = 0;
            while (split < address.Length && char.IsAsciiLetter(address[split])) col = col * 26 + address[split++] - 'A' + 1;
            return new(Math.Min(WorkbookLimits.MaxRows, int.Parse(address[split..])) - 1, Math.Min(256, col) - 1);
        }
        foreach (var item in fixture.RootElement.GetProperty("Navigation").EnumerateArray())
        {
            var cells = new List<KeyValuePair<CellPosition, CellSnapshot>>();
            foreach (var row in new[] { 0, 1, 2, 4 }) cells.Add(new(new(row, 0), new(row == 2 && item.TryGetProperty("FormulaEmptyA3", out var empty) && empty.GetBoolean()
                ? CellContent.FromFormula("=\"\"") : CellContent.FromLiteral(CellValue.FromNumber(row + 1)))));
            var sheet = new WorksheetSnapshot(SheetId.New(), "Entry", cells);
            var direction = item.GetProperty("Direction").GetInt32() == -4121 ? 1 : -1;
            var hidden = item.TryGetProperty("HiddenRow2", out var hiddenRow) && hiddenRow.GetBoolean();
            Assert.Equal(Position(item.GetProperty("To").GetString()!), EntryNavigation.Edge(sheet, Position(item.GetProperty("From").GetString()!), direction, 0, WorkbookLimits.MaxRows, 256,
                r => hidden && r == 1));
        }
        var formula = "=A1+B2";
        foreach (var item in fixture.RootElement.GetProperty("ReferenceModes").EnumerateArray())
        {
            formula = Toggle(formula, 1, formula.Length - 1);
            Assert.Equal(item.GetProperty("Formula").GetString(), formula);
        }
    }
    private static string Toggle(string text, int start, int length = 0)
    {
        var edit = FormulaEditing.ToggleReference(text, start, length)!.Value;
        return text.Remove(edit.Start, edit.Length).Insert(edit.Start, edit.Text);
    }
    [Theory]
    [InlineData("=A1", "=$A$1", "=A$1", "=$A1", "=A1")]
    [InlineData("=SUM(A1:B2", "=SUM($A$1:$B$2", "=SUM(A$1:B$2", "=SUM($A1:$B2", "=SUM(A1:B2")]
    [InlineData("='O''Brien 1'!A1", "='O''Brien 1'!$A$1", "='O''Brien 1'!A$1", "='O''Brien 1'!$A1", "='O''Brien 1'!A1")]
    [InlineData("=Sheet1!A1", "=Sheet1!$A$1", "=Sheet1!A$1", "=Sheet1!$A1", "=Sheet1!A1")]
    public void F4CyclesIncompleteDraftReferencesOnly(string text, string first, string second, string third, string fourth)
    {
        text = Toggle(text, text.Length); Assert.Equal(first, text);
        text = Toggle(text, text.Length); Assert.Equal(second, text);
        text = Toggle(text, text.Length); Assert.Equal(third, text);
        Assert.Equal(fourth, Toggle(text, text.Length));
    }
    [Theory]
    [InlineData("=\"A1\"")]
    [InlineData("=\"text \"\" A1\"")]
    [InlineData("=A1_Name")]
    [InlineData("=1E10")]
    [InlineData("=LOG10(2)")]
    [InlineData("='A1'")]
    [InlineData("=Table[A1]")]
    public void F4DoesNotRewriteStringsNamesNumbersOrUnsupportedTokens(string text) => Assert.Null(FormulaEditing.ToggleReference(text, 1, text.Length - 1));
    [Fact]
    public void WholeColumnsAndMultipleSelectedReferencesAreTokenBased()
    {
        Assert.Equal("=SUM($A:$B)", Toggle("=SUM(A:B)", 6));
        Assert.Equal("=SUM(A:B)", Toggle("=SUM($A:$B)", 6));
        Assert.Equal("=$A$1+\"B2\"+$C$3", Toggle("=A1+\"B2\"+C3", 1, 10));
        Assert.Equal("=A$1+$B2+C3", Toggle("=$A$1+B$2+C3", 1, 8));
    }
    [Theory]
    [InlineData("=su", "SUM")]
    [InlineData("=IF(A1>0,av", "AVERAGE")]
    [InlineData("=rou", "ROUND")]
    public void CompletionUsesOnlyRegisteredFunctions(string text, string name) => Assert.Contains(name, FormulaEditing.Complete(text, text.Length)!.Names);
    [Theory]
    [InlineData("=\"SUM")]
    [InlineData("=Table[SU")]
    [InlineData("=unknown")]
    public void CompletionDoesNotInventFunctions(string text) => Assert.Null(FormulaEditing.Complete(text, text.Length));
    [Theory]
    [InlineData("=IF(A1,SUM(B1,C1),", "IF", 3)]
    [InlineData("=IF(A1,SUM(B1,", "SUM", 2)]
    [InlineData("=IF(\"a,b\"=A1,", "IF", 2)]
    [InlineData("=ROUND((A1+B1),", "ROUND", 2)]
    public void ParameterHintsTrackNestedCallsAndQuotedCommas(string text, string function, int argument)
    {
        var hint = FormulaEditing.Hint(text, text.Length)!; Assert.Equal(function, hint.Name); Assert.Equal(argument, hint.Argument);
    }
    [Theory]
    [InlineData("=SUM(", true)]
    [InlineData("=SUM(A1", true)]
    [InlineData("=A1+", true)]
    [InlineData("=\"A1", false)]
    [InlineData("=SUM(\"x\"", false)]
    [InlineData("=name", false)]
    [InlineData("=123", false)]
    [InlineData("plain", false)]
    public void ReferencePickingRequiresAnOperandOrExistingReference(string text, bool valid) => Assert.Equal(valid, FormulaEditing.ReferenceSlot(text, text.Length, 0).HasValue);
    [Fact]
    public void ReferenceAtCaretResolvesOnlyVisibleReferenceTokens()
    {
        Assert.Equal(new GridRange(0, 0, 0, 0), FormulaEditing.ReferenceAt("=A1+B2", 2)!.Range);
        Assert.Equal(new GridRange(1, 1, 1, 1), FormulaEditing.ReferenceAt("=A1+B2", 5)!.Range);
        Assert.Equal(new GridRange(0, 0, 2, 0), FormulaEditing.ReferenceAt("=SUM(A1:A3", 10)!.Range);
        Assert.Null(FormulaEditing.ReferenceAt("=SUM(A1:A3", 4));
        Assert.Null(FormulaEditing.ReferenceAt("=\"A1\"", 3));
        Assert.Null(FormulaEditing.ReferenceAt("plain A1", 6));
        var qualified = FormulaEditing.ReferenceAt("=Sheet1!A1", 4)!;
        Assert.Equal("Sheet1", qualified.Sheet); Assert.Equal(new GridRange(0, 0, 0, 0), qualified.Range);
        var whole = FormulaEditing.ReferenceAt("=SUM(A:A)", 6)!;
        Assert.Equal(0, whole.Range.Left); Assert.Equal(0, whole.Range.Right);
        Assert.Null(FormulaEditing.ReferenceAt("=A99999999999", 5));
    }
    [Fact]
    public void NavigationFollowsContentBlocksAndSkipsHiddenCells()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Data", [
            new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(1)))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromText("")))),
            new(new(2, 0), new(CellContent.FromFormula("=\"\""))),
            new(new(4, 0), new(CellContent.FromLiteral(CellValue.FromNumber(5)))),
            new(new(8, 0), new(CellContent.FromLiteral(CellValue.Blank, 1)))]);
        CellPosition Down(int row, Func<int, bool>? hidden = null) => EntryNavigation.Edge(sheet, new(row, 0), 1, 0, 20, 10, hidden);
        Assert.Equal(new(2, 0), Down(0)); Assert.Equal(new(4, 0), Down(2)); Assert.Equal(new(4, 0), Down(3)); Assert.Equal(new(19, 0), Down(4));
        Assert.Equal(new(2, 0), Down(0, row => row == 1)); Assert.Equal(new(19, 0), Down(2, row => row == 4));
        Assert.Equal(new(2, 0), EntryNavigation.Edge(sheet, new(4, 0), -1, 0, 20, 10));
        Assert.Equal(new(0, 0), EntryNavigation.Edge(sheet, new(2, 0), -1, 0, 20, 10));
    }
    [Fact]
    public void AutoSumSuggestsAboveBeforeLeftWithoutChangingCells()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Data", [
            new(new(0, 1), new(CellContent.FromLiteral(CellValue.FromNumber(1)))),
            new(new(1, 1), new(CellContent.FromLiteral(CellValue.FromNumber(2)))),
            new(new(2, 0), new(CellContent.FromLiteral(CellValue.FromNumber(3))))]);
        Assert.Equal(new GridRange(0, 1, 1, 1), EntryNavigation.AutoSumSource(sheet, new(2, 1)));
        Assert.Equal(new GridRange(0, 1, 0, 1), EntryNavigation.AutoSumSource(sheet, new(0, 2)));
        Assert.Null(EntryNavigation.AutoSumSource(sheet, new(0, 0)));
    }
}
