using System.Text.Json;
using QuickExcel.CompatibilityProbe;
using QuickExcel.Domain;

namespace QuickExcel.CompatibilityProbe.Tests;

public sealed class BusinessBaselineTests
{
    private static JsonDocument Reference() => JsonDocument.Parse("""
        {"Date1904":false,"Sheets":[{"Name":"基线","Cells":[
          {"Row":1,"Column":1,"Kind":"Number","Value":42,"Formula":"=21*2","Text":"42.00","Font":"Arial","Size":11,"Bold":true,"Italic":false},
          {"Row":2,"Column":1,"Kind":"Text","Value":"00123","Formula":null,"Text":"00123","Font":"Arial","Size":11,"Bold":false,"Italic":false}
        ]}]}
        """);
    private static WorkbookSnapshot Book() => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [new(SheetId.New(), "基线", [new(new(0, 0), new(CellContent.FromFormula("=21*2", 1), new(CellValue.FromNumber(42), CalculationState.Stale, 0, 0, 0))),
            new(new(1, 0), new(CellContent.FromLiteral(CellValue.FromText("00123"))))])],
        styles: [new() { FontName = "Arial", FontSize = 11 }, new() { FontName = "Arial", FontSize = 11, Bold = true, NumberFormatId = 2 }]);
    [Fact]
    public void IndependentCaptureChecksTypesFormulasStylesAndDisplay()
    {
        using var expected = Reference();
        var checks = BusinessBaselineProbe.Compare(Book(), expected.RootElement);
        Assert.All(checks, c => Assert.True(c.Passed, c.ToString()));
        Assert.Contains(checks, c => c.Property == "Formula" && c.Expected == "=21*2");
        Assert.Contains(checks, c => c.Property == "Value" && c.Expected == "00123");
    }
    [Theory]
    [InlineData("formula", "Formula")]
    [InlineData("value", "Value")]
    [InlineData("style", "Text")]
    public void SemanticallyWrongButReadableWorkbookFailsAtLocatedProperty(string fault, string property)
    {
        using var expected = Reference(); var book = Book(); var sheet = book.Sheets[0];
        var cells = sheet.Cells.ToDictionary();
        cells[new(0, 0)] = new(CellContent.FromFormula(fault == "formula" ? "=21*3" : "=21*2", fault == "style" ? 0 : 1),
            new(CellValue.FromNumber(fault == "value" ? 63 : 42), CalculationState.Stale, 0, 0, 0));
        book = book.WithSheets([new(sheet.Id, sheet.Name, cells)]);
        Assert.Contains(BusinessBaselineProbe.Compare(book, expected.RootElement), c => !c.Passed && c.Property == property && c.Location == "基线!R1C1");
    }
}
