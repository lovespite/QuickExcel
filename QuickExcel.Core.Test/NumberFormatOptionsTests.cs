using System.Globalization;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class NumberFormatOptionsTests
{
    // Display strings cross-checked with a separate Excel COM workbook, including '?' padding.
    [Theory]
    [InlineData("# ?/?", 1.25, "1 1/4")]
    [InlineData("# ?/?", -2.5, "-2 1/2")]
    [InlineData("# ?/?", .3, " 2/7")]
    [InlineData("# ?/?", .9999, "1    ")]
    [InlineData("# ?/?", 0, "0    ")]
    [InlineData("# ??/??", 1.25, "1  1/4 ")]
    [InlineData("# ??/??", 12.3, "12  3/10")]
    [InlineData("# ???/???", .3, "   3/10 ")]
    [InlineData("# ???/???", 0, "0        ")]
    [InlineData("# ?/8", 1.25, "1 2/8")]
    [InlineData("# ??/16", .3, "  5/16")]
    [InlineData("# ??/100", .9999, "1       ")]
    public void FractionDisplayPreservesPaddingAndRounding(string code, double number, string expected)
        => Assert.Equal(expected, CellDisplayFormatter.Format(CellValue.FromNumber(number), new() { NumberFormatCode = code }, ExcelDateSystem.Windows1900, CultureInfo.InvariantCulture));

    [Fact]
    public void FractionBuiltinsColorsAndUnsupportedCases()
    {
        Assert.Equal("1 1/4", CellDisplayFormatter.Format(CellValue.FromNumber(1.25), new() { NumberFormatId = 12 }, ExcelDateSystem.Windows1900));
        Assert.Equal("1  1/4 ", CellDisplayFormatter.Format(CellValue.FromNumber(1.25), new() { NumberFormatId = 13 }, ExcelDateSystem.Windows1900));
        var display = CellDisplayFormatter.FormatDisplay(CellValue.FromNumber(-1.25), new() { NumberFormatCode = "# ?/?;[Red](# ?/?)" }, ExcelDateSystem.Windows1900);
        Assert.Equal("(1 1/4)", display.Text); Assert.Equal("FFFF0000", display.Color);
        foreach (var code in new[] { "# ?/0", "# ?/10000", "# ????/????", "# ?/?/?", "# ?/1?", "0\"oops\"0", "[DBNum1]0", "[颜色1]0", "0;0;0;0;0", "\"unclosed" })
            Assert.False(CellDisplayFormatter.TryNormalizeNumberFormat(code, out _), code);
        Assert.False(CellDisplayFormatter.TryAdjustDecimalPlaces(new() { NumberFormatId = 12 }, CellValue.FromNumber(1.25), 1, out _));
        Assert.False(CellDisplayFormatter.TryNormalizeNumberFormat(new string('0', 1025), out _));
    }

    [Fact]
    public void PresetsAndNumericSettingsAlwaysHaveRenderableCodes()
    {
        var codes = NumberFormatOptions.Quick.Select(p => p.Code).Concat(NumberFormatOptions.CustomCodes)
            .Concat(Enum.GetValues<NumberFormatCategory>().SelectMany(NumberFormatOptions.Types).Select(p => p.Code));
        foreach (var code in codes) Assert.True(CellDisplayFormatter.TryNormalizeNumberFormat(code, out _), code);
        foreach (var kind in new[] { NumberFormatCategory.Number, NumberFormatCategory.Currency, NumberFormatCategory.Accounting, NumberFormatCategory.Percentage, NumberFormatCategory.Scientific })
            foreach (var digits in new[] { 0, 2, 30 })
                foreach (var symbol in NumberFormatOptions.Symbols)
                    foreach (var negative in Enumerable.Range(0, 4))
                        Assert.True(CellDisplayFormatter.TryNormalizeNumberFormat(NumberFormatOptions.Numeric(kind, digits, true, symbol, negative), out _));
        Assert.True(CellDisplayFormatter.TryNormalizeNumberFormat("G/通用格式", out var general)); Assert.Equal("General", general);
        Assert.True(CellDisplayFormatter.TryNormalizeNumberFormat("0.00;[红色](0.00)", out var localized)); Assert.Equal("0.00;[Red](0.00)", localized);
        Assert.Equal(NumberFormatCategory.Custom, NumberFormatOptions.Category("[>=10]0.000;[Red]0"));
        foreach (var category in new[] { NumberFormatCategory.Number, NumberFormatCategory.Currency, NumberFormatCategory.Accounting, NumberFormatCategory.Percentage, NumberFormatCategory.Scientific })
        {
            var code = NumberFormatOptions.Numeric(category, 3, true, "€", 3);
            var settings = NumberFormatOptions.ReadNumeric(code)!;
            Assert.Equal(category, settings.Category); Assert.Equal(3, settings.Decimals);
            Assert.Equal(code, NumberFormatOptions.Numeric(settings.Category, settings.Decimals, settings.Grouping, settings.Symbol, settings.Negative));
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void PresetsRoundTripWithoutChangingValuesAndUndoAsOneTransaction(string extension)
    {
        var sheet = Sheet("S", (0, 0, N(1.25)), (1, 0, F("=1+2")));
        using var session = new DocumentSession(Book(sheet));
        foreach (var preset in NumberFormatOptions.Quick)
        {
            var before = session.Snapshot;
            session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0), new(sheet.Id, 1, 0)], preset.Patch));
            Assert.Equal(CellValue.FromNumber(1.25), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
            Assert.Equal("=1+2", session.Snapshot.Sheets[0].GetCell(1, 0)!.Content.FormulaText);
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            try
            {
                var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var reopened = storage.Open(path);
                var reopenedStyle = reopened.Styles[reopened.Sheets[0].GetStyleId(0, 0)];
                Assert.Equal(preset.Code, NumberFormatOptions.Code(reopenedStyle));
            }
            finally { File.Delete(path); }
            if (session.Snapshot.Revision != before.Revision)
            { Assert.True(session.Undo()); Assert.Equal(before.Sheets[0].GetStyleId(0, 0), session.Snapshot.Sheets[0].GetStyleId(0, 0)); Assert.True(session.Redo()); }
        }
    }

    [Theory]
    [InlineData("123", "123")]
    [InlineData("00123", "00123")]
    [InlineData("=1+2", "=1+2")]
    [InlineData("TRUE", "TRUE")]
    [InlineData("'123", "123")]
    public void TextFormatControlsNewEntryWithoutCoercingExistingValues(string input, string expected)
    {
        var sheet = Sheet("S", (0, 0, N(42)));
        using var session = new DocumentSession(Book(sheet));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new(NumberFormatId: 49)));
        Assert.Equal(CellValue.FromNumber(42), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        Assert.Equal(CellValue.FromText(expected), CellInput.ParseForCell(input, session.Snapshot, session.Snapshot.Sheets[0], 0, 0).LiteralValue);
        Assert.Equal(CellValue.Blank, CellInput.ParseForCell("", session.Snapshot, session.Snapshot.Sheets[0], 0, 0).LiteralValue);
    }
}
