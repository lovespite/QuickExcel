using System.Globalization;
using System.Text.Json;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class NumberDisplayTests
{
    public static IEnumerable<object[]> ExcelCases()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-number-formats.json")));
        foreach (var item in json.RootElement.GetProperty("Cases").EnumerateArray())
            yield return [item.GetRawText()];
    }

    [Theory]
    [MemberData(nameof(ExcelCases))]
    public void MatchesExcel16Build20527FixedResults(string json)
    {
        using var item = JsonDocument.Parse(json);
        var row = item.RootElement;
        var value = row.GetProperty("Value");
        var cell = value.ValueKind == JsonValueKind.String ? CellValue.FromText(value.GetString()!) : CellValue.FromNumber(value.GetDouble());
        var display = CellDisplayFormatter.FormatDisplay(cell, new() { NumberFormatCode = row.GetProperty("Format").GetString() }, ExcelDateSystem.Windows1900, CultureInfo.InvariantCulture);
        var expected = row.GetProperty("Text").GetString()!;
        if (expected.Length > 0 && expected.All(c => c == '#')) Assert.Equal("########", display.Text);
        else if (display.FillOffset >= 0)
        {
            // Excel Text repeats '*' to column width; the Core result gives the renderer that same fill point.
            var count = expected.Length - display.Text.Length;
            Assert.True(count >= 0);
            Assert.All(expected.Substring(display.FillOffset, count), c => Assert.Equal(display.FillCharacter, c));
            Assert.Equal(expected.Remove(display.FillOffset, count), display.Text);
        }
        else Assert.Equal(expected, display.Text);
        var bgr = (int)row.GetProperty("Color").GetDouble();
        Assert.Equal($"FF{bgr & 255:X2}{(bgr >> 8) & 255:X2}{(bgr >> 16) & 255:X2}", display.Color ?? "FF000000");
    }

    [Theory]
    [InlineData("# ?/?")]
    [InlineData("[DBNum1]0")]
    [InlineData("[unknown]0")]
    [InlineData("0.0%%")]
    [InlineData("0/0")]
    public void UnsupportedCodesKeepOriginalValueVisible(string code) => Assert.Equal("12.3",
        CellDisplayFormatter.Format(CellValue.FromNumber(12.3), new() { NumberFormatCode = code }, ExcelDateSystem.Windows1900, CultureInfo.InvariantCulture));

    [Fact]
    public void CultureAndTextSectionsDoNotCoerceRawValues()
    {
        var style = new StyleDefinition { NumberFormatCode = "#,##0.000;[Red](#,##0.000);\"zero\";\"ID \"@" };
        Assert.Equal("1.234,500", CellDisplayFormatter.Format(CellValue.FromNumber(1234.5), style, ExcelDateSystem.Windows1900, CultureInfo.GetCultureInfo("de-DE")));
        Assert.Equal("ID 00123", CellDisplayFormatter.Format(CellValue.FromText("00123"), style, ExcelDateSystem.Windows1900));
        Assert.Equal("", CellDisplayFormatter.Format(CellValue.Blank, style, ExcelDateSystem.Windows1900));
        Assert.Equal("TRUE", CellDisplayFormatter.Format(CellValue.FromBoolean(true), style, ExcelDateSystem.Windows1900));
        Assert.Equal("#REF!", CellDisplayFormatter.Format(CellValue.FromError(CellErrorCode.Reference), style, ExcelDateSystem.Windows1900));
        Assert.Equal("\uFFF012.3", CellDisplayFormatter.Format(CellValue.FromNumber(12.3), new() { NumberFormatCode = "\"\uFFF0\"0.0" }, ExcelDateSystem.Windows1900));
    }
}
