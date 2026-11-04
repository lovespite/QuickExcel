using System.Text;
using QuickExcel.Clipboard;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class ClipboardStyleTests
{
    private static StyleDefinition Roundtrip(StyleDefinition style, bool html)
    {
        var region = new ExternalClipboardRegion(1, 1, [new(CellValue.FromText("样式"), style)]);
        return (html ? HtmlClipboardCodec.Read(HtmlClipboardCodec.Write(region)) : SpreadsheetClipboardCodec.Read(SpreadsheetClipboardCodec.Write(region))).Cells[0].Style!;
    }

    [Theory]
    [InlineData(null)] [InlineData("single")] [InlineData("double")] [InlineData("singleAccounting")] [InlineData("doubleAccounting")]
    public void DecorationsRoundtripThroughBothFormats(string? underline)
    {
        foreach (var strike in new[] { false, true }) foreach (var html in new[] { false, true })
        {
            var result = Roundtrip(new() { Underline = underline, Strike = strike }, html);
            Assert.Equal(underline, result.Underline); Assert.Equal(strike, result.Strike);
        }
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)] [InlineData(4)] [InlineData(5)] [InlineData(6)]
    [InlineData(7)]
    public void BorderWeightPatternAndColorRoundtrip(int border)
    {
        foreach (var html in new[] { false, true })
        {
            var result = Roundtrip(new() { LeftBorder = border, LeftBorderColor = border == 0 ? null : "FFFF0000", RightBorder = 5, RightBorderColor = "FF00FF00" }, html);
            Assert.Equal(border, result.LeftBorder); Assert.Equal(border == 0 ? null : "FFFF0000", result.LeftBorderColor);
            Assert.Equal(5, result.RightBorder); Assert.Equal("FF00FF00", result.RightBorderColor); Assert.Equal(0, result.BottomBorder);
        }
    }

    [Theory]
    [InlineData(0)] [InlineData(45)] [InlineData(90)] [InlineData(135)] [InlineData(180)] [InlineData(255)]
    public void RotationAndIndentRoundtrip(int rotation)
    {
        foreach (var html in new[] { false, true })
        {
            var result = Roundtrip(new() { TextRotation = rotation, Indent = 3, Horizontal = 1 }, html);
            Assert.Equal(rotation, result.TextRotation); Assert.Equal(3, result.Indent);
        }
    }

    [Fact]
    public void WholeCellStrikeMarkupImportsWhilePartialDecorationRequiresExplicitValues()
    {
        var html = "<table><tr><td><s>value</s></td><td>plain <s>part</s></td></tr></table>";
        var source = HtmlClipboardCodec.Read(HtmlClipboardCodec.Wrap(html, 0, html.Length));
        Assert.True(source.Cells[0].Style!.Strike); Assert.False(source.Cells[1].Style!.Strike);
        Assert.NotEmpty(source.Limitations); Assert.Throws<NotSupportedException>(() => source.RequireMode(CellPasteMode.Formats));
        source.RequireMode(CellPasteMode.Values);
    }

    [Theory]
    [InlineData("mso-rotate:91")]
    [InlineData("mso-rotate:-91")]
    [InlineData("mso-char-indent-count:-1")]
    public void InvalidStyleValuesCannotEnterWorkbook(string css)
    {
        var html = $"<table><tr><td style='{css}'>value</td></tr></table>";
        Assert.ThrowsAny<Exception>(() => HtmlClipboardCodec.Read(HtmlClipboardCodec.Wrap(html, 0, html.Length)));
    }
}
