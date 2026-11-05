using System.Text;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class TextExchangeTests
{
    [Theory]
    [InlineData(',')]
    [InlineData('\t')]
    public void ExportsQuotedUnicodeAndTrailingEmptyColumnsWithoutChangingInputs(char separator)
    {
        var sheet = Sheet("D", (0, 0, T("001234567890123456789")), (0, 1, T("中文,\t\"引用\"\n下一行")), (1, 0, T("=1+1")));
        var book = Book(sheet);
        var bytes = WorkbookTextExchange.Export(book, sheet.Id, new(0, 0, 1, 2), new(separator));
        Assert.True(bytes.AsSpan().StartsWith(Encoding.UTF8.Preamble));
        var rows = DelimitedText.ReadRows(Encoding.UTF8.GetString(bytes.AsSpan(3)), separator, 10).ToArray();
        Assert.Equal(2, rows.Length); Assert.All(rows, row => Assert.Equal(3, row.Length));
        Assert.Equal("001234567890123456789", rows[0][0]); Assert.Equal(sheet.GetCell(0, 1)!.Value.TextValue, rows[0][1]);
        Assert.Equal("", rows[0][2]); Assert.Equal("=1+1", rows[1][0]); Assert.Equal(0, book.Revision);
    }

    [Fact]
    public void RejectsStaleMissingResultsOversizedRangesAndCancellation()
    {
        var sheet = Sheet("D", (0, 0, F("=1+1")));
        Assert.Throws<InvalidOperationException>(() => WorkbookTextExchange.Export(Book(sheet), sheet.Id, new(0, 0, 0, 0), new()));
        Assert.Throws<InvalidOperationException>(() => WorkbookTextExchange.Export(Book(sheet), sheet.Id, new(0, 0, 199999, 100), new()));
        Assert.Throws<OperationCanceledException>(() => WorkbookTextExchange.Export(Book(sheet), sheet.Id, new(0, 0, 0, 0), new(), new(true)));
    }

    [Fact]
    public void RawAndDisplayValuesAreExplicitAndBomIsOptional()
    {
        var sheet = Sheet("D", (0, 0, N(1.25)), (0, 1, N(2)));
        var book = Book(sheet);
        Assert.Equal("1.25,2\r\n", Encoding.UTF8.GetString(WorkbookTextExchange.Export(book, sheet.Id, new(0, 0, 0, 1), new(Bom: false))));
    }
}
