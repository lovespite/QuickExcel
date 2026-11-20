using System.Text;
using System.Text.Json;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class TextImportTests
{
    [Fact]
    public void RealExcelUtf8CsvMatchesIndependentlyCapturedText()
    {
        using var fixture = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "text-exchange-excel.json")));
        var root = fixture.RootElement;
        var actual = TextImportReader.Read(Encoding.UTF8.GetBytes(root.GetProperty("text").GetString()!), ',', TextFileEncoding.Utf8);
        var expected = root.GetProperty("rows").EnumerateArray().Select(row => row.EnumerateArray().Select(cell => cell.GetString()!).ToArray()).ToArray();
        Assert.Equal(expected.Length, actual.Length);
        for (var i = 0; i < actual.Length; i++) Assert.Equal(expected[i], actual[i]);
        var book = TextImport.Convert(actual, [CleaningColumnType.Text, CleaningColumnType.Text, CleaningColumnType.Text], new()).Workbook;
        Assert.Equal("001234567890123456789", book.Sheets[0].GetCell(0, 0)!.Value.TextValue);
    }
    [Fact]
    public void DefaultTextPreservesIdentifiersFormulasNewlinesAndTrailingEmptyFields()
    {
        var rows = TextImportReader.Read(Encoding.UTF8.GetBytes("00123456789012345678,=1+1,\"中文\n文字\",\r\n"), ',', TextFileEncoding.Utf8);
        var result = TextImport.Convert(rows, Enumerable.Repeat(CleaningColumnType.Text, 4).ToArray(), new());
        Assert.Equal(4, result.Workbook.ScriptModules.ScriptApiVersion);
        var sheet = result.Workbook.Sheets[0];
        Assert.Equal("00123456789012345678", sheet.GetCell(0, 0)!.Value.TextValue);
        Assert.Equal(CellContentKind.Literal, sheet.GetCell(0, 1)!.Content.Kind);
        Assert.Equal("=1+1", sheet.GetCell(0, 1)!.Value.TextValue);
        Assert.Equal("中文\n文字", sheet.GetCell(0, 2)!.Value.TextValue);
        Assert.Equal(4, WorkbookTextExchange.UsedRange(sheet).Columns);
    }

    [Theory]
    [InlineData(ExcelDateSystem.Windows1900, 1462)]
    [InlineData(ExcelDateSystem.Mac1904, 0)]
    public void ExplicitDatesAndRegionalNumbersReuseSplitRules(ExcelDateSystem system, double expected)
    {
        var result = TextImport.Convert([["Date", "Amount"], ["1904/1/1", "1.234,50"]],
            [CleaningColumnType.DateYmd, CleaningColumnType.Number], new(',', '.'), system, header: true);
        var sheet = result.Workbook.Sheets[0];
        Assert.Equal("Date", sheet.GetCell(0, 0)!.Value.TextValue);
        Assert.Equal(expected, sheet.GetCell(1, 0)!.Value.NumberValue);
        Assert.Equal(1234.5, sheet.GetCell(1, 1)!.Value.NumberValue);
        Assert.Equal("yyyy/m/d", result.Workbook.Styles[sheet.GetStyleId(1, 0)].NumberFormatCode);
    }

    [Fact]
    public void InvalidColumnsAndCancellationRejectWholeImport()
    {
        Assert.Throws<ArgumentException>(() => TextImport.Convert([["1"], ["not a number"]], [CleaningColumnType.Number], new()));
        Assert.Throws<ArgumentException>(() => TextImport.Convert([["1234567890123456"]], [CleaningColumnType.Number], new()));
        Assert.Throws<OperationCanceledException>(() => TextImport.Convert([["x"]], [CleaningColumnType.Text], new(), cancellation: new(true)));
        Assert.Throws<DecoderFallbackException>(() => TextImportReader.Read([0xff], ',', TextFileEncoding.Utf8));
    }

    [Theory]
    [InlineData(TextFileEncoding.Utf8, 65001)]
    [InlineData(TextFileEncoding.Utf16LittleEndian, 1200)]
    [InlineData(TextFileEncoding.Utf16BigEndian, 1201)]
    [InlineData(TextFileEncoding.Gb18030, 54936)]
    public void ExplicitEncodingPreservesChinese(TextFileEncoding kind, int page)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); var encoding = Encoding.GetEncoding(page);
        var bytes = encoding.GetPreamble().Concat(encoding.GetBytes("中文\t001\t")).ToArray();
        var rows = TextImportReader.Read(bytes, '\t', kind);
        Assert.Equal(new[] { "中文", "001", "" }, rows.Single());
    }

    [Fact]
    public void SaveReopenPreservesTypedData()
    {
        var result = TextImport.Convert([["001", "2.5", "2026/11/5"]], [CleaningColumnType.Text, CleaningColumnType.Number, CleaningColumnType.DateYmd], new());
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsq");
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, result.Workbook); var reopened = storage.Open(path).Sheets[0];
            Assert.Equal("001", reopened.GetCell(0, 0)!.Value.TextValue); Assert.Equal(2.5, reopened.GetCell(0, 1)!.Value.NumberValue);
            Assert.Equal(result.Workbook.Sheets[0].GetCell(0, 2)!.Value, reopened.GetCell(0, 2)!.Value);
        }
        finally { File.Delete(path); }
    }
}
