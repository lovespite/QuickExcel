using System.Globalization;
using System.IO.Compression;

using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class DateDisplayTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-dates-" + Guid.NewGuid().ToString("N"));
    public DateDisplayTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private static string Display(double serial, string code, ExcelDateSystem system = ExcelDateSystem.Windows1900) =>
        CellDisplayFormatter.Format(CellValue.FromNumber(serial), new() { NumberFormatCode = code }, system, CultureInfo.InvariantCulture);

    [Fact]
    public void SpreadsheetMlEnglishMonthMarkerOverridesHostCultureAndRetainsDateClassification()
    {
        var style = new StyleDefinition { NumberFormatCode = "[ENG]mmmm" };
        Assert.True(CellDisplayFormatter.IsDateFormat(style, 45000.5));
        Assert.False(CellDisplayFormatter.IsTimeOnlyFormat(style, 45000.5));
        Assert.Equal("March", CellDisplayFormatter.Format(CellValue.FromNumber(45000.5), style,
            ExcelDateSystem.Windows1900, CultureInfo.GetCultureInfo("zh-CN")));
    }

    [Theory]
    [InlineData(14, 61, "03-01-00")]
    [InlineData(15, 61, "1-Mar-00")]
    [InlineData(16, 61, "1-Mar")]
    [InlineData(17, 61, "Mar-00")]
    [InlineData(18, 0.5, "12:00 PM")]
    [InlineData(19, 0.5, "12:00:00 PM")]
    [InlineData(20, 0.5, "12:00")]
    [InlineData(21, 0.5, "12:00:00")]
    [InlineData(22, 61.5, "3/1/00 12:00")]
    [InlineData(27, 61, "1900年3月")]
    [InlineData(28, 61, "3月1日")]
    [InlineData(30, 61, "3-1-00")]
    [InlineData(31, 61, "1900年3月1日")]
    [InlineData(32, 0.5, "12时00分")]
    [InlineData(33, 0.5, "12时00分00秒")]
    [InlineData(34, 0.5, "下午12时00分")]
    [InlineData(35, 0.5, "下午12时00分00秒")]
    [InlineData(45, 0.0625, "30:00")]
    [InlineData(46, 1.5, "36:00:00")]
    [InlineData(47, 0.0625, "3000.0")]
    [InlineData(57, 61, "1900年3月")]
    [InlineData(58, 61, "3月1日")]
    public void BuiltInDateAndTimeFormats(int id, double serial, string expected) =>
        Assert.Equal(expected, CellDisplayFormatter.Format(CellValue.FromNumber(serial),
            new() { NumberFormatId = id }, ExcelDateSystem.Windows1900, CultureInfo.InvariantCulture));

    [Theory]
    [InlineData("yyyy-mm-dd hh:mm:ss", "2026-10-27 13:05:09")]
    [InlineData("yyyy年m月d日 h时mm分ss秒", "2026年10月27日 13时05分09秒")]
    [InlineData("[$-409]dddd, mmmm d, yyyy", "Tuesday, October 27, 2026")]
    [InlineData("[$-804]yyyy\"年\"m\"月\"d\"日\"", "2026年10月27日")]
    [InlineData("m/d/yyyy h:mm AM/PM", "10/27/2026 1:05 PM")]
    [InlineData("hh:mm:ss a/p", "01:05:09 p")]
    [InlineData("mm:ss", "05:09")]
    [InlineData("hh:mm:ss.000", "13:05:09.125")]
    [InlineData("mmmmm d", "O 27")]
    [InlineData("[Blue]yyyy\\-mm\\-dd;[Red]yyyy-mm-dd;\"-\";@", "2026-10-27")]
    [InlineData("[>=40000]yyyy-mm-dd;0", "2026-10-27")]
    [InlineData("\"date; \"yyyy-mm-dd", "date; 2026-10-27")]
    [InlineData("yyyy-mm-dd\";time \"hh:mm:ss", "2026-10-27;time 13:05:09")]
    [InlineData("yyyy-mm-dd_);@", "2026-10-27 ")]
    public void CustomFormatsRespectMonthMinuteAndLiteralSyntax(string code, string expected)
    {
        var serial = ExcelSerialDate.FromDateTime(new(2026, 10, 27, 13, 5, 9, 125), ExcelDateSystem.Windows1900).Serial;
        Assert.Equal(expected, Display(serial, code));
    }

    [Theory]
    [InlineData("[h]:mm:ss", "36:00:00")]
    [InlineData("[mm]:ss", "2160:00")]
    [InlineData("[ss].00", "129600.00")]
    public void ElapsedTimeDoesNotWrapAtMidnight(string code, string expected) => Assert.Equal(expected, Display(1.5, code));

    [Theory]
    [InlineData(59, ExcelDateSystem.Windows1900, "1900-02-28")]
    [InlineData(60, ExcelDateSystem.Windows1900, "1900-02-29")]
    [InlineData(61, ExcelDateSystem.Windows1900, "1900-03-01")]
    [InlineData(0, ExcelDateSystem.Windows1900, "1900-01-00")]
    [InlineData(0, ExcelDateSystem.Mac1904, "1904-01-01")]
    [InlineData(59, ExcelDateSystem.Mac1904, "1904-02-29")]
    [InlineData(-1, ExcelDateSystem.Windows1900, "########")]
    [InlineData(1e20, ExcelDateSystem.Windows1900, "########")]
    public void EpochLeapDayAndOverflowAreExplicit(double serial, ExcelDateSystem system, string expected) =>
        Assert.Equal(expected, Display(serial, "yyyy-mm-dd", system));

    [Fact]
    public void FractionalSecondsRoundWithCarryWithoutChangingTheDateOfDateOnlyFormats()
    {
        Assert.Equal("00:00:00.000", Display(1 - 0.0004 / 86400, "hh:mm:ss.000"));
        Assert.Equal("1900-03-01 00:00:00", Display(61 - 0.4 / 86400, "yyyy-mm-dd hh:mm:ss"));
        Assert.Equal("1900-02-29", Display(61 - 0.4 / 86400, "yyyy-mm-dd"));
        Assert.Equal("1900-02-29 12:00", Display(60.5, "yyyy-mm-dd hh:mm"));
        Assert.Equal("12:00 AM", Display(0, "hh:mm AM/PM"));
    }

    [Theory]
    [InlineData("General")]
    [InlineData("0.00")]
    [InlineData("0.00E+00")]
    [InlineData("0%")]
    [InlineData("0\" days\"")]
    [InlineData("\"yyyy-mm-dd\"0")]
    [InlineData("0\\m")]
    [InlineData("[Red]0")]
    [InlineData("[DBNum1]yyyy-mm-dd")]
    [InlineData("yyyy-mm-dd\"")]
    [InlineData("[oops]yyyy-mm-dd")]
    [InlineData("[>=10]yyyy-mm-dd;0")]
    public void BasicNumericFormatsDisplayNumbersAndUnsupportedFormatsKeepGeneralFallback(string code) =>
        Assert.Equal(code switch { "0.00" => "1.50", "0%" => "150%", "0.00E+00" => "1.50E+00",
            "0\" days\"" => "2 days", "\"yyyy-mm-dd\"0" => "yyyy-mm-dd2", "0\\m" => "2m",
            "[Red]0" or "[>=10]yyyy-mm-dd;0" => "2", _ => "1.5" }, Display(1.5, code));

    [Fact]
    public void SectionSelectionAndExplicitFormatOverridesAreRespected()
    {
        Assert.Equal("-", Display(0, "yyyy-mm-dd;yyyy-mm-dd;\"-\""));
        Assert.Equal("", Display(0, "yyyy-mm-dd;yyyy-mm-dd;;@"));
        Assert.Equal("1900-03-01", Display(61, "[>100]hh:mm;[>=1]yyyy-mm-dd;\"zero\""));
        Assert.Equal("1900-03-01", CellDisplayFormatter.Format(CellValue.FromNumber(61),
            new() { NumberFormatId = 20, NumberFormatCode = "yyyy-mm-dd" }, ExcelDateSystem.Windows1900));
        Assert.Equal("1.5", Display(1.5, new string('y', 1025)));
    }

    [Fact]
    public void TextBlankBooleanAndErrorsAreNeverCoerced()
    {
        var style = new StyleDefinition { NumberFormatCode = "yyyy-mm-dd" };
        foreach (var value in new[] { CellValue.Blank, CellValue.FromText("46100"), CellValue.FromBoolean(true), CellValue.FromError(CellErrorCode.Reference) })
            Assert.Equal(CellInput.DisplayText(value), CellDisplayFormatter.Format(value, style, ExcelDateSystem.Windows1900));
    }

    [Theory]
    [InlineData(ExcelDateSystem.Windows1900)]
    [InlineData(ExcelDateSystem.Mac1904)]
    public void ImportedXlsxDisplaysDatesWhileEditCalculateSaveAndReopenPreserveNumbers(ExcelDateSystem system)
    {
        var storage = new TypedWorkbookStorage();
        var services = new WorkbookServices(storage, new WorkbookCalculator(new M3CalculationTests.NoScripts()));
        var serial = ExcelSerialDate.FromDateTime(new(2026, 10, 27), system).Serial;
        var sheet = new WorksheetSnapshot(SheetId.New(), "Dates", new Dictionary<CellPosition, CellSnapshot>
        {
            [new(0, 0)] = new(CellContent.FromLiteral(CellValue.FromNumber(serial), 1)),
            [new(0, 1)] = new(CellContent.FromFormula("=A1+1", 1), new(CellValue.FromNumber(serial + 1), CalculationState.Stale, 0, 0, 0)),
            [new(1, 0)] = new(CellContent.FromLiteral(CellValue.FromNumber(0.5), 2)),
            [new(2, 0)] = new(CellContent.FromLiteral(CellValue.FromNumber(serial), 0))
        });
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, system, [sheet],
            styles: [new(), new() { NumberFormatCode = "yyyy-mm-dd" }, new() { NumberFormatId = 21 }]);
        var xlsx = Path.Combine(_folder, "dates.xlsx");
        var seed = Path.Combine(_folder, "seed.xlsq");
        storage.Save(seed, book);
        using (var zip = ZipFile.OpenRead(seed))
            zip.GetEntry("workbook.xlsx")!.ExtractToFile(xlsx);
        using var doc = WorkbookDocument.Open(xlsx, services);
        Assert.Equal("2026-10-27", doc.GetCellValue(0, 0, 0));
        Assert.Equal("2026-10-28", doc.GetCellValue(0, 0, 1));
        Assert.Equal("12:00:00", doc.GetCellValue(0, 1, 0));
        Assert.Equal(serial.ToString(CultureInfo.InvariantCulture), doc.GetCellValue(0, 2, 0));
        Assert.Equal(serial.ToString(CultureInfo.InvariantCulture), doc.GetCellInput(0, 0, 0));
        Assert.Equal("=A1+1", doc.GetCellInput(0, 0, 1));
        doc.SetCell(0, 0, 0, (serial + 1).ToString(CultureInfo.InvariantCulture));
        doc.Calculate();
        Assert.Equal("2026-10-29", doc.GetCellValue(0, 0, 1));
        foreach (var extension in new[] { "xlsx", "xlsq" })
        {
            var path = Path.Combine(_folder, "saved." + extension);
            if (extension == "xlsx") doc.ExportPortableXlsx(path); else doc.Save(path);
            using var reopened = WorkbookDocument.Open(path, services);
            Assert.Equal("2026-10-28", reopened.GetCellValue(0, 0, 0));
            Assert.Equal("2026-10-29", reopened.GetCellValue(0, 0, 1));
            Assert.Equal(CellValue.FromNumber(serial + 1), reopened.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
        }
    }

    [Fact]
    public void MissingCellStyleInheritsRowThenColumnWhileExplicitGeneralOverridesBoth()
    {
        var storage = new TypedWorkbookStorage();
        var path = Path.Combine(_folder, "inherited.xlsx");
        storage.Export(path, new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(SheetId.New(), "S")], styles: [new(), new() { NumberFormatCode = "yyyy-mm-dd" }, new() { NumberFormatId = 21 }]));
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; entry.Delete();
            using var writer = new StreamWriter(zip.CreateEntry("xl/worksheets/sheet1.xml").Open());
            writer.Write("""
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <cols><col min="1" max="3" style="1"/></cols><sheetData>
                  <row r="1"><c r="A1"><v>61</v></c><c r="B1" s="0"><v>61</v></c></row>
                  <row r="2" s="2" customFormat="1"><c r="A2"><v>0.5</v></c><c r="B2" s="1"><v>61</v></c></row>
                  <row r="3" s="2" customFormat="0"><c r="A3"><v>61</v></c></row>
                  </sheetData>
                </worksheet>
                """);
        }
        var services = new WorkbookServices(storage, new WorkbookCalculator(new M3CalculationTests.NoScripts()));
        using var doc = WorkbookDocument.Open(path, services);
        Assert.Equal("1900-03-01", doc.GetCellValue(0, 0, 0));
        Assert.Equal("61", doc.GetCellValue(0, 0, 1));
        Assert.Equal("12:00:00", doc.GetCellValue(0, 1, 0));
        Assert.Equal("1900-03-01", doc.GetCellValue(0, 1, 1));
        Assert.Equal("1900-03-01", doc.GetCellValue(0, 2, 0));
        doc.SetCell(0, 0, 2, "61");
        Assert.Equal("1900-03-01", doc.GetCellValue(0, 0, 2));
        doc.Save(Path.Combine(_folder, "inherited.xlsq"));
        using var reopened = WorkbookDocument.Open(Path.Combine(_folder, "inherited.xlsq"), services);
        Assert.Equal("1900-03-01", reopened.GetCellValue(0, 0, 0));
        Assert.Equal("61", reopened.GetCellValue(0, 0, 1));
    }
}
