using System.Text;
using OpenMcdf;
using QuickExcel.Clipboard;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class Biff8ClipboardTests
{
    // Hand-authored BIFF records independent of the production reader. CFB only supplies the envelope.
    private static byte[] Fixture(Action<BinaryWriter> records, int rows = 1, int columns = 1)
    {
        using var binary = new MemoryStream(); using var w = new BinaryWriter(binary, Encoding.UTF8, true);
        void Bof(ushort type) { w.Write((ushort)0x0809); w.Write((ushort)16); w.Write((ushort)0x0600); w.Write(type); w.Write(new byte[12]); }
        Bof(5); w.Write((ushort)0x0085); w.Write((ushort)8); w.Write((uint)36); w.Write(new byte[4]); w.Write((ushort)0x000A); w.Write((ushort)0);
        Bof(16); w.Write((ushort)0x0200); w.Write((ushort)14); w.Write(3); w.Write(3 + rows); w.Write((ushort)2); w.Write((ushort)(2 + columns)); w.Write((ushort)0);
        records(w); w.Write((ushort)0x000A); w.Write((ushort)0); w.Flush();
        using var output = new MemoryStream();
        using (var root = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen))
        { using var stream = root.CreateStream("Workbook"); stream.Write(binary.ToArray()); }
        return output.ToArray();
    }
    private static void Number(BinaryWriter w, double value, ushort row = 3, ushort col = 2)
    { w.Write((ushort)0x0203); w.Write((ushort)14); w.Write(row); w.Write(col); w.Write((ushort)0); w.Write(value); }
    [Fact]
    public void BoundedCompanionRetainsOriginalNumberAndDimensions()
    {
        var payload = Fixture(w => Number(w, 12.5)); var result = Biff8ClipboardNumbers.Read(payload);
        Assert.Equal(3, result.Top); Assert.Equal(2, result.Left); Assert.Equal(12.5, result.Numbers[new(0, 0)]);
    }
    [Fact]
    public void RkMulRkAndNumericFormulaCacheAreDecodedWithoutExecutingTokens()
    {
        var payload = Fixture(w =>
        {
            w.Write((ushort)0x027E); w.Write((ushort)10); w.Write((ushort)3); w.Write((ushort)2); w.Write((ushort)0); w.Write((uint)(1250 << 2 | 3));
            w.Write((ushort)0x00BD); w.Write((ushort)18); w.Write((ushort)3); w.Write((ushort)3);
            w.Write((ushort)0); w.Write((uint)(8 << 2 | 2)); w.Write((ushort)0); w.Write(unchecked((uint)(-5 << 2 | 2))); w.Write((ushort)4);
            w.Write((ushort)0x0006); w.Write((ushort)22); w.Write((ushort)3); w.Write((ushort)5); w.Write((ushort)0); w.Write(1.25); w.Write(new byte[8]);
        }, columns: 4);
        Assert.Equal(new[] { 12.5, 8d, -5d, 1.25 }, Biff8ClipboardNumbers.Read(payload).Numbers.OrderBy(p => p.Key.Column).Select(p => p.Value));
    }
    [Theory]
    [InlineData("duplicate")] [InlineData("outside")] [InlineData("nan")]
    public void InvalidNumbersRejectWholePayload(string kind)
    {
        var payload = Fixture(w => { Number(w, kind == "nan" ? double.NaN : 1, col: kind == "outside" ? (ushort)3 : (ushort)2); if (kind == "duplicate") Number(w, 2); });
        Assert.Throws<InvalidDataException>(() => Biff8ClipboardNumbers.Read(payload));
    }
    [Fact]
    public void CorruptContainerAndSizeBudgetsFailClosed()
    {
        Assert.ThrowsAny<Exception>(() => Biff8ClipboardNumbers.Read(new byte[1024]));
        Assert.Throws<InvalidDataException>(() => Biff8ClipboardNumbers.Read(new byte[SpreadsheetClipboardCodec.MaxBytes + 1]));
        var payload = Fixture(w => Number(w, .5)); Assert.ThrowsAny<Exception>(() => Biff8ClipboardNumbers.Read(payload[..600]));
    }
    [Theory]
    [InlineData("1904-01-13T12:00:00.000", "[h]:mm:ss", 12.5, false)]
    [InlineData("1904-01-01T12:00:00.000", "hh:mm:ss", .5, false)]
    [InlineData("2027-03-16T12:00:00.000", "yyyy-mm-dd hh:mm:ss", 45000.5, true)]
    [InlineData("2027-03-16T12:00:00.000", "[ENG]mmmm", 45000.5, true)]
    public void RawNumbersDisambiguateTimeAndPreserveCalendarDates(string calendar, string format, double raw, bool date)
    {
        var xml = Encoding.UTF8.GetBytes($"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"><Styles><Style ss:ID=\"Default\"><NumberFormat ss:Format=\"{format}\"/></Style></Styles><Worksheet><Table ss:ExpandedRowCount=\"1\" ss:ExpandedColumnCount=\"1\"><Row><Cell><Data ss:Type=\"DateTime\">{calendar}</Data></Cell></Row></Table></Worksheet></Workbook>");
        var result = SpreadsheetClipboardCodec.Read(xml, Fixture(w => Number(w, raw)));
        Assert.Equal(raw, result.Cells[0].Value.NumberValue);
        Assert.Equal(date, result.Cells[0].Date.HasValue);
        if (date) Assert.Equal(ExcelDateSystem.Mac1904, result.Cells[0].Date!.Value.DateSystem);
        else Assert.Throws<NotSupportedException>(() => SpreadsheetClipboardCodec.Read(xml));
        Assert.Throws<InvalidDataException>(() => SpreadsheetClipboardCodec.Read(xml, Fixture(w => Number(w, raw + 1))));
    }
}
