using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using QuickExcel.Clipboard;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class ChartClipboardTests
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static (WorkbookSnapshot Book, WorksheetChart Chart) Create(WorksheetChartKind kind = WorksheetChartKind.Column)
    {
        var cells = new List<KeyValuePair<CellPosition, CellSnapshot>>();
        void Cell(int r, int c, CellValue v) => cells.Add(new(new(r, c), new(CellContent.FromLiteral(v))));
        Cell(0, 0, CellValue.FromText("分类")); Cell(0, 1, CellValue.FromText("收入"));
        for (var i = 1; i <= 3; i++) { Cell(i, 0, kind == WorksheetChartKind.Scatter ? CellValue.FromNumber(i) : CellValue.FromText("Q" + i)); Cell(i, 1, CellValue.FromNumber(i * 10)); }
        var sheet = new WorksheetSnapshot(SheetId.New(), "O'Brien", cells);
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 3, 1), kind, new() { Kind = DrawingAnchorKind.OneCell, Width = 4_762_500, Height = 2_857_500 }, "中文图表");
        return (new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet.WithData(sheet.Data with { Charts = [chart] })]), chart);
    }
    private static ChartClipboard Capture(WorksheetChartKind kind = WorksheetChartKind.Column) { var (book, chart) = Create(kind); return ChartClipboard.Capture(book, chart); }
    // Independent minimal Office fixture: external references plus the cache copied by Excel, no workbook cells.
    private static byte[] NativeFixture()
    {
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/charts/chart1-1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>""",
            ["_rels/.rels"] = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="drawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="xl/drawings/drawing1.xml"/></Relationships>""",
            ["xl/drawings/_rels/drawing1.xml.rels"] = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1-1.xml"/></Relationships>""",
            ["xl/drawings/drawing1.xml"] = """
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"><lc:lockedCanvas xmlns:lc="http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas">
                <a:nvGrpSpPr><a:cNvPr id="0" name=""/><a:cNvGrpSpPr/></a:nvGrpSpPr><a:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4762500" cy="2857500"/><a:chOff x="0" y="0"/><a:chExt cx="4762500" cy="2857500"/></a:xfrm></a:grpSpPr>
                <a:graphicFrame><a:nvGraphicFramePr><a:cNvPr id="1" name="Chart"/><a:cNvGraphicFramePr/></a:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="chart"/></a:graphicData></a:graphic><a:xfrm><a:off x="0" y="0"/><a:ext cx="4762500" cy="2857500"/></a:xfrm></a:graphicFrame>
                </lc:lockedCanvas></a:graphicData></a:graphic>
                """,
            ["xl/charts/chart1-1.xml"] = """
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strRef><c:f>[source.xlsx]Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
                <c:val><c:numRef><c:f>[source.xlsx]Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
                <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
                """
        };
        using var output = new MemoryStream(); using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var (name, value) in parts) { using var stream = zip.CreateEntry(name).Open(); stream.Write(Encoding.UTF8.GetBytes(value)); }
        return output.ToArray();
    }
    private static byte[] Rewrite(byte[] bytes, string part, Action<XDocument> change)
    {
        using var source = new MemoryStream(bytes); using var read = new ZipArchive(source); using var output = new MemoryStream();
        using (var write = new ZipArchive(output, ZipArchiveMode.Create, true)) foreach (var entry in read.Entries)
        {
            using var input = entry.Open(); using var target = write.CreateEntry(entry.FullName).Open();
            if (entry.FullName == part) { var xml = XDocument.Load(input); change(xml); xml.Save(target); } else input.CopyTo(target);
        }
        return output.ToArray();
    }

    [Theory]
    [InlineData(WorksheetChartKind.Column)] [InlineData(WorksheetChartKind.Line)] [InlineData(WorksheetChartKind.SmoothLine)] [InlineData(WorksheetChartKind.Scatter)] [InlineData(WorksheetChartKind.Pie)] [InlineData(WorksheetChartKind.Waterfall)] [InlineData(WorksheetChartKind.Funnel)]
    public void EditableSnapshotSurvivesClipboardAndAtomicUndoRedo(WorksheetChartKind kind)
    {
        var clip = OfficeDrawingClipboardCodec.ReadPrivateChart(OfficeDrawingClipboardCodec.WritePrivateChart(Capture(kind)));
        var target = new WorksheetSnapshot(SheetId.New(), "图表数据");
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [target]));
        session.Execute(new PasteChartCommand(clip, target.Id, new(5, 4)));
        Assert.Equal(2, session.Snapshot.Sheets.Length); Assert.Equal("图表数据2", session.Snapshot.Sheets[1].Name);
        Assert.Empty(session.Snapshot.Sheets[0].Cells); var chart = Assert.Single(session.Snapshot.Sheets[0].Data.Charts);
        Assert.Equal(5, chart.Anchor.Row); Assert.Equal(4, chart.Anchor.Column);
        Assert.Equal(new double?[] { 10, 20, 30 }, Assert.Single(ChartData.Read(session.Snapshot, chart)).Values);
        Assert.All(chart.Series, s => Assert.Contains("图表数据2", s.Values));
        Assert.All(session.Snapshot.Sheets[1].Cells.Values, cell => Assert.Equal(CellContentKind.Literal, cell.Content.Kind));
        session.Undo(); Assert.Single(session.Snapshot.Sheets); Assert.Empty(session.Snapshot.Sheets[0].Data.Charts);
        session.Redo(); Assert.Equal(2, session.Snapshot.Sheets.Length);
        var data = session.Snapshot.Sheets[1]; var reference = WorksheetChart.ParseReference(chart.Series[0].Values);
        session.Execute(new SetCellsCommand([new(new(data.Id, reference.Row, reference.Column), CellContent.FromLiteral(CellValue.FromNumber(99)))]));
        Assert.Equal(99, ChartData.Read(session.Snapshot, chart)[0].Values[0]);
    }

    [Fact]
    public void PointFiltersAndRotationRemainEditable()
    {
        var (book, chart) = Create(); chart = chart with { Series = [chart.Series[0] with { HiddenPoints = [1] }], Anchor = chart.Anchor with { Rotation = 30 } };
        var copy = OfficeDrawingClipboardCodec.ReadPrivateChart(OfficeDrawingClipboardCodec.WritePrivateChart(ChartClipboard.Capture(book, chart)));
        Assert.Equal(30, copy.Chart.Anchor.Rotation); Assert.Equal(new[] { 1 }, copy.Chart.Series[0].HiddenPoints);
        Assert.Equal(new double?[] { 10, 30 }, ChartData.Read(copy.Book, copy.Chart)[0].Values);
    }
    [Fact]
    public void StaleSourceFormulaCannotBeExportedAsFreshChartData()
    {
        var (book, chart) = Create(); using var session = new DocumentSession(book);
        session.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 1, 1), CellContent.FromFormula("=10"))]));
        Assert.Throws<InvalidOperationException>(() => ChartClipboard.Capture(session.Snapshot, chart));
    }
    [Fact]
    public void ProtectedDestinationRejectsChartAndDataSheetTogether()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Target"); using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]));
        session.Execute(new SetProtectionCommand(sheet.Id, new() { Sheet = true, Objects = true })); var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new PasteChartCommand(Capture(), sheet.Id, new(0, 0)))); Assert.Same(before, session.Snapshot);
    }
    [Theory]
    [InlineData("missing-cache")] [InlineData("duplicate-point")] [InlineData("nan")] [InlineData("count")] [InlineData("stacked")] [InlineData("unknown")] [InlineData("mismatch")]
    public void InvalidOrUnmodeledChartNeverSilentlyBecomesAPicture(string mutation)
    {
        var bytes = Rewrite(NativeFixture(), "xl/charts/chart1-1.xml", xml =>
        {
            var cache = xml.Descendants(C + "numCache").First();
            switch (mutation)
            {
                case "missing-cache": cache.Remove(); break;
                case "duplicate-point": cache.Add(new XElement(cache.Element(C + "pt")!)); break;
                case "nan": cache.Descendants(C + "v").First().Value = "NaN"; break;
                case "count": cache.Element(C + "ptCount")!.SetAttributeValue("val", 1000000); break;
                case "stacked": xml.Descendants(C + "grouping").First().SetAttributeValue("val", "stacked"); break;
                case "unknown": xml.Root!.Add(new XElement(C + "unknownFeature")); break;
                case "mismatch": cache.Elements(C + "pt").Last().Remove(); cache.Element(C + "ptCount")!.SetAttributeValue("val", 2); break;
            }
        });
        Assert.ThrowsAny<Exception>(() => OfficeDrawingClipboardCodec.ReadChart(bytes));
    }
    [Fact]
    public void SourceReferencesAreReplacedWithoutOpeningExternalFiles()
    {
        var bytes = Rewrite(NativeFixture(), "xl/charts/chart1-1.xml", xml =>
        {
            foreach (var f in xml.Descendants(C + "f")) f.Value = "'https://unreachable.invalid/[source.xlsx]Sheet1'!" + f.Value.Split('!')[1];
        });
        var copied = OfficeDrawingClipboardCodec.ReadChart(bytes);
        Assert.All(copied.Chart.Series, s => Assert.DoesNotContain("https", s.Values));
        Assert.Equal(new double?[] { 10, 20, 30 }, ChartData.Read(copied.Book, copied.Chart)[0].Values);
    }
    [Fact]
    public void UnknownDrawingAttachmentIsRejected()
    {
        var bytes = Rewrite(NativeFixture(), "xl/drawings/drawing1.xml", xml => xml.Descendants(A + "cNvPr").First().Add(new XElement(A + "hlinkClick")));
        Assert.Throws<NotSupportedException>(() => OfficeDrawingClipboardCodec.ReadChart(bytes));
    }
    [Fact]
    public void ZipExpansionIsBoundedBeforeChartParsing()
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true)) using (var output = zip.CreateEntry("expanded.xml").Open()) output.Write(new byte[OfficeDrawingClipboardCodec.MaxBytes + 1]);
        Assert.Throws<InvalidDataException>(() => OfficeDrawingClipboardCodec.ReadChart(stream.ToArray()));
    }
}
