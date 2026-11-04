using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Core.Test.Fixtures;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class ChartFilteringTests : IDisposable
{
    private readonly string _folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../validation/chart-filters", Guid.NewGuid().ToString("N")));
    public ChartFilteringTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace C15 = "http://schemas.microsoft.com/office/drawing/2012/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string FilterUri = "{02D57815-91ED-43cb-92C2-25804820EDAC}";
    private static (WorkbookSnapshot Book, WorksheetChart Chart) Create(WorksheetChartKind kind = WorksheetChartKind.Column, bool multiple = false)
    {
        var cells = new List<KeyValuePair<CellPosition, CellSnapshot>>();
        void Cell(int r, int c, CellValue value) => cells.Add(new(new(r, c), new(CellContent.FromLiteral(value))));
        Cell(0, 0, CellValue.FromText("类别")); Cell(0, 1, CellValue.FromText("价格")); Cell(0, 2, CellValue.FromText("数量"));
        for (var r = 1; r <= 4; r++) { Cell(r, 0, kind == WorksheetChartKind.Scatter ? CellValue.FromNumber(r * 10) : CellValue.FromText("X" + r)); Cell(r, 1, CellValue.FromNumber(r * 20)); Cell(r, 2, CellValue.FromNumber(r * 3)); }
        var sheet = new WorksheetSnapshot(SheetId.New(), "O'Brien, Sales", cells);
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 4, multiple ? 2 : 1), kind, new() { Kind = DrawingAnchorKind.Absolute, Width = 4762500, Height = 2857500 }, "价格");
        chart = chart with { Series = [.. chart.Series.Select((s, i) => s with { HiddenPoints = [1], IsFiltered = multiple && i == 0,
            PointColors = ["FF4472C4", "FFFF0000", "FF00FF00", "FF4472C4"], TotalPoints = kind == WorksheetChartKind.Waterfall ? [2] : [] })] };
        return (new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet.WithData(sheet.Data with { Charts = [chart] })]), chart);
    }
    [Theory]
    [InlineData(WorksheetChartKind.Column)] [InlineData(WorksheetChartKind.Line)] [InlineData(WorksheetChartKind.SmoothLine)]
    [InlineData(WorksheetChartKind.Pie)] [InlineData(WorksheetChartKind.Scatter)] [InlineData(WorksheetChartKind.Waterfall)] [InlineData(WorksheetChartKind.Funnel)]
    public void FilteredChartsRoundTripAndCanRestoreEverySourcePoint(WorksheetChartKind kind)
    {
        var (book, chart) = Create(kind, kind is WorksheetChartKind.Column or WorksheetChartKind.Line or WorksheetChartKind.SmoothLine or WorksheetChartKind.Scatter);
        var storage = new TypedWorkbookStorage();
        foreach (var extension in new[] { ".xlsx", ".xlsq" })
        {
            var path = Path.Combine(_folder, "filtered" + extension); storage.Save(path, book); var read = storage.Open(path);
            Assert.Empty(read.FileState!.CompatibilityIssues); var actual = Assert.Single(read.Sheets[0].Data.Charts);
            Assert.True(chart.ContentEquals(actual), $"{kind}{extension}: {string.Join(';', read.FileState.CompatibilityIssues)}");
            var data = Assert.Single(ChartData.Read(read, actual)); Assert.Equal(new[] { 0, 2, 3 }, data.SourceIndices);
            Assert.Equal(new[] { "FF4472C4", "FF00FF00", "FF4472C4" }, data.PointColors);
            if (kind == WorksheetChartKind.Waterfall) Assert.Equal(new[] { false, true, false }, data.Totals);
            using var session = new DocumentSession(read); session.Execute(new RenameSheetCommand(read.ActiveSheetId, "Renamed"));
            var renamed = session.Snapshot.Sheets[0].Data.Charts[0];
            session.Execute(new SetCellsCommand([new(new(read.ActiveSheetId, 2, 1), CellContent.FromLiteral(CellValue.FromNumber(999)))]));
            session.Execute(new SetChartCommand(read.ActiveSheetId, 0, renamed with { Series = [.. renamed.Series.Select(s => s with { IsFiltered = false, HiddenPoints = [] })] }));
            var restoredPath = Path.Combine(_folder, "restored" + extension); storage.Save(restoredPath, session.Snapshot); var restored = storage.Open(restoredPath);
            Assert.Empty(restored.FileState!.CompatibilityIssues); var restoredChart = restored.Sheets[0].Data.Charts[0];
            Assert.All(restoredChart.Series, s => { Assert.False(s.IsFiltered); Assert.Empty(s.HiddenPoints); Assert.Contains("Renamed", s.Values); });
            Assert.Equal(999, ChartData.Read(restored, restoredChart)[0].Values[1]); Assert.Equal("FFFF0000", restoredChart.Series[0].PointColors[1]);
            if (extension == ".xlsq") { using var zip = ZipFile.OpenRead(path); using var text = new StreamReader(zip.GetEntry("manifest.json")!.Open()); Assert.Contains("chart-filters", text.ReadToEnd()); }
        }
    }
    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void EmptySelectionIsPersistedWithoutDroppingSeriesOrSource(bool hideSeries)
    {
        var (book, chart) = Create(); chart = chart with { Series = [chart.Series[0] with { IsFiltered = hideSeries, HiddenPoints = [0, 1, 2, 3] }] };
        book = book.WithSheets([book.Sheets[0].WithData(book.Sheets[0].Data with { Charts = [chart] })]);
        var storage = new TypedWorkbookStorage(); var path = Path.Combine(_folder, "empty.xlsx"); storage.Save(path, book); var read = storage.Open(path);
        Assert.Empty(read.FileState!.CompatibilityIssues); Assert.True(chart.ContentEquals(read.Sheets[0].Data.Charts[0]));
        Assert.Equal(4, ChartData.Read(read, read.Sheets[0].Data.Charts[0], includeFiltered: true)[0].Values.Length);
    }
    [Fact]
    public void SourceSelectionKeepsCompleteRangesAndDeduplicatesSharedCategories()
    {
        var (_, chart) = Create(multiple: true); var sources = ChartSelection.Sources(chart);
        Assert.Equal(5, sources.Count); Assert.Single(sources, s => s.Role == ChartSourceRole.Categories);
        Assert.All(sources, s => Assert.Equal("O'Brien, Sales", s.Sheet));
        Assert.Equal(new CellRectangle(1, 0, 4, 0), sources.First(s => s.Role == ChartSourceRole.Categories).Range);
        Assert.Equal(3, sources.Select(s => s.Color).Distinct().Count());
    }
    [Fact]
    public void ChartFilterIsAnUndoableProtectedObjectEditAndDoesNotChangeCells()
    {
        var (book, chart) = Create(); using var session = new DocumentSession(book); var cells = book.Sheets[0].Cells;
        var next = chart with { Series = [chart.Series[0] with { IsFiltered = true }] }; session.Execute(new SetChartCommand(book.ActiveSheetId, 0, next));
        Assert.True(session.IsDirty); Assert.Same(cells, session.Snapshot.Sheets[0].Cells); Assert.Empty(ChartData.Read(session.Snapshot, next));
        session.Undo(); Assert.True(chart.ContentEquals(session.Snapshot.Sheets[0].Data.Charts[0])); session.Redo(); Assert.True(session.Snapshot.Sheets[0].Data.Charts[0].Series[0].IsFiltered);
        session.Execute(new SetProtectionCommand(book.ActiveSheetId, new() { Objects = true }));
        Assert.Throws<InvalidOperationException>(() => session.Execute(new SetChartCommand(book.ActiveSheetId, 0, chart)));
    }
    [Theory]
    [InlineData(-1)] [InlineData(4)]
    public void InvalidSourceIndexesAreRejected(int index)
    { var (_, chart) = Create(); Assert.Throws<ArgumentException>(() => (chart with { Series = [chart.Series[0] with { HiddenPoints = [index] }] }).Validate()); }

    [Theory]
    [InlineData(WorksheetChartKind.Waterfall, false)] [InlineData(WorksheetChartKind.Waterfall, true)]
    [InlineData(WorksheetChartKind.Funnel, false)] [InlineData(WorksheetChartKind.Funnel, true)]
    public void ExtendedFiltersRestoreEmptyCategoriesAndHiddenSeries(WorksheetChartKind kind, bool hidden)
    {
        var (book, chart) = Create(kind); chart = chart with { Series = [chart.Series[0] with { IsFiltered = hidden, HiddenPoints = [0, 1, 2, 3] }] };
        book = book.WithSheets([book.Sheets[0].WithData(book.Sheets[0].Data with { Charts = [chart] })]);
        var path = Path.Combine(_folder, "extended.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, book); var read = storage.Open(path);
        Assert.Empty(read.FileState!.CompatibilityIssues); Assert.True(chart.ContentEquals(read.Sheets[0].Data.Charts[0]));
    }

    [Theory]
    [InlineData(WorksheetChartKind.Waterfall)] [InlineData(WorksheetChartKind.Funnel)]
    public void EveryXlsxWritePathUsesExcelCompatibleChartExAndRetainsFilters(WorksheetChartKind kind)
    {
        XNamespace cx = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        var (book, chart) = Create(kind); var storage = new TypedWorkbookStorage();
        foreach (var state in new[] { "all", "filtered", "empty", "hidden" })
        foreach (var mode in new[] { "save", "values", "portable" })
        {
            var expected = chart with { Series = [chart.Series[0] with { IsFiltered = state == "hidden", HiddenPoints = state switch { "all" => [], "empty" => [0, 1, 2, 3], _ => [1] } }] };
            var source = book.WithSheets([book.Sheets[0].WithData(book.Sheets[0].Data with { Charts = [expected] })]);
            var path = Path.Combine(_folder, $"{state}-{mode}.xlsx");
            if (mode == "save") storage.SaveChecked(path, source, false, new(null));
            else if (mode == "values") storage.Export(path, source);
            else storage.ExportPortable(path, source);
            var read = storage.Open(path); Assert.Empty(read.FileState!.CompatibilityIssues);
            Assert.True(expected.ContentEquals(Assert.Single(read.Sheets[0].Data.Charts)));
            using var zip = ZipFile.OpenRead(path);
            using var xml = zip.GetEntry("xl/charts/chart1-1.xml")!.Open(); Assert.Empty(XDocument.Load(xml).Descendants(cx + "axisId"));
            using var relationships = zip.GetEntry("xl/charts/_rels/chart1-1.xml.rels")!.Open();
            var types = XDocument.Load(relationships).Root!.Elements(rel + "Relationship").Select(r => (string)r.Attribute("Type")!).ToArray();
            Assert.Contains("http://schemas.microsoft.com/office/2011/relationships/chartStyle", types);
            Assert.Contains("http://schemas.microsoft.com/office/2011/relationships/chartColorStyle", types);
        }
    }

    [Fact]
    public void LegacyFilteredChartExRepairUpdatesVisibleAndFullSeriesTogether()
    {
        var (book, chart) = Create(WorksheetChartKind.Waterfall); var storage = new TypedWorkbookStorage();
        var path = Path.Combine(_folder, "legacy-filtered.xlsx"); storage.Save(path, book);
        XNamespace cx = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/charts/chart1-1.xml")!; XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            foreach (var series in xml.Descendants(cx + "series")) series.Add(new XElement(cx + "axisId", 0), new XElement(cx + "axisId", 1));
            entry.Delete(); using var output = zip.CreateEntry("xl/charts/chart1-1.xml").Open(); xml.Save(output);
        }
        var read = storage.Open(path); Assert.Empty(read.FileState!.CompatibilityIssues);
        var repaired = Path.Combine(_folder, "filtered-repaired.xlsx"); storage.Save(repaired, read);
        Assert.True(chart.ContentEquals(Assert.Single(storage.Open(repaired).Sheets[0].Data.Charts)));
        using var saved = ZipFile.OpenRead(repaired); using var part = saved.GetEntry("xl/charts/chart1-1.xml")!.Open();
        Assert.Empty(XDocument.Load(part).Descendants(cx + "axisId"));
    }

    [Fact]
    public void LegacyRepairKeepsDistinctStylesForIdenticalChartXml()
    {
        var (book, chart) = Create(WorksheetChartKind.Waterfall); var second = chart with { IncreaseColor = "FFFF0000", DecreaseColor = "FF00FF00" };
        book = book.WithSheets([book.Sheets[0].WithData(book.Sheets[0].Data with { Charts = [chart, second] })]);
        var storage = new TypedWorkbookStorage(); var path = Path.Combine(_folder, "two-legacy.xlsx"); storage.Save(path, book);
        XNamespace cx = "http://schemas.microsoft.com/office/drawing/2014/chartex";
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        foreach (var name in new[] { "xl/charts/chart1-1.xml", "xl/charts/chart1-2.xml" })
        {
            var entry = zip.GetEntry(name)!; XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            foreach (var series in xml.Descendants(cx + "series")) series.Add(new XElement(cx + "axisId", 0));
            entry.Delete(); using var output = zip.CreateEntry(name).Open(); xml.Save(output);
        }
        var read = storage.Open(path); Assert.Equal(2, read.Sheets[0].Data.Charts.Length);
        Assert.True(read.Sheets[0].Data.Charts[0].OriginalXml.SequenceEqual(read.Sheets[0].Data.Charts[1].OriginalXml));
        var repaired = Path.Combine(_folder, "two-repaired.xlsx"); storage.Save(repaired, read);
        var actual = storage.Open(repaired).Sheets[0].Data.Charts;
        Assert.True(chart.ContentEquals(actual[0])); Assert.True(second.ContentEquals(actual[1]));
    }
    [Fact]
    public void HorizontalSourcesRetainQuotedSheetNameAndHiddenPointIndexes()
    {
        var (book, chart) = Create(); chart = chart with { Series = [new("横向", null, "'O''Brien, Sales'!$A$2:$D$2", "'O''Brien, Sales'!$A$3:$D$3") { HiddenPoints = [0, 2] }] };
        book = book.WithSheets([book.Sheets[0].WithData(book.Sheets[0].Data with { Charts = [chart] })]);
        var path = Path.Combine(_folder, "horizontal.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, book); var read = storage.Open(path);
        Assert.Empty(read.FileState!.CompatibilityIssues); Assert.True(chart.ContentEquals(read.Sheets[0].Data.Charts[0]));
        Assert.Equal(new[] { 1, 3 }, ChartData.Read(read, read.Sheets[0].Data.Charts[0])[0].SourceIndices);
    }
    [Fact]
    public void DifferentSeriesFiltersLeaveGapsOnTheSharedCategoryAxis()
    {
        var (book, chart) = Create(multiple: true);
        chart = chart with { Series = [chart.Series[0] with { IsFiltered = false, HiddenPoints = [1] }, chart.Series[1] with { HiddenPoints = [] }] };
        var data = ChartData.Read(book, chart);
        Assert.Equal(data[1].Categories, data[0].Categories); Assert.Equal(new[] { 0, 1, 2, 3 }, data[0].SourceIndices);
        Assert.Equal(new double?[] { 20, null, 60, 80 }, data[0].Values); Assert.Equal(new double?[] { 3, 6, 9, 12 }, data[1].Values);
    }

    // Independent OOXML fixture modeled after Office's filteredBarSeries/fullRef/formulaRef encoding.
    // It never uses QuickExcel's chart writer to create the input.
    private static byte[] OfficeFixture(string mode = "valid")
    {
        const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string pkg = "http://schemas.openxmlformats.org/package/2006/relationships";
        string Full(string column) => $"Sheet1!${column}$2:${column}$5";
        string Visible(string column) => mode == "outside" && column == "B" ? "Sheet1!$B$6" : $"(Sheet1!${column}$2,Sheet1!${column}$4:${column}$5)";
        string Ref(string column, bool hidden, bool numeric) => $"""
            <c:{(numeric ? "numRef" : "strRef")}><c:extLst><c:ext uri="{FilterUri}"><c15:fullRef><c15:sqref>{Full(column)}</c15:sqref></c15:fullRef>{(hidden ? $"<c15:formulaRef><c15:sqref>{Visible(column)}</c15:sqref></c15:formulaRef>" : "")}</c:ext></c:extLst>{(hidden ? "" : $"<c:f>{Visible(column)}</c:f>")}</c:{(numeric ? "numRef" : "strRef")}>
            """;
        string Series(bool hidden) => $"""
            <{(hidden ? "c15" : "c")}:ser><c:idx val="{(hidden ? 0 : 1)}"/><c:order val="{(hidden ? 0 : 1)}"/><c:tx><c:v>{(hidden ? "Hidden" : "Visible")}</c:v></c:tx>
            <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></c:spPr></c:dPt>
            <c:cat>{Ref("A", hidden, false)}</c:cat><c:val>{Ref("B", hidden, true)}</c:val>
            <c:extLst><c:ext uri="{FilterUri}"><c15:categoryFilterExceptions><c15:categoryFilterException><c15:sqref>Sheet1!$B$3</c15:sqref><c15:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>{(mode == "effect" ? "<a:effectLst><a:glow rad=\"100\"/></a:effectLst>" : "")}</c15:spPr></c15:categoryFilterException></c15:categoryFilterExceptions></c:ext></c:extLst>
            </{(hidden ? "c15" : "c")}:ser>
            """;
        var chart = $"""
            <c:chartSpace xmlns:c="{C}" xmlns:c15="{C15}" xmlns:a="{A}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{Series(false)}<c:axId val="1"/><c:axId val="2"/>
            <c:extLst><c:ext uri="{FilterUri}"><c15:filteredBarSeries>{Series(true)}</c15:filteredBarSeries>{(mode == "unknown" ? "<c15:unknown/>" : "")}</c:ext>{(mode == "empty-unknown" ? "<c:ext uri=\"unknown\"/>" : "")}</c:extLst></c:barChart>
            <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
            </c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>
            """;
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
                <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>
                """,
            ["xl/worksheets/sheet1.xml"] = $"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"{r}\"><sheetData>{string.Concat(Enumerable.Range(2, 4).Select(i => $"<row r=\"{i}\"><c r=\"A{i}\"><v>{i}</v></c><c r=\"B{i}\"><v>{i * 10}</v></c></row>"))}</sheetData><drawing r:id=\"d\"/></worksheet>",
            ["xl/worksheets/_rels/sheet1.xml.rels"] = $"<Relationships xmlns=\"{pkg}\"><Relationship Id=\"d\" Type=\"{r}/drawing\" Target=\"../drawings/drawing1.xml\"/></Relationships>",
            ["xl/drawings/drawing1.xml"] = $"<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"{A}\" xmlns:c=\"{C}\" xmlns:r=\"{r}\"><xdr:absoluteAnchor><xdr:pos x=\"0\" y=\"0\"/><xdr:ext cx=\"4762500\" cy=\"2857500\"/><xdr:graphicFrame><a:graphic><a:graphicData uri=\"{C}\"><c:chart r:id=\"c\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>",
            ["xl/drawings/_rels/drawing1.xml.rels"] = $"<Relationships xmlns=\"{pkg}\"><Relationship Id=\"c\" Type=\"{r}/chart\" Target=\"../charts/chart1.xml\"/></Relationships>",
            ["xl/charts/chart1.xml"] = chart
        };
        return WorkbookPackageFixture.Xlsx("", parts);
    }
    [Fact]
    public void ImportsOfficeFilteredSeriesReferencesAndHiddenPointStyles()
    {
        var path = Path.Combine(_folder, "office.xlsx"); File.WriteAllBytes(path, OfficeFixture()); var storage = new TypedWorkbookStorage(); var book = storage.Open(path);
        Assert.Empty(book.FileState!.CompatibilityIssues); var chart = Assert.Single(book.Sheets[0].Data.Charts);
        Assert.True(chart.Series[0].IsFiltered); Assert.False(chart.Series[1].IsFiltered);
        Assert.All(chart.Series, s => { Assert.Equal(new[] { 1 }, s.HiddenPoints); Assert.Equal("FFFF0000", s.PointColors[1]); Assert.Equal("FF00FF00", s.PointColors[2]); });
        Assert.Equal(new double?[] { 20, 40, 50 }, Assert.Single(ChartData.Read(book, chart)).Values);
        var output = Path.Combine(_folder, "office-saved.xlsx"); storage.Save(output, book); var read = storage.Open(output);
        Assert.Empty(read.FileState!.CompatibilityIssues); Assert.True(chart.ContentEquals(read.Sheets[0].Data.Charts[0]));
    }
    [Theory]
    [InlineData("outside")] [InlineData("effect")] [InlineData("unknown")] [InlineData("empty-unknown")]
    public void UnmodeledOrMalformedFiltersKeepOriginalPackageProtection(string mode)
    {
        var input = OfficeFixture(mode); var path = Path.Combine(_folder, "bad.xlsx"); File.WriteAllBytes(path, input); var storage = new TypedWorkbookStorage(); var book = storage.Open(path);
        Assert.Empty(book.Sheets[0].Data.Charts); Assert.Contains(book.FileState!.CompatibilityDiagnostics, d => d.Feature == "chart/unsupported");
        var output = Path.Combine(_folder, "copy.xlsx"); storage.Save(output, book); Assert.Equal(input, File.ReadAllBytes(output));
    }
}
