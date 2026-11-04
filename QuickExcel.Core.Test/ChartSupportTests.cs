using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Core.Test.Fixtures;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class ChartSupportTests : IDisposable
{
    private readonly string _folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../validation/chart-support", Guid.NewGuid().ToString("N")));
    public ChartSupportTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Cx = "http://schemas.microsoft.com/office/drawing/2014/chartex";
    private static WorkbookSnapshot Book()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Sheet1", new[] { "X", "Sales" }.Select((v, c) => new KeyValuePair<CellPosition, CellSnapshot>(new(0, c), new(CellContent.FromLiteral(CellValue.FromText(v)))))
            .Concat(new[] { (1d, 10d), (2d, 30d), (100d, 20d) }.SelectMany((v, i) => new[] {
                new KeyValuePair<CellPosition, CellSnapshot>(new(i + 1, 0), new(CellContent.FromLiteral(CellValue.FromNumber(v.Item1)))),
                new KeyValuePair<CellPosition, CellSnapshot>(new(i + 1, 1), new(CellContent.FromLiteral(CellValue.FromNumber(v.Item2)))) })));
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
    }
    [Theory]
    [InlineData(WorksheetChartKind.Column)] [InlineData(WorksheetChartKind.Line)] [InlineData(WorksheetChartKind.SmoothLine)]
    [InlineData(WorksheetChartKind.Scatter)] [InlineData(WorksheetChartKind.Waterfall)] [InlineData(WorksheetChartKind.Funnel)]
    public void NewTypesPreserveDataAndRebuildCachesAfterEdits(WorksheetChartKind kind)
    {
        using var session = new DocumentSession(Book()); var sheet = session.Snapshot.Sheets[0];
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 3, 1), kind, DrawingGeometry.FromBounds(sheet, new(20, 140, 500, 300), DrawingAnchorKind.OneCell), "Sales");
        if (kind == WorksheetChartKind.Waterfall) chart = chart with { Series = [chart.Series[0] with { TotalPoints = [2] }] };
        if (kind == WorksheetChartKind.Line) chart = chart with { Series = [chart.Series[0] with { Marker = "diamond", MarkerSize = 7 }] };
        session.Execute(new SetChartCommand(sheet.Id, 0, chart));
        var storage = new TypedWorkbookStorage();
        foreach (var ext in new[] { ".xlsx", ".xlsq" })
        {
            var path = Path.Combine(_folder, kind + ext); storage.Save(path, session.Snapshot);
            var read = storage.Open(path); Assert.Empty(read.FileState!.CompatibilityIssues);
            var actual = Assert.Single(read.Sheets[0].Data.Charts); Assert.True(chart.ContentEquals(actual), $"{kind}: {chart} vs {actual}");
            Assert.Equal(new double?[] { 1, 2, 100 }, ChartData.Read(read, actual)[0].XValues);
            using var edit = new DocumentSession(read); edit.Execute(new RenameSheetCommand(read.Sheets[0].Id, "O'Brien"));
            edit.Execute(new SetCellsCommand([new(new(read.Sheets[0].Id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(77)))]));
            var edited = Path.Combine(_folder, "edited" + kind + ext); storage.Save(edited, edit.Snapshot);
            var reopened = storage.Open(edited); Assert.Empty(reopened.FileState!.CompatibilityIssues);
            Assert.Equal(77, ChartData.Read(reopened, reopened.Sheets[0].Data.Charts[0])[0].Values[0]);
            Assert.Contains("O''Brien", reopened.Sheets[0].Data.Charts[0].Series[0].Values);
        }
    }

    [Fact]
    public void OfficeStylesAndLayoutSurviveSourceEditsWithoutLosingThemeNodes()
    {
        var path = Path.Combine(_folder, "styled.xlsx"); File.WriteAllBytes(path, Fixture());
        var storage = new TypedWorkbookStorage(); var book = storage.Open(path); Assert.Empty(book.FileState!.CompatibilityIssues);
        var chart = Assert.Single(book.Sheets[0].Data.Charts);
        Assert.Equal(219, chart.GapWidth); Assert.Equal(-27, chart.Overlap); Assert.Equal(2, chart.StyleParts.Length);
        Assert.Equal("FFFF0000", chart.Series[0].Color); Assert.Equal(14, chart.Appearance.Title.Size);
        using var edit = new DocumentSession(book); edit.Execute(new SetCellsCommand([new(new(book.Sheets[0].Id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(77)))]));
        var output = Path.Combine(_folder, "styled-edited.xlsx"); storage.Save(output, edit.Snapshot);
        var read = storage.Open(output); Assert.Empty(read.FileState!.CompatibilityIssues); var actual = Assert.Single(read.Sheets[0].Data.Charts);
        Assert.True(chart.ContentEquals(actual)); Assert.Equal(77, ChartData.Read(read, actual)[0].Values[0]);
        Assert.True(chart.StyleParts[0].Bytes.SequenceEqual(actual.StyleParts[0].Bytes));
        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(actual.OriginalXml.AsSpan()));
        Assert.Contains(xml.Descendants(A + "schemeClr"), e => (string?)e.Attribute("val") == "tx1");
        Assert.Contains(xml.Descendants(C + "v"), e => e.Value == "77");
    }

    [Theory]
    [InlineData("external")] [InlineData("unknown")] [InlineData("dependency")] [InlineData("effect")]
    public void UnmodeledStyleDependenciesRemainOriginalOnly(string mode)
    {
        var bytes = Fixture(mode); var path = Path.Combine(_folder, mode + ".xlsx"); File.WriteAllBytes(path, bytes);
        var storage = new TypedWorkbookStorage(); var read = storage.Open(path);
        Assert.Empty(read.Sheets[0].Data.Charts); Assert.Contains(read.FileState!.CompatibilityDiagnostics, d => d.Feature == "chart/unsupported");
        var output = Path.Combine(_folder, mode + "-copy.xlsx"); storage.Save(output, read); Assert.Equal(bytes, File.ReadAllBytes(output));
    }

    [Fact]
    public void WaterfallUsesSubtotalsAndOriginalPointIndexesAfterFiltering()
    {
        var data = new ChartSeriesData("", ["A", "B", "Total", "C"], [100, -30, 70, 20], "FF000000", []) { Totals = [false, false, true, false] };
        Assert.Equal(new[] { (0d, 100d), (100d, 70d), (0d, 70d), (70d, 90d) }, ChartData.Waterfall(data));
        var book = Book(); var sheet = book.Sheets[0]; var chart = SetChartCommand.FromRange(sheet, new(0, 0, 3, 1), WorksheetChartKind.Waterfall, new() { Kind = DrawingAnchorKind.Absolute, Width = 952500, Height = 952500 });
        chart = chart with { Series = [chart.Series[0] with { TotalPoints = [2] }] };
        using var session = new DocumentSession(book); session.Execute(new SetAxisMetadataCommand(sheet.Id, SheetAxis.Row, 2, 1, hidden: true));
        var filtered = ChartData.Read(session.Snapshot, chart)[0]; Assert.Equal(new[] { false, true }, filtered.Totals);
    }

    [Fact]
    public void PartialWaterfallPointColorsKeepAutomaticColorsForOtherPoints()
    {
        var book = Book(); var sheet = book.Sheets[0];
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 3, 1), WorksheetChartKind.Waterfall, new() { Kind = DrawingAnchorKind.Absolute, Width = 952500, Height = 952500 });
        chart = chart with { Series = [chart.Series[0] with { PointColors = ["FFFF00FF", "", ""], TotalPoints = [2] }] };
        using var session = new DocumentSession(book); session.Execute(new SetChartCommand(sheet.Id, 0, chart));
        var path = Path.Combine(_folder, "partial-colors.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot);
        var read = storage.Open(path); var actual = Assert.Single(read.Sheets[0].Data.Charts); Assert.True(chart.ContentEquals(actual));
        using var edit = new DocumentSession(read); edit.Execute(new SetCellsCommand([new(new(read.Sheets[0].Id, 2, 1), CellContent.FromLiteral(CellValue.FromNumber(-30)))]));
        storage.Save(Path.Combine(_folder, "partial-colors-edited.xlsx"), edit.Snapshot);
    }

    [Fact]
    public void ChangingStyledChartTypePreservesCommonAppearanceAndCanBeSaved()
    {
        var input = Path.Combine(_folder, "styled-switch.xlsx"); File.WriteAllBytes(input, Fixture()); var storage = new TypedWorkbookStorage();
        var read = storage.Open(input); var original = read.Sheets[0].Data.Charts[0] with { Title = "Styled", Appearance = read.Sheets[0].Data.Charts[0].Appearance with { Axis = new("FF123456", "Arial", 11), Legend = new("FF345678", "Arial", 10) } };
        foreach (var kind in Enum.GetValues<WorksheetChartKind>().Where(k => k != original.Kind))
        {
            var converted = SetChartCommand.ChangeKind(original, kind); using var session = new DocumentSession(read);
            session.Execute(new SetChartCommand(read.Sheets[0].Id, 0, converted)); var output = Path.Combine(_folder, "converted-" + kind + ".xlsx"); storage.Save(output, session.Snapshot);
            var actual = storage.Open(output).Sheets[0].Data.Charts[0]; Assert.True(converted.ContentEquals(actual), kind.ToString());
            Assert.Equal(original.Appearance.Background, actual.Appearance.Background); Assert.Equal(original.Appearance.Title, actual.Appearance.Title);
        }
    }

    [Fact]
    public void ScatterCanMixConnectedAndDisconnectedSeries()
    {
        var book = Book(); var sheet = book.Sheets[0]; var chart = SetChartCommand.FromRange(sheet, new(0, 0, 3, 1), WorksheetChartKind.Scatter, new() { Kind = DrawingAnchorKind.Absolute, Width = 952500, Height = 952500 });
        chart = chart with { Series = [chart.Series[0], chart.Series[0] with { Color = "FFFF00FF", ConnectPoints = true, Smooth = true }] };
        using var session = new DocumentSession(book); session.Execute(new SetChartCommand(sheet.Id, 0, chart));
        var path = Path.Combine(_folder, "mixed-scatter.xlsx"); var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot);
        Assert.True(chart.ContentEquals(storage.Open(path).Sheets[0].Data.Charts[0]));
    }

    [Theory]
    [InlineData("waterfall", WorksheetChartKind.Waterfall)] [InlineData("funnel", WorksheetChartKind.Funnel)]
    public void IndependentChartExFixtureImportsReferencesAndSubtotals(string layout, WorksheetChartKind kind)
    {
        var xml = $"""
            <cx:chartSpace xmlns:cx="{Cx}" xmlns:a="{A}"><cx:chartData><cx:data id="9">
            <cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$3</cx:f><cx:lvl ptCount="2"><cx:pt idx="0">old</cx:pt><cx:pt idx="1">cache</cx:pt></cx:lvl></cx:strDim>
            <cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$3</cx:f><cx:lvl ptCount="2" formatCode="General"><cx:pt idx="0">9999</cx:pt></cx:lvl></cx:numDim>
            </cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="{layout}"><cx:tx><cx:txData><cx:v>Independent</cx:v></cx:txData></cx:tx>
            <cx:dataLabels><cx:visibility value="1"/></cx:dataLabels><cx:dataId val="9"/>{(layout == "waterfall" ? "<cx:layoutPr><cx:visibility connectorLines=\"1\"/><cx:subtotals><cx:idx val=\"1\"/></cx:subtotals></cx:layoutPr>" : "")}
            </cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>
            """;
        var bytes = ReplaceChart(Fixture(), xml, true); var path = Path.Combine(_folder, layout + "-independent.xlsx"); File.WriteAllBytes(path, bytes);
        var storage = new TypedWorkbookStorage(); var read = storage.Open(path); Assert.Empty(read.FileState!.CompatibilityIssues);
        var chart = Assert.Single(read.Sheets[0].Data.Charts); Assert.Equal(kind, chart.Kind); Assert.Equal(new double?[] { 10, 20 }, ChartData.Read(read, chart)[0].Values);
        if (kind == WorksheetChartKind.Waterfall) Assert.Equal(new[] { 1 }, chart.Series[0].TotalPoints);
        using var edit = new DocumentSession(read); edit.Execute(new SetCellsCommand([new(new(read.Sheets[0].Id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(45)))]));
        var output = Path.Combine(_folder, layout + "-edited.xlsx"); storage.Save(output, edit.Snapshot); Assert.Equal(45, ChartData.Read(storage.Open(output), chart)[0].Values[0]);
    }

    [Fact]
    public void IndependentSmoothScatterUsesXYReferencesAndKeepsMarkerAndAxisLimits()
    {
        var xml = $"""
            <c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:plotArea><c:scatterChart><c:scatterStyle val="smoothMarker"/>
            <c:ser><c:idx val="0"/><c:order val="0"/><c:marker><c:symbol val="diamond"/><c:size val="7"/></c:marker>
            <c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$3</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:yVal><c:smooth val="1"/></c:ser><c:axId val="10"/><c:axId val="20"/></c:scatterChart>
            <c:valAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/><c:max val="200"/><c:min val="0"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/></c:valAx>
            <c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """;
        var path = Path.Combine(_folder, "scatter.xlsx"); File.WriteAllBytes(path, ReplaceChart(Fixture(), xml, false));
        var read = new TypedWorkbookStorage().Open(path); Assert.Empty(read.FileState!.CompatibilityIssues); var chart = Assert.Single(read.Sheets[0].Data.Charts);
        Assert.Equal(WorksheetChartKind.Scatter, chart.Kind); Assert.Equal(200, chart.XMaximum); Assert.True(chart.Series[0].Smooth); Assert.Equal("diamond", chart.Series[0].Marker);
        Assert.Equal(new double?[] { 1, 2 }, ChartData.Read(read, chart)[0].XValues);
    }

    private static byte[] ReplaceChart(byte[] source, string xml, bool extended)
    {
        using var output = new MemoryStream(); output.Write(source); output.Position = 0;
        using (var zip = new ZipArchive(output, ZipArchiveMode.Update, true))
            foreach (var entry in zip.Entries.ToArray())
            {
                string content; using (var reader = new StreamReader(entry.Open())) content = reader.ReadToEnd();
                if (entry.FullName == "xl/charts/chart1.xml") content = xml;
                if (extended && entry.FullName == "[Content_Types].xml") content = content.Replace("application/vnd.openxmlformats-officedocument.drawingml.chart+xml", "application/vnd.ms-office.chartex+xml");
                if (extended && entry.FullName == "xl/drawings/drawing1.xml") content = content.Replace(C.ToString(), Cx.ToString());
                if (extended && entry.FullName == "xl/drawings/_rels/drawing1.xml.rels") content = content.Replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "http://schemas.microsoft.com/office/2014/relationships/chartEx");
                var name = entry.FullName; entry.Delete(); using var writer = new StreamWriter(zip.CreateEntry(name).Open()); writer.Write(content);
            }
        return output.ToArray();
    }

    [Theory]
    [InlineData("waterfall", false)] [InlineData("funnel", false)]
    [InlineData("waterfall", true)] [InlineData("funnel", true)]
    public void UnchangedLegacyChartExSaveRepairsExcelDependenciesWithoutRebuildingOtherParts(string layout, bool partialStyles)
    {
        // Independent legacy package: explicit axisId values and absent/incomplete style
        // dependencies are known to make Excel reject these otherwise readable files.
        var xml = $"""
            <cx:chartSpace xmlns:cx="{Cx}"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$3</cx:f></cx:strDim>
            <cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$3</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>
            <cx:series layoutId="{layout}"><cx:dataId val="0"/><cx:axisId>0</cx:axisId><cx:axisId>1</cx:axisId></cx:series>
            </cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>
            """;
        using var stream = new MemoryStream(); stream.Write(ReplaceChart(Fixture(), xml, true)); stream.Position = 0;
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true))
        {
            if (!partialStyles)
            {
                foreach (var path in new[] { "xl/charts/style1.xml", "xl/charts/colors1.xml", "xl/charts/_rels/chart1.xml.rels" }) zip.GetEntry(path)!.Delete();
                var entry = zip.GetEntry("[Content_Types].xml")!; XDocument types;
                using (var input = entry.Open()) types = XDocument.Load(input);
                types.Root!.Elements().Where(e => ((string?)e.Attribute("PartName")) is "/xl/charts/style1.xml" or "/xl/charts/colors1.xml").Remove();
                entry.Delete(); using var output = zip.CreateEntry("[Content_Types].xml").Open(); types.Save(output);
            }
            using var opaque = new StreamWriter(zip.CreateEntry("custom/opaque.xml").Open()); opaque.Write("<keep value=\"untouched\"/>");
        }
        var inputPath = Path.Combine(_folder, "legacy.xlsx"); File.WriteAllBytes(inputPath, stream.ToArray());
        var storage = new TypedWorkbookStorage(); var read = storage.Open(inputPath); var chart = Assert.Single(read.Sheets[0].Data.Charts);
        Assert.Contains(TypedWorkbookStorage.InspectSavePlan(read), p => p.Action == "repair-excel-compatibility");
        var outputPath = Path.Combine(_folder, "repaired.xlsx"); storage.SaveChecked(outputPath, read, false, new(null));
        var reopened = storage.Open(outputPath); Assert.True(chart.ContentEquals(Assert.Single(reopened.Sheets[0].Data.Charts)));
        using (var before = ZipFile.OpenRead(inputPath))
        using (var after = ZipFile.OpenRead(outputPath))
        {
            foreach (var entry in before.Entries.Where(e => e.FullName != "[Content_Types].xml" && !e.FullName.StartsWith("xl/charts/", StringComparison.Ordinal)))
            {
                using var original = new StreamReader(entry.Open()); using var saved = new StreamReader(after.GetEntry(entry.FullName)!.Open());
                Assert.Equal(original.ReadToEnd(), saved.ReadToEnd());
            }
            using var chartXml = after.GetEntry("xl/charts/chart1.xml")!.Open(); Assert.Empty(XDocument.Load(chartXml).Descendants(Cx + "axisId"));
        }
        var copy = Path.Combine(_folder, "already-compatible.xlsx"); storage.Save(copy, reopened);
        Assert.Equal(File.ReadAllBytes(outputPath), File.ReadAllBytes(copy)); // Idempotent original-package save.
    }

    private static byte[] Fixture(string mode = "")
    {
        const string rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string cs = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
        var types = """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
            <Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/><Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/></Types>
            """;
        var chart = $"""
            <c:chartSpace xmlns:c="{C}" xmlns:a="{A}"><c:chart><c:title><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr></a:p></c:txPr></c:title><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
            <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val></c:ser><c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="1"/><c:axId val="2"/></c:barChart>
            <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
            </c:plotArea><c:plotVisOnly val="1"/></c:chart><c:spPr><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:effectLst>{(mode == "effect" ? "<a:glow rad=\"100\"/>" : "")}</a:effectLst></c:spPr></c:chartSpace>
            """;
        var styleRelation = $"<Relationship Id=\"style\" Type=\"http://schemas.microsoft.com/office/2011/relationships/{(mode == "unknown" ? "unknownStyle" : "chartStyle")}\" Target=\"{(mode == "external" ? "https://example.invalid/style.xml" : "style1.xml")}\"{(mode == "external" ? " TargetMode=\"External\"" : "")}/>";
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = types,
            ["xl/worksheets/sheet1.xml"] = $"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"{rel}\"><sheetData><row r=\"2\"><c r=\"A2\"><v>1</v></c><c r=\"B2\"><v>10</v></c></row><row r=\"3\"><c r=\"A3\"><v>2</v></c><c r=\"B3\"><v>20</v></c></row></sheetData><drawing r:id=\"drawing\"/></worksheet>",
            ["xl/worksheets/_rels/sheet1.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"drawing\" Type=\"{rel}/drawing\" Target=\"../drawings/drawing1.xml\"/></Relationships>",
            ["xl/drawings/drawing1.xml"] = $"<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"{A}\" xmlns:c=\"{C}\" xmlns:r=\"{rel}\"><xdr:absoluteAnchor><xdr:pos x=\"0\" y=\"0\"/><xdr:ext cx=\"4762500\" cy=\"2857500\"/><xdr:graphicFrame><a:graphic><a:graphicData uri=\"{C}\"><c:chart r:id=\"chart\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>",
            ["xl/drawings/_rels/drawing1.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"chart\" Type=\"{rel}/chart\" Target=\"../charts/chart1.xml\"/></Relationships>",
            ["xl/charts/chart1.xml"] = chart,
            ["xl/charts/_rels/chart1.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{styleRelation}<Relationship Id=\"colors\" Type=\"http://schemas.microsoft.com/office/2011/relationships/chartColorStyle\" Target=\"colors1.xml\"/></Relationships>",
            ["xl/charts/style1.xml"] = $"<cs:chartStyle xmlns:cs=\"{cs}\" id=\"201\"><cs:title><cs:defRPr sz=\"1400\"/></cs:title></cs:chartStyle>",
            ["xl/charts/colors1.xml"] = $"<cs:colorStyle xmlns:cs=\"{cs}\" xmlns:a=\"{A}\" meth=\"cycle\" id=\"10\"><a:srgbClr val=\"FF0000\"/><a:srgbClr val=\"00FF00\"/><cs:variation/></cs:colorStyle>"
        };
        if (mode == "dependency") parts["xl/charts/_rels/style1.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"link\" Type=\"{rel}/hyperlink\" Target=\"https://example.invalid\" TargetMode=\"External\"/></Relationships>";
        return WorkbookPackageFixture.Xlsx("", parts);
    }
}
