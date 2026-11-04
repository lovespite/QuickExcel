using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ConditionalVisualTests
{
    private static ConditionalCellVisual? Visual(WorkbookSnapshot book, int row, int column = 0) => ConditionalPresentation.GetVisual(book, book.Sheets[0], row, column);
    private static string? Fill(WorkbookSnapshot book, int row) => ConditionalPresentation.GetStyle(book, book.Sheets[0], row, 0).FillColor;
    private static WorkbookSnapshot With(WorksheetSnapshot sheet, params ConditionalFormatRule[] rules) => Book(sheet.WithData(sheet.Data with { ConditionalFormats = [.. rules] }));
    private static readonly CellRectangle Range = new(0, 0, 4, 0);

    [Fact]
    public void ColorScalesInterpolateTwoAndThreeStopsAndRecomputeAfterEdits()
    {
        var sheet = Sheet("Data", (0, 0, N(0)), (1, 0, N(25)), (2, 0, N(50)), (3, 0, N(75)), (4, 0, N(100)));
        var rule = ConditionalVisualPresets.ColorScale(Range, ["FF000000", "FFFFFFFF"]);
        using var session = new DocumentSession(With(sheet, rule));
        Assert.Equal("FF000000", Fill(session.Snapshot, 0)); Assert.Equal("FF808080", Fill(session.Snapshot, 2)); Assert.Equal("FFFFFFFF", Fill(session.Snapshot, 4));
        rule = ConditionalVisualPresets.ColorScale(Range, ["FFFF0000", "FFFFFF00", "FF00FF00"]);
        session.Execute(new SetConditionalFormatsCommand(sheet.Id, [rule]));
        Assert.Equal("FFFFFF00", Fill(session.Snapshot, 2)); Assert.Equal("FF80FF00", Fill(session.Snapshot, 3));
        session.Execute(new SetCellsCommand([new(new(sheet.Id, 4, 0), N(200).Content)]));
        Assert.NotEqual("FF80FF00", Fill(session.Snapshot, 3));
        session.Undo(); Assert.Equal("FF80FF00", Fill(session.Snapshot, 3));
        var higher = new ConditionalFormatRule([Range], ConditionalFormatKind.Expression, 1, new(FillColor: "FF123456"), "=TRUE");
        session.Execute(new SetConditionalFormatsCommand(sheet.Id, [higher, rule.WithPriority(2)])); Assert.Equal("FF123456", Fill(session.Snapshot, 2));
    }

    [Theory]
    [InlineData(-10, 10, 0.5, 0, 1)]
    [InlineData(5, 10, 0, 0.5, 1)]
    [InlineData(-10, -5, 1, 0, 0.5)]
    [InlineData(0, 0, 0, 0, 0)]
    public void DataBarsHaveCorrectZeroAxisAndLengths(double low, double high, double axis, double firstEnd, double lastEnd)
    {
        var sheet = Sheet("Data", (0, 0, N(low)), (1, 0, N(high)));
        var rule = ConditionalVisualPresets.DataBar(new(0, 0, 1, 0), "FF638EC6"); var book = With(sheet, rule);
        Assert.Equal(axis, Visual(book, 0)!.Bar!.Axis); Assert.Equal(firstEnd, Visual(book, 0)!.Bar!.End, 8); Assert.Equal(lastEnd, Visual(book, 1)!.Bar!.End, 8);
        Assert.Equal(low < 0 ? "FFFF0000" : "FF638EC6", Visual(book, 0)!.Bar!.Color);
        rule = new(rule.Ranges, rule.Kind, 1, new(), dataBar: rule.DataBar! with { Gradient = false, ShowValue = false, Direction = DataBarDirection.RightToLeft });
        book = With(sheet, rule); Assert.False(Visual(book, 0)!.Bar!.Gradient); Assert.True(Visual(book, 0)!.HideValue); Assert.Equal(1 - firstEnd, Visual(book, 0)!.Bar!.End, 8);
    }

    [Fact]
    public void ThresholdTypesIconsReverseBoundariesAndStopIfTrueAreRespected()
    {
        var sheet = Sheet("Data", (0, 0, N(0)), (1, 0, N(50)), (2, 0, N(100)), (0, 2, N(50)));
        ConditionalThreshold[] thresholds = [new(ConditionalThresholdKind.Number, 0), new(ConditionalThresholdKind.Formula, Formula: "=$C$1", Inclusive: false), new(ConditionalThresholdKind.Percentile, 100)];
        var rule = new ConditionalFormatRule([Range], ConditionalFormatKind.IconSet, 2, new(), iconSet: new("3Arrows", thresholds, false, true));
        var book = With(sheet, rule);
        Assert.Equal(2, Visual(book, 0)!.Icon!.Index); Assert.Equal(2, Visual(book, 1)!.Icon!.Index); Assert.Equal(0, Visual(book, 2)!.Icon!.Index); Assert.True(Visual(book, 2)!.HideValue);
        var stop = new ConditionalFormatRule([Range], ConditionalFormatKind.CellIs, 1, new(Bold: true), "=50", comparison: ConditionalComparison.Equal, stopIfTrue: true);
        book = With(sheet, stop, rule); Assert.Null(Visual(book, 1)); Assert.NotNull(Visual(book, 2));
        foreach (var name in ConditionalIconSet.Names)
        {
            book = With(sheet, ConditionalVisualPresets.IconSet(Range, name));
            Assert.Equal(0, Visual(book, 0)!.Icon!.Index); Assert.Equal(name[0] - '1', Visual(book, 2)!.Icon!.Index);
        }
    }

    [Fact]
    public void BlankTextErrorAndBooleanCellsHaveNoVisualAndWholeColumnRemainsSparse()
    {
        var sheet = Sheet("Data", (0, 0, N(5)), (2, 0, new CellSnapshot(CellContent.FromLiteral(CellValue.FromText("10")))), (3, 0, F("=1/0")), (4, 0, new CellSnapshot(CellContent.FromLiteral(CellValue.FromBoolean(true)))));
        var rule = ConditionalVisualPresets.DataBar(new(0, 0, WorkbookLimits.MaxRows - 1, 0), "FF638EC6"); var book = With(sheet, rule);
        Assert.NotNull(Visual(book, 0)); foreach (var row in new[] { 1, 2, 3, 4, WorkbookLimits.MaxRows - 1 }) Assert.Null(Visual(book, row));
        Assert.Empty(ConditionalPresentation.Diagnostics(book));
    }

    [Fact]
    public void EveryPresetAndCustomThresholdSurvivesXlsxAndXlsq()
    {
        var sheet = Sheet("Visual", (0, 0, N(-20)), (1, 0, N(20)), (0, 2, N(10)));
        var rules = new List<ConditionalFormatRule>();
        foreach (var (_, color) in ConditionalVisualPresets.BarColors)
            foreach (var gradient in new[] { true, false }) rules.Add(ConditionalVisualPresets.DataBar(Range, color, gradient));
        foreach (var (_, colors) in ConditionalVisualPresets.Scales) rules.Add(ConditionalVisualPresets.ColorScale(Range, colors));
        foreach (var name in ConditionalIconSet.Names) rules.Add(ConditionalVisualPresets.IconSet(Range, name));
        rules.Add(new([Range], ConditionalFormatKind.DataBar, 1, new(), dataBar: new(new(ConditionalThresholdKind.Number, -20), new(ConditionalThresholdKind.Formula, Formula: "=$C$1"), "FF00FF00",
            Gradient: false, ShowValue: false, MinLength: 5, MaxLength: 80, Axis: DataBarAxis.Middle, Direction: DataBarDirection.RightToLeft, BorderColor: "FF112233", NegativeBorderColor: "FF445566")));
        var book = With(sheet, rules.Select((r, i) => r.WithPriority(i + 1)).ToArray());
        foreach (var extension in new[] { ".xlsx", ".xlsq" })
        {
            var path = Path.Combine(Path.GetTempPath(), "QuickExcel-visual-" + Guid.NewGuid().ToString("N") + extension);
            try
            {
                var storage = new TypedWorkbookStorage(); storage.Save(path, book); var reopened = storage.Open(path);
                if (extension == ".xlsx")
                {
                    using var zip = ZipFile.OpenRead(path); using var reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
                    var xml = reader.ReadToEnd(); Assert.Contains("{78C0D931-6437-407d-A8EE-F0AAD7539E65}", xml); Assert.Contains("<x14:id>", xml);
                }
                Assert.Empty(reopened.FileState!.CompatibilityDiagnostics); Assert.Equal(rules.Count, reopened.Sheets[0].Data.ConditionalFormats.Length);
                Assert.All(book.Sheets[0].Data.ConditionalFormats.Zip(reopened.Sheets[0].Data.ConditionalFormats), pair => Assert.True(pair.First.ContentEquals(pair.Second), pair.First.Kind + " / " + pair.First.Priority));
            }
            finally { File.Delete(path); }
        }
    }

    [Fact]
    public void IndependentOoxmlFixturesReadBasicVisualsAndPreserveUnknownExtensions()
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-visual-fixture-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, Book(Sheet("Fixture", (0, 0, N(0)), (1, 0, N(100)))));
            void ReplaceRules(string xml)
            {
                using var zip = ZipFile.Open(path, ZipArchiveMode.Update); var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; XDocument doc;
                using (var stream = entry.Open()) doc = XDocument.Load(stream);
                doc.Root!.Elements(ns + "conditionalFormatting").Remove(); doc.Root.Add(XElement.Parse(xml)); entry.Delete();
                using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); doc.Save(output);
            }
            ReplaceRules($"""<conditionalFormatting xmlns="{ns}" sqref="A1:A2"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min" val="0"/><cfvo type="max" val="0"/><color rgb="FF3366CC"/></dataBar></cfRule><cfRule type="iconSet" priority="2"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>""");
            var opened = storage.Open(path); Assert.Empty(opened.FileState!.CompatibilityDiagnostics); Assert.Equal(.1, Visual(opened, 0)!.Bar!.End, 8); Assert.Equal(.9, Visual(opened, 1)!.Bar!.End, 8);
            ReplaceRules($"""<conditionalFormatting xmlns="{ns}" sqref="A1:A2"><cfRule type="iconSet" priority="1"><iconSet iconSet="CustomUnsupported"><cfvo type="percent" val="0"/></iconSet></cfRule></conditionalFormatting>""");
            var original = File.ReadAllBytes(path); opened = storage.Open(path); Assert.Empty(opened.Sheets[0].Data.ConditionalFormats); Assert.Contains(opened.FileState!.CompatibilityDiagnostics, d => d.Feature == "worksheet/conditionalFormatting");
            storage.Save(path, opened); Assert.Equal(original, File.ReadAllBytes(path));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void StructuralCommandsClearAndHistoryPreserveVisualMetadataAndIdentity()
    {
        var sheet = Sheet("Visual", (0, 0, N(1)), (1, 0, N(2)));
        var rule = ConditionalVisualPresets.DataBar(Range, "FF638EC6", false);
        using var session = new DocumentSession(With(sheet, rule));
        session.Execute(new ClearConditionalFormatsCommand(sheet.Id, [new(0, 0, 0, 0)]));
        var remaining = session.Snapshot.Sheets[0].Data.ConditionalFormats.Single(); Assert.Equal(rule.ScriptId, remaining.ScriptId); Assert.Equal(rule.DataBar, remaining.DataBar);
        session.Undo(); Assert.True(rule.ContentEquals(session.Snapshot.Sheets[0].Data.ConditionalFormats.Single()));
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 1, 1, false));
        Assert.Equal(rule.DataBar, session.Snapshot.Sheets[0].Data.ConditionalFormats.Single().DataBar);
    }

    [Fact]
    public void DescendingLiteralThresholdsAreRejectedBeforeTheyEnterHistory()
    {
        var sheet = Sheet("Visual", (0, 0, N(10)));
        using var session = new DocumentSession(Book(sheet));
        Assert.Throws<ArgumentException>(() => session.Execute(new SetConditionalFormatsCommand(sheet.Id, [
            new ConditionalFormatRule([Range], ConditionalFormatKind.DataBar, 1, new(), dataBar: new(new(ConditionalThresholdKind.Number, 20), new(ConditionalThresholdKind.Number, 10), "FF638EC6"))])));
        Assert.Empty(session.Snapshot.Sheets[0].Data.ConditionalFormats);
        Assert.False(session.Undo());
        Assert.Throws<ArgumentException>(() => new ConditionalColorScale([new(ConditionalThresholdKind.Percentile, 90), new(ConditionalThresholdKind.Percentile, 10)], ["FFFF0000", "FF00FF00"]));
        Assert.Throws<ArgumentException>(() => new ConditionalIconSet("3Arrows", [new(ConditionalThresholdKind.Percent, 0), new(ConditionalThresholdKind.Percent, 80), new(ConditionalThresholdKind.Percent, 20)]));
    }

    [Fact]
    public void NativeDataBarExtensionFixtureIsModeledAndUnknownPropertiesRemainPreserved()
    {
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-native-bar-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, Book(Sheet("Native", (0, 0, N(-10)), (1, 0, N(20)))));
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            void Replace(string extra)
            {
                using var zip = ZipFile.Open(path, ZipArchiveMode.Update); var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; XDocument doc;
                using (var stream = entry.Open()) doc = XDocument.Load(stream);
                doc.Root!.Elements(ns + "conditionalFormatting").Remove(); doc.Root.Elements(ns + "extLst").Remove();
                doc.Root.Add(XElement.Parse($$"""<conditionalFormatting xmlns="{{ns}}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" sqref="A1:A2"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{4CD94529-38E4-4E98-AEE5-3F97388DA51E}</x14:id></ext></extLst></cfRule></conditionalFormatting>"""));
                doc.Root.Add(XElement.Parse($$"""<extLst xmlns="{{ns}}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{4CD94529-38E4-4E98-AEE5-3F97388DA51E}"><x14:dataBar minLength="0" maxLength="100" gradient="0" negativeBarColorSameAsPositive="1" {{extra}}><x14:cfvo type="min"/><x14:cfvo type="max"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A2</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>"""));
                entry.Delete(); using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); doc.Save(output);
            }
            Replace(""); var book = storage.Open(path); var rule = Assert.Single(book.Sheets[0].Data.ConditionalFormats);
            Assert.False(rule.DataBar!.Gradient); Assert.Equal(rule.DataBar.Color, rule.DataBar.NegativeColor); Assert.Empty(book.FileState!.CompatibilityDiagnostics);
            Replace("unknownVisualSetting=\"1\""); var bytes = File.ReadAllBytes(path); book = storage.Open(path);
            Assert.Empty(book.Sheets[0].Data.ConditionalFormats); Assert.Contains(book.FileState!.CompatibilityDiagnostics, d => d.Feature == "worksheet/conditionalFormatting");
            storage.Save(path, book); Assert.Equal(bytes, File.ReadAllBytes(path));
        }
        finally { File.Delete(path); }
    }
}
