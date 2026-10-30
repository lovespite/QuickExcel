using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class TextFunctionTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-text-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Compatibility = "http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion";
    public TextFunctionTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private string PathFor(string name) => Path.Combine(_folder, name);
    private static byte[] Part(string path, string part)
    { using var zip = ZipFile.OpenRead(path); using var stream = zip.GetEntry(part)!.Open(); using var memory = new MemoryStream(); stream.CopyTo(memory); return memory.ToArray(); }
    private static void WritePart(string path, string part, byte[] bytes)
    { using var zip = ZipFile.Open(path, ZipArchiveMode.Update); zip.GetEntry(part)?.Delete(); using var stream = zip.CreateEntry(part).Open(); stream.Write(bytes); }
    private static WorkbookSnapshot UnicodeBook(int version) => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [Sheet("S", (0, 0, T("😀a😀")), (0, 1, F("=LEN(A1)")), (1, 1, F("=MID(A1,2,2)")),
            (2, 1, F("=REPLACE(A1,2,1,\"Z\")")), (3, 1, F("=LEFT(A1,1)")), (4, 1, F("=RIGHT(A1,1)")),
            (5, 1, F("=FIND(\"a\",A1)")), (6, 1, F("=SEARCH(\"a\",A1)")))], textCompatibilityVersion: version);

    [Theory]
    [InlineData(1900, "xlsx")]
    [InlineData(1904, "xlsx")]
    [InlineData(1900, "xlsq")]
    [InlineData(1904, "xlsq")]
    public void IndependentExcelTextMatrixMatchesEditUndoAndRoundtrip(int system, string extension)
    {
        using var specJson = ScalarFormulaFixture.Load($"text-type-{system}-spec"); using var expectedJson = ScalarFormulaFixture.Load($"excel-text-type-{system}");
        var spec = specJson.RootElement; var expected = expectedJson.RootElement;
        using var session = new DocumentSession(ScalarFormulaFixture.Book(spec)); var calculator = new WorkbookCalculator(new NoScripts());
        void Check(string state) { ScalarFormulaFixture.Calculate(session, calculator); ScalarFormulaFixture.Compare(session.Snapshot, spec, expected.GetProperty(state)); }
        Check("Initial"); ScalarFormulaFixture.Edit(session, spec); Check("Edited");
        Assert.True(session.Undo()); Assert.False(session.CanUndo); Check("Initial"); Assert.True(session.Redo()); Check("Edited");
        var path = PathFor("matrix." + extension); _storage.Save(path, session.Snapshot); using var reopened = new DocumentSession(_storage.Open(path));
        Assert.False(reopened.Snapshot.FileState!.FormulaReadOnly); Assert.Equal(2, reopened.Snapshot.TextCompatibilityVersion);
        ScalarFormulaFixture.Calculate(reopened, new WorkbookCalculator(new NoScripts())); ScalarFormulaFixture.Compare(reopened.Snapshot, spec, expected.GetProperty("Edited"));
    }

    [Theory]
    [InlineData(1, "xlsx")]
    [InlineData(2, "xlsx")]
    [InlineData(1, "xlsq")]
    [InlineData(2, "xlsq")]
    public void CompatibilityVersionsAndSplitSurrogatesSurviveStorage(int version, string extension)
    {
        using var session = new DocumentSession(UnicodeBook(version)); ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts()));
        var sheet = session.Snapshot.Sheets[0];
        Assert.Equal(version == 1 ? 5 : 3, sheet.GetCell(0, 1)!.Value.NumberValue);
        Assert.Equal(version == 1 ? "\uDE00a" : "a😀", sheet.GetCell(1, 1)!.Value.TextValue);
        Assert.Equal(version == 1 ? "\uD83DZa😀" : "😀Z😀", sheet.GetCell(2, 1)!.Value.TextValue);
        Assert.Equal("😀", sheet.GetCell(3, 1)!.Value.TextValue); Assert.Equal("😀", sheet.GetCell(4, 1)!.Value.TextValue);
        Assert.Equal(version == 1 ? 3 : 2, sheet.GetCell(5, 1)!.Value.NumberValue); Assert.Equal(sheet.GetCell(5, 1)!.Value, sheet.GetCell(6, 1)!.Value);
        var path = PathFor("unicode." + extension); _storage.Save(path, session.Snapshot); using var opened = new DocumentSession(_storage.Open(path));
        Assert.Empty(opened.Snapshot.FileState!.CompatibilityIssues); Assert.Equal(version, opened.Snapshot.TextCompatibilityVersion);
        for (var row = 0; row < 7; row++) Assert.Equal(sheet.GetCell(row, 1)!.Value, opened.Snapshot.Sheets[0].GetCell(row, 1)!.Value);
        ScalarFormulaFixture.Calculate(opened, new WorkbookCalculator(new NoScripts()));
        for (var row = 0; row < 7; row++) Assert.Equal(sheet.GetCell(row, 1)!.Value, opened.Snapshot.Sheets[0].GetCell(row, 1)!.Value);
        if (version == 1 && extension == "xlsx") Assert.Contains("_xDE00_a", Encoding.UTF8.GetString(Part(path, "xl/worksheets/sheet1.xml")));
    }

    [Fact]
    public void CommandsKeepTextVersionAndOneStepHistory()
    {
        using var session = new DocumentSession(UnicodeBook(2)); var id = session.Snapshot.Sheets[0].Id;
        DocumentCommand[] commands = [new SetCellsCommand([new(new(id, 0, 0), CellContent.FromLiteral(CellValue.FromText("👩‍💻")))]),
            new SetCellStyleCommand([new(id, 0, 0)], new(Bold: true)), new RenameSheetCommand(id, "Renamed"),
            new ShiftAxisCommand(id, SheetAxis.Row, 0, 1), new CopySheetCommand(id, "Copy", 1)];
        foreach (var command in commands) { session.Execute(command); Assert.Equal(2, session.Snapshot.TextCompatibilityVersion); }
        foreach (var _ in commands) { Assert.True(session.Undo()); Assert.Equal(2, session.Snapshot.TextCompatibilityVersion); }
        Assert.False(session.CanUndo);
        foreach (var _ in commands) { Assert.True(session.Redo()); Assert.Equal(2, session.Snapshot.TextCompatibilityVersion); }
    }

    [Fact]
    public void ChangedCompatibilityVersionInvalidatesReusableCalculation()
    {
        var original = UnicodeBook(1); var calculator = new WorkbookCalculator(new NoScripts());
        Assert.Equal(5, C(calculator.Calculate(Request(original)), original.Sheets[0], 0, 1).Value?.NumberValue);
        var next = new WorkbookSnapshot(original.DocumentId, original.Revision, original.DateSystem, original.Sheets, textCompatibilityVersion: 2);
        Assert.Equal(3, C(calculator.Calculate(Request(next)), original.Sheets[0], 0, 1).Value?.NumberValue);
    }

    [Theory]
    [InlineData("setVersion", "0")]
    [InlineData("setVersion", "-1")]
    [InlineData("setVersion", "bad")]
    [InlineData("warnBelowVersion", "3")]
    [InlineData("unknown", "1")]
    public void MalformedCompatibilityMetadataIsRejected(string attribute, string value)
    {
        var path = PathFor("bad.xlsx"); _storage.Save(path, UnicodeBook(2)); var xml = XDocument.Parse(Encoding.UTF8.GetString(Part(path, "xl/workbook.xml")));
        xml.Descendants(Compatibility + "version").Single().SetAttributeValue(attribute, value); WritePart(path, "xl/workbook.xml", Encoding.UTF8.GetBytes(xml.ToString()));
        Assert.Throws<InvalidDataException>(() => _storage.Open(path));
    }

    [Fact]
    public void FutureCompatibilityVersionRetainsOriginalPackageAndBlocksFormulaChanges()
    {
        var path = PathFor("future.xlsx"); _storage.Save(path, UnicodeBook(2)); var xml = XDocument.Parse(Encoding.UTF8.GetString(Part(path, "xl/workbook.xml")));
        xml.Descendants(Compatibility + "version").Single().SetAttributeValue("setVersion", 3); WritePart(path, "xl/workbook.xml", Encoding.UTF8.GetBytes(xml.ToString()));
        using var session = new DocumentSession(_storage.Open(path)); Assert.True(session.Snapshot.FileState!.FormulaReadOnly);
        ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts()));
        Assert.Equal(CalculationState.Blocked, session.Snapshot.Sheets[0].GetCell(0, 1)!.Calculation!.State);
        var copy = PathFor("future-copy.xlsx"); _storage.Save(copy, session.Snapshot); Assert.Equal(File.ReadAllBytes(path), File.ReadAllBytes(copy));
    }

    [Theory]
    [InlineData("text-compatibility-v2", 2)]
    [InlineData("extended-cell-errors", 1)]
    public void RequiredFeaturesCannotBeRemoved(string feature, int version)
    {
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [Sheet("S", (0, 0, new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.Calculation)))))], textCompatibilityVersion: version);
        var path = PathFor("missing.xlsq"); _storage.Save(path, book); var manifest = JsonNode.Parse(Part(path, "manifest.json"))!.AsObject();
        var features = manifest["requiredFeatures"]!.AsArray(); Assert.Contains(features, n => n!.GetValue<string>() == feature);
        features.Remove(features.Single(n => n!.GetValue<string>() == feature)); WritePart(path, "manifest.json", Encoding.UTF8.GetBytes(manifest.ToJsonString()));
        Assert.Throws<InvalidDataException>(() => _storage.Open(path));
    }

    [Theory]
    [InlineData("=TEXT(1,\"# ?/?\")")]
    [InlineData("=TEXT(1,\"[DBNum1]0\")")]
    [InlineData("=TEXT(1,\"0\"&\".0\")")]
    [InlineData("=VALUE(A1:A2)")]
    [InlineData("=CONCAT(IF(TRUE,A1:A2,B1:B2))")]
    public void UnimplementedTextFormatsAndArraysKeepOriginalCaches(string formula)
    {
        var path = PathFor("guard.xlsx"); _storage.Save(path, Book(Sheet("S", (0, 5, F(formula, 123)))));
        using var session = new DocumentSession(_storage.Open(path)); Assert.True(session.Snapshot.FileState!.FormulaReadOnly);
        ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts())); var cell = session.Snapshot.Sheets[0].GetCell(0, 5)!;
        Assert.Equal(123, cell.Value.NumberValue); Assert.Equal(CalculationState.Blocked, cell.Calculation!.State);
        var copy = PathFor("guard-copy.xlsx"); _storage.Save(copy, session.Snapshot); Assert.Equal(File.ReadAllBytes(path), File.ReadAllBytes(copy));
    }

    [Theory]
    [InlineData("zh-CN")]
    [InlineData("en-US")]
    [InlineData("de-DE")]
    public void DeclaredTextLanguageIsIndependentOfHostCulture(string culture)
    {
        var old = CultureInfo.CurrentCulture; try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(culture);
            var sheet = Sheet("S", (0, 0, F("=TEXT(1234.5,\"#,##0.00\")")), (1, 0, F("=VALUE(\"1,234.5\")")),
                (2, 0, F("=TEXT(12.5,\"G/通用格式\")")), (3, 0, F("=TEXT(12,\"General\")")));
            using var session = new DocumentSession(Book(sheet)); ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts()));
            var result = session.Snapshot.Sheets[0]; Assert.Equal("1,234.50", result.GetCell(0, 0)!.Value.TextValue);
            Assert.Equal(1234.5, result.GetCell(1, 0)!.Value.NumberValue); Assert.Equal("12.5", result.GetCell(2, 0)!.Value.TextValue);
            Assert.Equal(CellErrorCode.Value, result.GetCell(3, 0)!.Value.ErrorCode);
        }
        finally { CultureInfo.CurrentCulture = old; }
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void CalcErrorsRetainIdentityBeforeRecalculationAndDisappearAfterFix(string extension)
    {
        using var session = new DocumentSession(Book(Sheet("S", (0, 0, T(new string('a', 32767))), (0, 1, F("=CONCAT(A1,\"b\")")), (1, 1, F("=ISERROR(B1)")))));
        ScalarFormulaFixture.Calculate(session, new WorkbookCalculator(new NoScripts())); var path = PathFor("calc." + extension); _storage.Save(path, session.Snapshot);
        using var opened = new DocumentSession(_storage.Open(path)); Assert.Empty(opened.Snapshot.FileState!.CompatibilityIssues);
        Assert.Equal(CellErrorCode.Calculation, opened.Snapshot.Sheets[0].GetCell(0, 1)!.Value.ErrorCode);
        Assert.True(opened.Snapshot.Sheets[0].GetCell(1, 1)!.Value.BooleanValue);
        var exported = PathFor("values.xlsx"); _storage.Export(exported, session.Snapshot);
        var values = _storage.Open(exported); Assert.Empty(values.FileState!.CompatibilityIssues);
        Assert.Equal(CellContentKind.Literal, values.Sheets[0].GetCell(0, 1)!.Content.Kind);
        Assert.Equal(CellErrorCode.Calculation, values.Sheets[0].GetCell(0, 1)!.Value.ErrorCode);
        opened.Execute(new SetCellsCommand([new(new(opened.Snapshot.Sheets[0].Id, 0, 0), CellContent.FromLiteral(CellValue.FromText("a")))]));
        ScalarFormulaFixture.Calculate(opened, new WorkbookCalculator(new NoScripts())); _storage.Save(path, opened.Snapshot);
        var fixedBook = _storage.Open(path); Assert.Empty(fixedBook.FileState!.CompatibilityIssues); Assert.Equal("ab", fixedBook.Sheets[0].GetCell(0, 1)!.Value.TextValue);
        if (extension == "xlsx") { using var zip = ZipFile.OpenRead(path); Assert.Null(zip.GetEntry("xl/metadata.xml")); }
    }

    [Theory]
    [InlineData("vm")]
    [InlineData("count")]
    [InlineData("binding")]
    [InlineData("fallback")]
    [InlineData("bare-error")]
    public void BrokenRichErrorReferencesAreRejected(string fault)
    {
        var path = PathFor("broken.xlsx"); _storage.Save(path, Book(Sheet("S", (0, 0, new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.Calculation)))))));
        var part = fault is "vm" or "fallback" or "bare-error" ? "xl/worksheets/sheet1.xml" : "xl/metadata.xml";
        var xml = XDocument.Parse(Encoding.UTF8.GetString(Part(path, part)));
        switch (fault)
        {
            case "vm": xml.Descendants(Main + "c").Single().SetAttributeValue("vm", "2"); break;
            case "count": xml.Descendants(Main + "valueMetadata").Single().SetAttributeValue("count", "2"); break;
            case "binding": xml.Descendants(Main + "rc").Single().SetAttributeValue("v", "2"); break;
            case "fallback": xml.Descendants(Main + "v").Single().Value = "#N/A"; break;
            default: xml.Descendants(Main + "v").Single().Value = "#CALC!"; break;
        }
        WritePart(path, part, Encoding.UTF8.GetBytes(xml.ToString())); Assert.Throws<InvalidDataException>(() => _storage.Open(path));
    }

    [Fact]
    public void UnmodeledCompatibilityWarningAndRichValuePolicyRemainDiagnosed()
    {
        var path = PathFor("warning.xlsx"); _storage.Save(path, UnicodeBook(2));
        var xml = XDocument.Parse(Encoding.UTF8.GetString(Part(path, "xl/workbook.xml")));
        xml.Descendants(Compatibility + "version").Single().SetAttributeValue("warnBelowVersion", "1");
        WritePart(path, "xl/workbook.xml", Encoding.UTF8.GetBytes(xml.ToString()));
        Assert.Contains(_storage.Open(path).FileState!.CompatibilityIssues, s => s.Contains("阈值"));
        path = PathFor("policy.xlsx"); _storage.Save(path, Book(Sheet("S", (0, 0, new(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.Calculation)))))));
        xml = XDocument.Parse(Encoding.UTF8.GetString(Part(path, "xl/metadata.xml")));
        xml.Descendants(Main + "metadataType").Single().SetAttributeValue("pasteValues", "0"); WritePart(path, "xl/metadata.xml", Encoding.UTF8.GetBytes(xml.ToString()));
        var opened = _storage.Open(path); Assert.Equal(CellErrorCode.Calculation, opened.Sheets[0].GetCell(0, 0)!.Value.ErrorCode);
        Assert.Contains(opened.FileState!.CompatibilityIssues, s => s.Contains("操作策略"));
    }
}
