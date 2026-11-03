using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Core.Test.Fixtures;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class DynamicArrayStorageTests
{
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void SpillErrorStoresNativeDimensionsAndCanBeClearedAfterReopen(string extension)
    {
        var sheet = M3CalculationTests.Sheet("Data", (0,0,M3CalculationTests.F("=SEQUENCE(3,2)")),
            (1,1,new CellSnapshot(CellContent.FromLiteral(CellValue.FromText("blocked")))), (6,0,M3CalculationTests.F("=SEQUENCE(0)")));
        using var session = new DocumentSession(M3CalculationTests.Book(sheet)); Calculate(session);
        Assert.Equal(CellValue.FromError(CellErrorCode.Spill),session.Snapshot.Sheets[0].GetCell(0,0)!.Value);
        var path = Path.Combine(EvidenceDirectory(),"spill." + extension); var storage = new TypedWorkbookStorage(); storage.Save(path,session.Snapshot);
        using var opened = new DocumentSession(storage.Open(path)); var read = opened.Snapshot.Sheets[0];
        Assert.False(opened.Snapshot.FileState!.FormulaReadOnly); Assert.Empty(opened.Snapshot.FileState.CompatibilityIssues);
        Assert.Equal(new FormulaArrayDimensions(3,2),read.GetCell(0,0)!.Calculation!.BlockedSpill);
        Assert.Equal(CellValue.FromError(CellErrorCode.Calculation),read.GetCell(6,0)!.Value);
        opened.Execute(new ClearCellsCommand([new(read.Id,1,1)])); Calculate(opened);
        Assert.Equal(CellValue.FromNumber(6),opened.Snapshot.Sheets[0].GetCell(2,1)!.Value);
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(path)!)!,"latest-spill-" + extension + ".txt"),path);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void ArrayFormulaMetadataAndDerivedCellsSurviveSaveReopen(string extension)
    {
        var sheet = M3CalculationTests.Sheet("Data", (0,0,M3CalculationTests.F("=SEQUENCE(3,2)")),
            (0,3,M3CalculationTests.F("=SUM(A1#)")), (5,0,M3CalculationTests.F("=FILTER(SEQUENCE(3),SEQUENCE(3)>10)")));
        using var session = new DocumentSession(M3CalculationTests.Book(sheet));
        Calculate(session);
        var path = Path.Combine(EvidenceDirectory(), "arrays." + extension);
        var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot);
        using var opened = new DocumentSession(storage.Open(path));
        Assert.False(opened.Snapshot.FileState!.FormulaReadOnly);
        Assert.Empty(opened.Snapshot.FileState.CompatibilityIssues);
        var readSheet = opened.Snapshot.Sheets[0];
        Assert.True(readSheet.GetCell(0,0)!.Content.DynamicArray);
        Assert.Equal(new CellAddress(readSheet.Id,0,0), readSheet.GetCell(2,1)!.SpillOwner);
        Assert.Equal(CellValue.Blank, readSheet.GetCell(2,1)!.Content.LiteralValue);
        Assert.Equal(CellValue.FromNumber(6), readSheet.GetCell(2,1)!.Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Calculation), readSheet.GetCell(5,0)!.Value);
        Calculate(opened);
        Assert.Equal(CellValue.FromNumber(21), opened.Snapshot.Sheets[0].GetCell(0,3)!.Value);
        opened.Execute(new SetCellsCommand([new(new(readSheet.Id,0,0),CellContent.FromFormula("=SEQUENCE(2)"))]));
        Calculate(opened); storage.Save(Path.Combine(Path.GetDirectoryName(path)!, "edited." + extension), opened.Snapshot);
        Assert.Null(opened.Snapshot.Sheets[0].GetCell(2,1));
        if (extension == "xlsq")
        {
            using var zip = ZipFile.OpenRead(path); using var stream = zip.GetEntry("manifest.json")!.Open();
            using var json = JsonDocument.Parse(stream);
            Assert.Contains(json.RootElement.GetProperty("requiredFeatures").EnumerateArray(), f => f.GetString() == "dynamic-arrays");
        }
    }

    [Fact]
    public void NamedArrayAndExplicitArrayRangeSaveWithNativeMetadataAndStaleSaveIsRejected()
    {
        var sheet = M3CalculationTests.Sheet("Data",(0,0,M3CalculationTests.F("=Numbers")),
            (0,3,new CellSnapshot(CellContent.FromFormula("=A1:A3",dynamicArray:true))));
        using var session = new DocumentSession(M3CalculationTests.Book(sheet).WithDefinedNames([new QuickExcel.Formulas.DefinedName("Numbers","=SEQUENCE(3)")]));
        Calculate(session); var storage = new TypedWorkbookStorage(); var path = Path.Combine(EvidenceDirectory(),"named.xlsx");
        storage.Save(path,session.Snapshot);
        using var opened = new DocumentSession(storage.Open(path)); Calculate(opened);
        Assert.False(opened.Snapshot.FileState!.FormulaReadOnly);
        Assert.Equal(CellValue.FromNumber(3),opened.Snapshot.Sheets[0].GetCell(2,3)!.Value);
        session.Execute(new SetCellsCommand([new(new(sheet.Id,9,9),CellContent.FromLiteral(CellValue.FromNumber(1)))]));
        Assert.Throws<InvalidOperationException>(() => storage.Save(Path.Combine(Path.GetDirectoryName(path)!,"stale.xlsx"),session.Snapshot));
    }

    [Fact]
    public void IndependentDynamicMetadataIsEditableButLegacyArrayStillUsesOriginalPackage()
    {
        var path = Path.Combine(EvidenceDirectory(), "independent.xlsx");
        File.WriteAllBytes(path, Fixture());
        var storage = new TypedWorkbookStorage(); using var session = new DocumentSession(storage.Open(path));
        Assert.False(session.Snapshot.FileState!.FormulaReadOnly);
        Assert.Equal(CellValue.FromNumber(3), session.Snapshot.Sheets[0].GetCell(2,0)!.Value);
        Assert.NotNull(session.Snapshot.Sheets[0].GetCell(2,0)!.SpillOwner);
        Calculate(session);
        Assert.Equal(CellValue.FromNumber(3), session.Snapshot.Sheets[0].GetCell(2,0)!.Value);
        var legacy = Path.Combine(Path.GetDirectoryName(path)!, "legacy.xlsx"); File.WriteAllBytes(legacy, Fixture(marker: false));
        Assert.True(storage.Open(legacy).FileState!.FormulaReadOnly);
    }

    [Theory]
    [InlineData("A1:A100001")]
    [InlineData("A1:A3", true)]
    public void InvalidDynamicCacheGeometryRejectsImport(string range, bool childFormula = false)
    {
        var path = Path.Combine(EvidenceDirectory(), "invalid.xlsx"); File.WriteAllBytes(path, Fixture(range: range, childFormula: childFormula));
        Assert.Throws<InvalidDataException>(() => new TypedWorkbookStorage().Open(path));
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void IndependentExcelArrayOracleMatchesCalculationAndRoundTrip(string extension)
    {
        using var specJson = ScalarFormulaFixture.Load("dynamic-array-spec"); using var oracleJson = ScalarFormulaFixture.Load("excel-dynamic-arrays");
        var spec = specJson.RootElement;
        var data = spec.GetProperty("Inputs").EnumerateArray().ToDictionary(i => ScalarFormulaFixture.Position(i.GetProperty("Address").GetString()!),
            i => new CellSnapshot(ScalarFormulaFixture.Content(i)));
        var results = spec.GetProperty("Cases").EnumerateArray().Where(i => i.TryGetProperty("Formula", out _)).ToDictionary(
            i => ScalarFormulaFixture.Position(i.GetProperty("Address").GetString()!), i => M3CalculationTests.F(i.GetProperty("Formula").GetString()!));
        using var session = new DocumentSession(new(Guid.NewGuid(),0,ExcelDateSystem.Windows1900,[new(SheetId.New(),"Data",data), new(SheetId.New(),"Results",results)]));
        void Compare(WorkbookSnapshot book)
        {
            var sheet = book.Sheets[1]; var failures = new List<string>();
            foreach (var item in oracleJson.RootElement.GetProperty("Initial").EnumerateArray())
            {
                var position = ScalarFormulaFixture.Position(item.GetProperty("Address").GetString()!);
                var actual = sheet.GetCell(position.Row,position.Column)?.Value ?? CellValue.Blank; var expected = ScalarFormulaFixture.Value(item);
                if (actual != expected) failures.Add($"{item.GetProperty("Id").GetString()}: expected {expected}, actual {actual}");
            }
            Assert.True(failures.Count == 0,string.Join(Environment.NewLine,failures));
        }
        Calculate(session); Compare(session.Snapshot);
        var path = Path.Combine(EvidenceDirectory(), "oracle." + extension); var storage = new TypedWorkbookStorage();
        var rawPath = Path.Combine(Path.GetDirectoryName(path)!, "raw.xlsx");
        File.WriteAllBytes(rawPath,TypedWorkbookXml.Write(session.Snapshot,false,CancellationToken.None));
        var raw = storage.Open(rawPath); var differences = new List<string>();
        for (var i = 0; i < raw.Sheets.Length; i++)
        foreach (var position in session.Snapshot.Sheets[i].Cells.Keys.Union(raw.Sheets[i].Cells.Keys))
        {
            var a = session.Snapshot.Sheets[i].Cells.GetValueOrDefault(position); var b = raw.Sheets[i].Cells.GetValueOrDefault(position);
            if (a?.Content != b?.Content || a?.Value != b?.Value || (a?.SpillOwner?.Row,a?.SpillOwner?.Column) != (b?.SpillOwner?.Row,b?.SpillOwner?.Column))
                differences.Add($"{i}:{position} source {a?.Content} dynamic={a?.Content.DynamicArray} value={a?.Value}, read {b?.Content} dynamic={b?.Content.DynamicArray} value={b?.Value}");
        }
        Assert.True(differences.Count == 0,string.Join(Environment.NewLine,differences));
        storage.Save(path,session.Snapshot);
        using var opened = new DocumentSession(storage.Open(path));
        Assert.False(opened.Snapshot.FileState!.FormulaReadOnly); Compare(opened.Snapshot); Calculate(opened); Compare(opened.Snapshot);
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(path)!)!, "latest-oracle-" + extension + ".txt"), path);
    }

    [Fact]
    public void PortableExportKeepsExplicitArrayModeAndFreezesNonPortableSpillClosure()
    {
        var sheet = M3CalculationTests.Sheet("Data", (0,0,M3CalculationTests.F("=SEQUENCE(3)")),
            (0,3,new CellSnapshot(CellContent.FromFormula("=A1:A3",dynamicArray:true))),
            (0,6,M3CalculationTests.F("=D2+1")));
        using var session = new DocumentSession(M3CalculationTests.Book(sheet)); Calculate(session);
        var storage = new TypedWorkbookStorage(); var folder = EvidenceDirectory(); var path = Path.Combine(folder,"portable.xlsx");
        storage.ExportPortable(path,session.Snapshot);
        using (var reopened = new DocumentSession(storage.Open(path)))
        {
            Assert.True(reopened.Snapshot.Sheets[0].GetCell(0,3)!.Content.DynamicArray);
            Calculate(reopened); Assert.Equal(CellValue.FromNumber(3),reopened.Snapshot.Sheets[0].GetCell(2,3)!.Value);
            Assert.Equal(CellValue.FromNumber(3),reopened.Snapshot.Sheets[0].GetCell(0,6)!.Value);
        }
        // Independent old-cache fixture: an unavailable producer and its existing spill footprint.
        var cachedSheet = session.Snapshot.Sheets[0]; var cached = cachedSheet.GetCell(0,0)!.Calculation!;
        var oldCells = cachedSheet.Cells.SetItem(new(0,0),new(CellContent.FromFormula("=UNKNOWN()",dynamicArray:true),
            new(cached.Value,CalculationState.Blocked,cached.InputRevision,cached.CalculationGeneration,cached.ScriptGeneration,array:cached.Array)));
        var oldBook = session.Snapshot.WithSheets([cachedSheet.WithCells(oldCells)]).WithRevision(session.Snapshot.Revision+1);
        Assert.Equal(3,storage.InspectPortableExport(oldBook).FrozenFormulas);
        var frozenPath = Path.Combine(folder,"frozen.xlsx");
        Assert.Throws<InvalidOperationException>(()=>storage.ExportPortable(frozenPath,oldBook));
        storage.ExportPortable(frozenPath,oldBook,allowStaleCaches:true);
        var frozen = storage.Open(frozenPath); Assert.All(frozen.Sheets[0].Cells.Values,c=>{
            Assert.Equal(CellContentKind.Literal,c.Content.Kind); Assert.Null(c.SpillOwner);
        });
        Assert.Equal(CellValue.FromNumber(3),frozen.Sheets[0].GetCell(2,3)!.Value);
        Assert.Equal(CellValue.FromNumber(3),frozen.Sheets[0].GetCell(0,6)!.Value);
    }

    internal static void Calculate(DocumentSession session) => Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new M3CalculationTests.NoScripts()).Calculate(session.BeginCalculation())));
    internal static string EvidenceDirectory()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName,"QuickExcel.slnx"))) directory = directory.Parent;
        if (directory is null) throw new InvalidOperationException("Repository root not found for test evidence.");
        return Directory.CreateDirectory(Path.Combine(directory.FullName,"artifacts","e5","storage",Guid.NewGuid().ToString("N"))).FullName;
    }
    private static byte[] Fixture(bool marker = true, string range = "A1:A3", bool childFormula = false) => WorkbookPackageFixture.Xlsx(
        $"""<row r="1"><c r="A1"{(marker ? " cm=\"1\"" : "")}><f t="array" ref="{range}">_xlfn.SEQUENCE(3)</f><v>1</v></c></row><row r="2"><c r="A2">{(childFormula ? "<f>1+1</f>" : "")}<v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>""",
        new Dictionary<string,string>
        {
            ["[Content_Types].xml"] = """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>""",
            ["xl/_rels/workbook.xml.rels"] = """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="meta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>""",
            ["xl/metadata.xml"] = """<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>"""
        });
}
