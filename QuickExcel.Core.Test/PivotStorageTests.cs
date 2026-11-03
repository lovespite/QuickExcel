using QuickExcel.Domain;
using QuickExcel.Storage;
using QuickExcel.Core.Test.Fixtures;

namespace QuickExcel.Core.Test;

public sealed class PivotStorageTests
{
    [Fact]
    public void IndependentNativeDefinitionAndIndexedCacheRemainEditableAndExportAsPlainValues()
    {
        var directory=DynamicArrayStorageTests.EvidenceDirectory(); var path=Path.Combine(directory,"independent-pivot.xlsx"); File.WriteAllBytes(path,Fixture());
        var storage=new TypedWorkbookStorage(); using var session=new DocumentSession(storage.Open(path));
        Assert.Empty(session.Snapshot.FileState!.CompatibilityIssues); var id=session.Snapshot.Sheets[0].Id;
        Assert.Single(session.Snapshot.Sheets[0].Data.Pivots);
        session.Execute(new SetCellsCommand([new(new(id,1,1),CellContent.FromLiteral(CellValue.FromNumber(3)))]));
        Assert.True(session.Snapshot.Sheets[0].Data.Pivots.Single().Stale);
        Assert.Throws<InvalidOperationException>(()=>storage.Export(Path.Combine(directory,"not-fresh.xlsx"),session.Snapshot));
        storage.Export(Path.Combine(directory,"old-values.xlsx"),session.Snapshot,allowStaleCaches:true,discardUnsupported:false);
        Assert.Empty(storage.Open(Path.Combine(directory,"old-values.xlsx")).Sheets[0].Data.Pivots);
        session.Execute(new RefreshPivotCommand(id,"Pivot"));
        Assert.Equal(CellValue.FromNumber(10),session.Snapshot.Sheets[0].GetCell(3,4)!.Value);
        storage.Save(Path.Combine(directory,"refreshed.xlsx"),session.Snapshot);
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(directory)!,"latest-pivot-independent.txt"),Path.Combine(directory,"refreshed.xlsx"));
    }

    [Theory]
    [InlineData("average",false)]
    [InlineData("sum",true)]
    public void UnsupportedAggregationAndExternalSourcePreserveOriginalAndNeverGuessRefresh(string aggregation,bool external)
    {
        var directory=DynamicArrayStorageTests.EvidenceDirectory(); var path=Path.Combine(directory,"opaque.xlsx"); var bytes=Fixture(aggregation:aggregation,external:external);
        File.WriteAllBytes(path,bytes); var storage=new TypedWorkbookStorage(); using var session=new DocumentSession(storage.Open(path));
        Assert.Empty(session.Snapshot.Sheets[0].Data.Pivots);
        Assert.Contains(session.Snapshot.FileState!.CompatibilityDiagnostics,d=>d.Feature.Contains("pivot",StringComparison.OrdinalIgnoreCase));
        Assert.Equal(CellValue.FromNumber(8),session.Snapshot.Sheets[0].GetCell(3,4)!.Value);
        Assert.Throws<NotSupportedException>(()=>session.Execute(new SetCellsCommand([new(new(session.Snapshot.Sheets[0].Id,1,1),CellContent.FromLiteral(CellValue.FromNumber(3)))])));
        var copy=Path.Combine(directory,"preserved.xlsx"); storage.Save(copy,session.Snapshot); Assert.Equal(bytes,File.ReadAllBytes(copy));
    }

    [Theory]
    [InlineData("9",3)]
    [InlineData("0",4)]
    public void InvalidCacheIndexOrRecordCountRejectsImport(string sharedIndex,int count)
    {
        var path=Path.Combine(DynamicArrayStorageTests.EvidenceDirectory(),"malformed.xlsx"); File.WriteAllBytes(path,Fixture(sharedIndex:sharedIndex,count:count));
        Assert.ThrowsAny<InvalidDataException>(()=>new TypedWorkbookStorage().Open(path));
    }

    private static byte[] Fixture(string aggregation="sum",bool external=false,string sharedIndex="0",int count=3) => WorkbookPackageFixture.Xlsx("""
        <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c><c r="D1" t="inlineStr"><is><t>Region</t></is></c><c r="E1" t="inlineStr"><is><t>Total</t></is></c></row>
        <row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>1</v></c><c r="D2" t="inlineStr"><is><t>A</t></is></c><c r="E2"><v>3</v></c></row>
        <row r="3"><c r="A3" t="inlineStr"><is><t>A</t></is></c><c r="B3"><v>2</v></c><c r="D3" t="inlineStr"><is><t>B</t></is></c><c r="E3"><v>5</v></c></row>
        <row r="4"><c r="A4" t="inlineStr"><is><t>B</t></is></c><c r="B4"><v>5</v></c><c r="D4" t="inlineStr"><is><t>总计</t></is></c><c r="E4"><v>8</v></c></row>
        """,new Dictionary<string,string>
        {
            ["[Content_Types].xml"]="""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/pivotTables/pivot.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/><Override PartName="/xl/pivotCache/definition.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/><Override PartName="/xl/pivotCache/records.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/></Types>""",
            ["xl/workbook.xml"]="""<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="sheet"/></sheets><pivotCaches><pivotCache cacheId="7" r:id="cache"/></pivotCaches></workbook>""",
            ["xl/_rels/workbook.xml.rels"]="""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="sheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="cache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/definition.xml"/></Relationships>""",
            ["xl/worksheets/_rels/sheet1.xml.rels"]="""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="pivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivot.xml"/></Relationships>""",
            ["xl/pivotTables/_rels/pivot.xml.rels"]="""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="cache" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/definition.xml"/></Relationships>""",
            ["xl/pivotCache/_rels/definition.xml.rels"]="""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="records" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="records.xml"/></Relationships>""",
            ["xl/pivotCache/definition.xml"]=$"""<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="records" recordCount="{count}"><cacheSource type="{(external ? "external" : "worksheet")}"><worksheetSource sheet="Sheet1" ref="A1:B4"/></cacheSource><cacheFields count="2"><cacheField name="Region"><sharedItems count="2"><s v="A"/><s v="B"/></sharedItems></cacheField><cacheField name="Amount"><sharedItems count="0"/></cacheField></cacheFields></pivotCacheDefinition>""",
            ["xl/pivotCache/records.xml"]=$"""<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3"><r><x v="{sharedIndex}"/><n v="1"/></r><r><x v="0"/><n v="2"/></r><r><x v="1"/><n v="5"/></r></pivotCacheRecords>""",
            ["xl/pivotTables/pivot.xml"]=$"""<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Pivot" cacheId="7" dataCaption="值" compact="0" compactData="0"><location ref="D1:E4" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow" defaultSubtotal="0"><items count="2"><item x="0"/><item x="1"/></items></pivotField><pivotField dataField="1"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i><x/></i></colItems><dataFields count="1"><dataField name="Total" fld="1" subtotal="{aggregation}"/></dataFields></pivotTableDefinition>"""
        });

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void DefinitionCacheDisplayAndStaleStateRoundTrip(string extension)
    {
        using var session=new DocumentSession(PivotReportTests.SourceBook()); var target=session.Snapshot.Sheets[1].Id;
        var source=new PivotSource("Data",new(0,0,8,2),"SalesData");
        session.Execute(new CreatePivotCommand(target,new(2,0),source,"SalesSum",0,1,2,PivotAggregation.Sum,"Total"));
        session.Execute(new CreatePivotCommand(target,new(11,0),source,"SalesCount",0,1,2,PivotAggregation.Count,"Count"));
        var path=Path.Combine(DynamicArrayStorageTests.EvidenceDirectory(),"pivots."+extension); var storage=new TypedWorkbookStorage();
        storage.Save(path,session.Snapshot);
        using var opened=new DocumentSession(storage.Open(path)); var sheet=opened.Snapshot.Sheets[1];
        Assert.Empty(opened.Snapshot.FileState!.CompatibilityIssues);
        Assert.Equal(2,sheet.Data.Pivots.Length);
        Assert.True(session.Snapshot.Sheets[1].Data.Pivots[0].ContentEquals(sheet.Data.Pivots[0]));
        opened.Execute(new SetCellsCommand([new(new(opened.Snapshot.Sheets[0].Id,1,2),CellContent.FromLiteral(CellValue.FromNumber(12)))]));
        Assert.All(opened.Snapshot.Sheets[1].Data.Pivots,p=>Assert.True(p.Stale));
        var stale=Path.Combine(Path.GetDirectoryName(path)!,"stale."+extension); storage.Save(stale,opened.Snapshot);
        var reread=storage.Open(stale); Assert.All(reread.Sheets[1].Data.Pivots,p=>Assert.True(p.Stale));
        Assert.Equal(CellValue.FromNumber(39),reread.Sheets[1].GetCell(7,3)!.Value);
        opened.Execute(new RefreshPivotCommand(sheet.Id,"SalesSum"));
        Assert.Equal(CellValue.FromNumber(41),opened.Snapshot.Sheets[1].GetCell(7,3)!.Value);
        storage.Save(Path.Combine(Path.GetDirectoryName(path)!,"refreshed."+extension),opened.Snapshot);
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(path)!)!,"latest-pivot-"+extension+".txt"),path);
    }
}
