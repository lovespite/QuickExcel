using System.Collections.Immutable;
using System.Globalization;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using static QuickExcel.Storage.PackageIO;
using static QuickExcel.Storage.TypedStyles;
using static QuickExcel.Storage.TypedWorkbookXml;
using static QuickExcel.Storage.XlsxPackage;

namespace QuickExcel.Storage;

internal static partial class XlsxPivots
{
    private static string Type(string suffix) => "application/vnd.openxmlformats-officedocument.spreadsheetml." + suffix + "+xml";
    private static readonly XNamespace X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    private static readonly XNamespace Layout = "http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout";
    private static readonly XNamespace Clean = "http://schemas.microsoft.com/office/spreadsheetml/2026/pivotDataSourceClean";

    public static bool Present(WorkbookSnapshot snapshot) => snapshot.Sheets.Any(s => !s.Data.Pivots.IsEmpty);

    public static void Write(WorkbookSnapshot book,Dictionary<string,byte[]> parts,XElement workbook,XElement relationships,XElement types,CancellationToken cancellationToken)
    {
        var caches=new XElement(Main+"pivotCaches"); var serial=0;
        for (var sheetIndex=0;sheetIndex<book.Sheets.Length;sheetIndex++)
        foreach (var pivot in book.Sheets[sheetIndex].Data.Pivots)
        {
            cancellationToken.ThrowIfCancellationRequested(); serial++;
            var id=serial; var cache=pivot.Cache;
            var result=PivotReport.Build(cache,new(pivot.Range.Top,pivot.Range.Left),pivot.RowField,pivot.ColumnField,pivot.DataField,pivot.Aggregation,pivot.DataCaption,pivot.GrandCaption,cancellationToken);
            if (result.Range!=pivot.Range) throw new InvalidOperationException("透视缓存与显示区域尺寸不一致。");
            var cachePart=$"xl/pivotCache/pivotCacheDefinition{id}.xml"; var recordsPart=$"xl/pivotCache/pivotCacheRecords{id}.xml"; var pivotPart=$"xl/pivotTables/pivotTable{id}.xml";
            foreach (var entry in new[] {(cachePart,"pivotCacheDefinition"),(recordsPart,"pivotCacheRecords"),(pivotPart,"pivotTable")})
                types.Add(new XElement(ContentTypes+"Override",new XAttribute("PartName","/"+entry.Item1),new XAttribute("ContentType",Type(entry.Item2))));
            relationships.Add(Relationship("pivotCache"+id,"pivotCacheDefinition",$"pivotCache/pivotCacheDefinition{id}.xml"));
            caches.Add(new XElement(Main+"pivotCache",new XAttribute("cacheId",id),new XAttribute(Rel+"id","pivotCache"+id)));
            parts[RelationshipPart(cachePart)]=XmlBytes(new(new XElement(PackageRel+"Relationships",Relationship("records","pivotCacheRecords",$"pivotCacheRecords{id}.xml"))));
            parts[RelationshipPart(pivotPart)]=XmlBytes(new(new XElement(PackageRel+"Relationships",Relationship("cache","pivotCacheDefinition",$"../pivotCache/pivotCacheDefinition{id}.xml"))));
            var sheetRel=RelationshipPart($"xl/worksheets/sheet{sheetIndex+1}.xml");
            var sheetRelations=parts.TryGetValue(sheetRel,out var prior) ? Xml(prior,new(cancellationToken)).Root! : new XElement(PackageRel+"Relationships");
            sheetRelations.Add(Relationship("pivotTable"+id,"pivotTable",$"../pivotTables/pivotTable{id}.xml")); parts[sheetRel]=XmlBytes(new(sheetRelations));
            var axes=new Dictionary<int,ImmutableArray<CellValue>> { [pivot.RowField]=result.RowItems };
            if (pivot.ColumnField is { } columnField) axes.Add(columnField,result.ColumnItems);
            var fields=new XElement(Main+"cacheFields",new XAttribute("count",cache.Fields.Length));
            foreach (var (field,index) in cache.Fields.Select((f,i)=>(f,i)))
            {
                var values=axes.GetValueOrDefault(index,[]);
                var shared=new XElement(Main+"sharedItems",new XAttribute("count",values.Length));
                var observed=cache.Records.Select(r=>r[index]).ToArray();
                shared.SetAttributeValue("containsBlank",observed.Any(v=>v.Kind==CellValueKind.Blank) ? 1 : 0);
                shared.SetAttributeValue("containsString",observed.Any(v=>v.Kind==CellValueKind.Text) ? 1 : 0);
                shared.SetAttributeValue("containsNumber",observed.Any(v=>v.Kind==CellValueKind.Number) ? 1 : 0);
                shared.SetAttributeValue("containsMixedTypes",observed.Select(v=>v.Kind).Distinct().Count()>1 ? 1 : 0);
                shared.SetAttributeValue("containsSemiMixedTypes",observed.Any(v=>v.Kind!=CellValueKind.Number) ? 1 : 0);
                shared.Add(values.Select(Value));
                fields.Add(new XElement(Main+"cacheField",new XAttribute("name",Encode(field)),new XAttribute("numFmtId",0),shared));
            }
            var source=new XElement(Main+"worksheetSource",pivot.Source.TableName is { } table ? new XAttribute("name",table) :
                new object[] {new XAttribute("sheet",Encode(pivot.Source.Sheet)),new XAttribute("ref",pivot.Source.Range.ToString())});
            var definition=new XElement(Main+"pivotCacheDefinition",new XAttribute(Rel+"id","records"),new XAttribute("recordCount",cache.Records.Length),
                new XAttribute("createdVersion",6),new XAttribute("refreshedVersion",6),new XAttribute("minRefreshableVersion",3),
                new XAttribute("refreshOnLoad",0),new XAttribute("enableRefresh",1),new XAttribute("invalid",pivot.Stale ? 1 : 0),
                new XElement(Main+"cacheSource",new XAttribute("type","worksheet"),source),fields);
            parts[cachePart]=XmlBytes(new(definition));
            var records=new XElement(Main+"pivotCacheRecords",new XAttribute("count",cache.Records.Length));
            foreach (var record in cache.Records)
            {
                cancellationToken.ThrowIfCancellationRequested();
                records.Add(new XElement(Main+"r",record.Select(Value)));
            }
            parts[recordsPart]=XmlBytes(new(records));
            var pivotFields=new XElement(Main+"pivotFields",new XAttribute("count",cache.Fields.Length));
            for (var index=0;index<cache.Fields.Length;index++)
            {
                var field=new XElement(Main+"pivotField",new XAttribute("compact",0),new XAttribute("outline",0),new XAttribute("showAll",0));
                if (index==pivot.DataField) field.SetAttributeValue("dataField",1);
                if (axes.TryGetValue(index,out var items))
                {
                    field.SetAttributeValue("axis",index==pivot.RowField ? "axisRow" : "axisCol"); field.SetAttributeValue("defaultSubtotal",0);
                    field.Add(new XElement(Main+"items",new XAttribute("count",items.Length),Enumerable.Range(0,items.Length).Select(i=>new XElement(Main+"item",new XAttribute("x",i)))));
                }
                pivotFields.Add(field);
            }
            XElement Axis(string name,int field) => new(Main+name,new XAttribute("count",1),new XElement(Main+"field",new XAttribute("x",field)));
            XElement Items(string name,int count) => new(Main+name,new XAttribute("count",count+1),Enumerable.Range(0,count).Select(i=>
                new XElement(Main+"i",new XElement(Main+"x",new XAttribute("v",i)))),new XElement(Main+"i",new XAttribute("t","grand"),new XElement(Main+"x")));
            var root=new XElement(Main+"pivotTableDefinition",new XAttribute("name",pivot.Name),new XAttribute("cacheId",id),
                new XAttribute("dataCaption","值"),new XAttribute("grandTotalCaption",pivot.GrandCaption),new XAttribute("createdVersion",6),new XAttribute("updatedVersion",6),
                new XAttribute("minRefreshableVersion",3),new XAttribute("compact",0),new XAttribute("compactData",0),new XAttribute("indent",0),
                new XAttribute("rowGrandTotals",1),new XAttribute("colGrandTotals",1),
                new XElement(Main+"location",new XAttribute("ref",pivot.Range.ToString()),new XAttribute("firstHeaderRow",pivot.ColumnField is null ? 0 : 1),
                    new XAttribute("firstDataRow",pivot.ColumnField is null ? 1 : 2),new XAttribute("firstDataCol",1)),pivotFields,Axis("rowFields",pivot.RowField),Items("rowItems",result.RowItems.Length));
            if (pivot.ColumnField is { } column) root.Add(Axis("colFields",column),Items("colItems",result.ColumnItems.Length));
            else root.Add(new XElement(Main+"colItems",new XAttribute("count",1),new XElement(Main+"i",new XElement(Main+"x"))));
            root.Add(new XElement(Main+"dataFields",new XAttribute("count",1),new XElement(Main+"dataField",new XAttribute("name",Encode(pivot.DataCaption)),
                new XAttribute("fld",pivot.DataField),new XAttribute("subtotal",pivot.Aggregation==PivotAggregation.Count ? "count" : "sum"),new XAttribute("baseField",0),new XAttribute("baseItem",0))),
                new XElement(Main+"pivotTableStyleInfo",new XAttribute("showRowHeaders",1),new XAttribute("showColHeaders",1),new XAttribute("showRowStripes",0),new XAttribute("showColStripes",0),new XAttribute("showLastColumn",1)),
                new XElement(Main+"extLst",new XElement(Main+"ext",new XAttribute("uri","{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}"),new XElement(X14+"pivotTableDefinition",new XAttribute("hideValuesRow",1)))));
            parts[pivotPart]=XmlBytes(new(root));
        }
        if (serial>0) workbook.Add(caches);
    }

    private static XElement Value(CellValue value) => value.Kind switch
    {
        CellValueKind.Blank=>new(Main+"m"), CellValueKind.Text=>new(Main+"s",new XAttribute("v",Encode(value.TextValue))),
        CellValueKind.Number=>new(Main+"n",new XAttribute("v",value.NumberValue.ToString("R",CultureInfo.InvariantCulture))),
        CellValueKind.Boolean=>new(Main+"b",new XAttribute("v",value.BooleanValue ? 1 : 0)),
        CellValueKind.Error=>new(Main+"e",new XAttribute("v",CellInput.DisplayText(value))), _=>throw new InvalidOperationException()
    };
}
