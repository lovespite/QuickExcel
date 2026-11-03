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
    private static bool Overlaps(CellRectangle a,CellRectangle b) => a.Top<=b.Bottom && a.Bottom>=b.Top && a.Left<=b.Right && a.Right>=b.Left;
    private static double Number(string text) => double.TryParse(text,NumberStyles.Float,CultureInfo.InvariantCulture,out var value) && double.IsFinite(value)
        ? value : throw new InvalidDataException("无效透视数字。");
    public static WorkbookSnapshot Read(XlsxPackage package,string workbookPart,XElement workbook,WorkbookSnapshot book,IReadOnlyList<WorkbookPartBinding> sheets)
    {
        if (workbook.Elements(Main+"pivotCaches").Count()>1) throw new InvalidDataException("工作簿透视缓存集合重复。");
        var caches=new Dictionary<int,string>();
        foreach (var entry in workbook.Elements(Main+"pivotCaches").Elements())
        {
            Shape(entry,"","cacheId r:id");
            var id=Int(Attr(entry,"cacheId"),int.MaxValue);
            if (caches.Count>=128 || !caches.TryAdd(id,package.Related(workbookPart,Attr(entry,Rel+"id"),"pivotCacheDefinition")))
                throw new InvalidDataException("透视缓存编号重复或数量超过 128。");
        }
        var parsedCaches=new Dictionary<string,(PivotSource Source,PivotCache Cache,ImmutableArray<ImmutableArray<CellValue>> Shared,bool Stale)>();
        var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result=book.Sheets.ToArray();
        for (var index=0;index<result.Length;index++)
        {
            var sheet=result[index]; var sheetPart=sheets[index].Part; var pivots=ImmutableArray.CreateBuilder<WorksheetPivot>();
            foreach (var relation in package.Relationships(sheetPart).Values.Where(r=>r.Type==Rel.NamespaceName+"/pivotTable"))
            {
                package.Budget.Check(); var part=package.Related(sheetPart,relation.Id,"pivotTable"); package.RequireType(part,Type("pivotTable"));
                var definition=package.XmlPart(part).Root!;
                try
                {
                    if (definition.Name!=Main+"pivotTableDefinition") throw new InvalidDataException("透视定义根元素无效。");
                    var id=Int(Attr(definition,"cacheId"),int.MaxValue);
                    if (!caches.TryGetValue(id,out var cachePart) || package.RelatedByType(part,"pivotCacheDefinition",true)!=cachePart)
                        throw new InvalidDataException("透视定义的缓存关系与工作簿不一致。");
                    if (!parsedCaches.TryGetValue(cachePart,out var cache)) parsedCaches[cachePart]=cache=ReadCache(package,cachePart,book);
                    var pivot=ReadDefinition(package,part,definition,sheet,cache.Source,cache.Cache,cache.Shared,cache.Stale);
                    var source=PivotReport.Resolve(book,pivot.Source);
                    if (source.Sheet.Id==sheet.Id && Overlaps(source.Range,pivot.Range)) throw new NotSupportedException("透视来源与自身结果重叠。");
                    if (source.Range.Rows-1!=cache.Cache.Records.Length || source.Range.Columns!=cache.Cache.Fields.Length) pivot=pivot.WithStale(true);
                    else
                    {
                        for (var c=0;c<cache.Cache.Fields.Length;c++)
                            if (source.Sheet.GetCell(source.Range.Top,source.Range.Left+c)?.Value!=CellValue.FromText(cache.Cache.Fields[c])) pivot=pivot.WithStale(true);
                        for (var r=0;r<cache.Cache.Records.Length;r++)
                        for (var c=0;c<cache.Cache.Fields.Length;c++)
                        {
                            package.Budget.Check();
                            if ((source.Sheet.GetCell(source.Range.Top+r+1,source.Range.Left+c)?.Value ?? CellValue.Blank)!=cache.Cache.Records[r][c]) pivot=pivot.WithStale(true);
                        }
                    }
                    if (!names.Add(pivot.Name) || pivots.Count>=16 || pivots.Any(p=>Overlaps(p.Range,pivot.Range)))
                        throw new InvalidDataException("透视表重名、重叠或超过每表 16 个。");
                    if (sheet.Data.Tables.Any(t=>Overlaps(t.Range,pivot.Range)) || sheet.MergedCells.Select(CellRectangle.Parse).Any(r=>Overlaps(r,pivot.Range)))
                        throw new InvalidDataException("透视结果与表格或合并区域重叠。");
                    pivots.Add(pivot);
                }
                catch (NotSupportedException error)
                { package.Loss("透视定义/缓存保留在原包，尚不能刷新："+error.Message,"worksheet/pivotTable",part,range:(string?)definition.Element(Main+"location")?.Attribute("ref"),sheetName:sheet.Name); }
            }
            if (pivots.Count>0) result[index]=sheet.WithData(sheet.Data with { Pivots=pivots.ToImmutable() });
        }
        return book.WithSheets(result);
    }

    private static void Shape(XElement element,string children,string attributes)
    {
        var childNames=children.Split(' ',StringSplitOptions.RemoveEmptyEntries); var names=attributes.Split(' ',StringSplitOptions.RemoveEmptyEntries);
        if (element.Elements().Any(e=>e.Name.Namespace!=Main || !childNames.Contains(e.Name.LocalName)) || element.Attributes().Any(a=>!a.IsNamespaceDeclaration &&
            !(a.Name.NamespaceName.Length==0 && names.Contains(a.Name.LocalName)) && !(a.Name==Rel+"id" && names.Contains("r:id")) &&
            !(a.Name.LocalName=="uid" && a.Name.NamespaceName=="http://schemas.microsoft.com/office/spreadsheetml/2014/revision") &&
            a.Name!=XName.Get("Ignorable","http://schemas.openxmlformats.org/markup-compatibility/2006")))
            throw new NotSupportedException("含当前未建模的透视属性或子元素。");
        if (element.Elements().GroupBy(e=>e.Name).Any(g=>g.Count()>1) && element.Name.LocalName is "pivotTableDefinition" or "pivotCacheDefinition" or "cacheField")
            throw new InvalidDataException("透视属性集合重复。");
    }

    private static XElement[] Items(XElement parent,string child,int limit)
    {
        var elements=parent.Elements(Main+child).Take(limit+1).ToArray();
        if (elements.Length>limit || parent.Elements().Count()!=elements.Length || parent.Attribute("count") is { } count && Int(count.Value,limit)!=elements.Length)
            throw new InvalidDataException("透视集合元素数量不一致或超限。");
        return elements;
    }
    private static XElement RequiredChild(XElement parent,string name) => parent.Element(Main+name) ?? throw new InvalidDataException("透视部件缺少 "+name);

    private static void Extensions(XElement root,bool cache)
    {
        if (root.Element(Main+"extLst") is not { } list) return;
        Shape(list,"ext","");
        foreach (var ext in list.Elements())
        {
            var contents=ext.Elements().ToArray();
            if (contents.Length!=1) throw new NotSupportedException("未知透视扩展。");
            var value=contents[0]; var uri=Attr(ext,"uri").ToUpperInvariant();
            var supported=cache ? uri switch
            {
                "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}"=>value.Name==X14+"pivotCacheDefinition" && !value.HasElements && !value.Attributes().Any(a=>!a.IsNamespaceDeclaration),
                "{F4336242-56EE-4DD3-A5F7-E900642BF6E9}"=>value.Name==Clean+"dataSourceClean" && value.Value is "0" or "1" && !value.HasElements,
                _=>false
            } : uri switch
            {
                "{962EF5D1-5CA2-4C93-8EF4-DBF5C05439D2}"=>value.Name==X14+"pivotTableDefinition" && (string?)value.Attribute("hideValuesRow")=="1" &&
                    !value.HasElements && value.Attributes().All(a=>a.IsNamespaceDeclaration || a.Name=="hideValuesRow"),
                "{747A6164-185A-40DC-8AA5-F01512510D54}"=>value.Name==Layout+"pivotTableDefinition16" && !value.HasElements && !value.Attributes().Any(a=>!a.IsNamespaceDeclaration),
                _=>false
            };
            if (!supported) throw new NotSupportedException("未知透视扩展或非默认布局。");
        }
    }

    private static (PivotSource,PivotCache,ImmutableArray<ImmutableArray<CellValue>>,bool) ReadCache(XlsxPackage package,string part,WorkbookSnapshot book)
    {
        package.RequireType(part,Type("pivotCacheDefinition")); var root=package.XmlPart(part).Root!;
        if (root.Name!=Main+"pivotCacheDefinition") throw new InvalidDataException("无效透视缓存定义。");
        Shape(root,"cacheSource cacheFields extLst","r:id refreshedBy refreshedDate createdVersion refreshedVersion minRefreshableVersion recordCount invalid refreshOnLoad enableRefresh saveData backgroundQuery missingItemsLimit");
        Extensions(root,true);
        if (root.Attribute("saveData") is { } save && !Bool(save.Value) || Bool((string?)root.Attribute("backgroundQuery")))
            throw new NotSupportedException("缓存缺失或后台刷新尚不支持。");
        var source=RequiredChild(root,"cacheSource"); Shape(source,"worksheetSource","type");
        if (Attr(source,"type")!="worksheet") throw new NotSupportedException("仅支持工作簿内工作表来源；外部数据不会连接。");
        var reference=RequiredChild(source,"worksheetSource"); Shape(reference,"","sheet ref name");
        PivotSource bound;
        if (reference.Attribute("name") is { } name)
        {
            if (reference.Attribute("ref") is not null) throw new NotSupportedException("同时包含命名来源与区域。");
            var table=book.Sheets.SelectMany(s=>s.Data.Tables.Select(t=>(Sheet:s,Table:t))).FirstOrDefault(p=>p.Table.Name.Equals(Decode(name.Value),StringComparison.OrdinalIgnoreCase));
            if (table.Table is null || !table.Table.HeaderRow) throw new NotSupportedException("命名来源不是带标题的本地表格。");
            bound=new(table.Sheet.Name,table.Table.Range with { Bottom=table.Table.Range.Bottom-(table.Table.TotalsRow ? 1 : 0) },table.Table.Name);
        }
        else bound=new(Decode(Attr(reference,"sheet")),CellRectangle.Parse(Attr(reference,"ref")));
        var fields=RequiredChild(root,"cacheFields"); Shape(fields,"cacheField","count");
        var fieldItems=Items(fields,"cacheField",PivotCache.MaxFields); var names=new List<string>(); var shared=ImmutableArray.CreateBuilder<ImmutableArray<CellValue>>();
        var sharedCount=0; var sharedBytes=0L;
        foreach (var field in fieldItems)
        {
            package.Budget.Check(); Shape(field,"sharedItems","name numFmtId");
            if ((string?)field.Attribute("numFmtId") is not (null or "0")) throw new NotSupportedException("透视字段数字格式尚未建模。");
            names.Add(Decode(Attr(field,"name"))); var items=RequiredChild(field,"sharedItems");
            Shape(items,"m n s b e","count containsBlank containsMixedTypes containsNumber containsInteger containsString containsSemiMixedTypes containsDate containsNonDate minValue maxValue");
            if (Bool((string?)items.Attribute("containsDate"))) throw new NotSupportedException("日期分组暂不支持。");
            var values=items.Elements().Take(PivotCache.MaxRows+1).Select(v=>
            {
                package.Budget.Check(); if (++sharedCount>PivotCache.MaxValues) throw new InvalidDataException("透视共享项累计超过 200,000 项。");
                var value=ReadValue(v,[]); sharedBytes+=32+(value.Kind==CellValueKind.Text ? value.TextValue.Length*2L : 0);
                if (sharedBytes>16*1024*1024) throw new InvalidDataException("透视共享项超过 16 MiB。");
                return value;
            }).ToImmutableArray();
            if (values.Length>PivotCache.MaxRows || items.Attribute("count") is { } count && Int(count.Value,PivotCache.MaxRows)!=values.Length) throw new InvalidDataException("共享透视项数量无效。");
            shared.Add(values);
        }
        package.Budget.Parse(sharedBytes);
        var recordsPart=package.Related(part,Attr(root,Rel+"id"),"pivotCacheRecords"); package.RequireType(recordsPart,Type("pivotCacheRecords"));
        var recordsRoot=package.XmlPart(recordsPart).Root!;
        if (recordsRoot.Name!=Main+"pivotCacheRecords") throw new InvalidDataException("无效透视缓存记录。");
        Shape(recordsRoot,"r","count"); var records=Items(recordsRoot,"r",PivotCache.MaxRows);
        if (root.Attribute("recordCount") is { } total && Int(total.Value,PivotCache.MaxRows)!=records.Length) throw new InvalidDataException("透视缓存记录数量与定义不一致。");
        if ((long)records.Length*names.Count>PivotCache.MaxValues) throw new InvalidDataException("透视缓存超过容量。");
        var valuesList=new List<CellValue[]>();
        foreach (var record in records)
        {
            package.Budget.Check(); Shape(record,"m n s b e x",""); var values=record.Elements().ToArray();
            if (values.Length!=names.Count) throw new InvalidDataException("透视缓存记录列数不一致。");
            valuesList.Add(values.Select((v,i)=>ReadValue(v,shared[i])).ToArray());
        }
        var cache=new PivotCache(names,valuesList); package.Budget.Parse(cache.EstimatedBytes);
        var stale=Bool((string?)root.Attribute("invalid")) || Bool((string?)root.Attribute("refreshOnLoad")) ||
            root.Descendants(Clean+"dataSourceClean").Any(v=>v.Value=="0");
        return (bound,cache,shared.ToImmutable(),stale);
    }

    private static CellValue ReadValue(XElement element,ImmutableArray<CellValue> shared)
    {
        Shape(element,"",element.Name==Main+"m" ? "" : "v");
        return element.Name.LocalName switch
        {
            "m"=>CellValue.Blank,"s"=>CellValue.FromText(Decode(Attr(element,"v"))),"n"=>CellValue.FromNumber(Number(Attr(element,"v"))),
            "b"=>CellValue.FromBoolean(Bool(Attr(element,"v"))),"x" when !shared.IsEmpty=>shared[Int(Attr(element,"v"),shared.Length-1)],
            "e"=>throw new NotSupportedException("含错误项的透视缓存暂仅保留。"),_=>throw new InvalidDataException("透视记录值或共享索引无效。")
        };
    }

    private static WorksheetPivot ReadDefinition(XlsxPackage package,string part,XElement root,WorksheetSnapshot sheet,PivotSource source,PivotCache cache,
        ImmutableArray<ImmutableArray<CellValue>> shared,bool stale)
    {
        Shape(root,"location pivotFields rowFields rowItems colFields colItems dataFields pivotTableStyleInfo extLst",
            "name cacheId dataCaption grandTotalCaption applyNumberFormats applyBorderFormats applyFontFormats applyPatternFormats applyAlignmentFormats applyWidthHeightFormats updatedVersion minRefreshableVersion useAutoFormatting itemPrintTitles createdVersion indent compact compactData multipleFieldFilters rowGrandTotals colGrandTotals");
        Extensions(root,false);
        if ((string?)root.Attribute("compact")!="0" || (string?)root.Attribute("compactData")!="0" ||
            root.Attribute("rowGrandTotals") is { } rg && !Bool(rg.Value) || root.Attribute("colGrandTotals") is { } cg && !Bool(cg.Value) || Bool((string?)root.Attribute("multipleFieldFilters")))
            throw new NotSupportedException("仅支持带行列总计的简单表格布局。");
        var location=RequiredChild(root,"location"); Shape(location,"","ref firstHeaderRow firstDataRow firstDataCol");
        var range=CellRectangle.Parse(Attr(location,"ref"));
        int? Axis(string name,bool required)
        {
            if (root.Element(Main+name) is not { } axis) return required ? throw new InvalidDataException("缺少行字段。") : null;
            Shape(axis,"field","count"); var items=Items(axis,"field",PivotCache.MaxFields);
            if (items.Length!=1) throw new NotSupportedException("仅支持一个行字段与可选的一个列字段。");
            Shape(items[0],"","x"); return Int(Attr(items[0],"x"),cache.Fields.Length-1);
        }
        var row=Axis("rowFields",true)!.Value; var column=Axis("colFields",false);
        var data=RequiredChild(root,"dataFields"); Shape(data,"dataField","count"); var measures=Items(data,"dataField",128);
        if (measures.Length!=1) throw new NotSupportedException("仅支持单个求和或计数度量。");
        var measure=measures[0]; Shape(measure,"","name fld subtotal baseField baseItem");
        var aggregation=(string?)measure.Attribute("subtotal") switch {null or "sum"=>PivotAggregation.Sum,"count"=>PivotAggregation.Count,_=>throw new NotSupportedException("此汇总函数尚未支持。")};
        var field=Int(Attr(measure,"fld"),cache.Fields.Length-1); var caption=Decode(Attr(measure,"name"));
        var grand=(string?)root.Attribute("grandTotalCaption") ?? (sheet.GetCell(range.Bottom,range.Left)?.Value is {Kind:CellValueKind.Text} label ? label.TextValue : "总计");
        var output=PivotReport.Build(cache,new(range.Top,range.Left),row,column,field,aggregation,caption,grand);
        if (output.Range!=range) throw new NotSupportedException("透视显示区域不属于当前基础布局。");
        if (Int(Attr(location,"firstDataCol"),128)!=1 || Int(Attr(location,"firstDataRow"),128)!=(column is null ? 1 : 2) ||
            Int(Attr(location,"firstHeaderRow"),128)!=(column is null ? 0 : 1)) throw new NotSupportedException("透视表头位置未建模。");
        var pivots=RequiredChild(root,"pivotFields"); Shape(pivots,"pivotField","count"); var fields=Items(pivots,"pivotField",PivotCache.MaxFields);
        if (fields.Length!=cache.Fields.Length) throw new InvalidDataException("透视字段与缓存不一致。");
        for (var i=0;i<fields.Length;i++)
        {
            var f=fields[i]; Shape(f,"items","axis compact outline showAll defaultSubtotal dataField");
            var axis=(string?)f.Attribute("axis");
            if (axis!=(i==row ? "axisRow" : i==column ? "axisCol" : null) || Bool((string?)f.Attribute("dataField"))!=(i==field)) throw new NotSupportedException("透视字段角色不一致。");
            if (axis is null) { if (f.HasElements) throw new NotSupportedException("隐藏维度项尚未建模。"); continue; }
            if ((string?)f.Attribute("defaultSubtotal")!="0" || Bool((string?)f.Attribute("outline"))) throw new NotSupportedException("层次/小计布局尚未支持。");
            var entries=RequiredChild(f,"items"); Shape(entries,"item","count"); var items=Items(entries,"item",PivotCache.MaxRows);
            var expected=i==row ? output.RowItems : output.ColumnItems;
            if (items.Length!=expected.Length) throw new NotSupportedException("透视筛选或缺失项尚未支持。");
            for (var n=0;n<items.Length;n++)
            {
                Shape(items[n],"","x"); var x=Int(Attr(items[n],"x"),shared[i].Length-1);
                if (shared[i][x]!=expected[n]) throw new NotSupportedException("手动分组、排序或不同项顺序尚未支持。");
            }
        }
        foreach (var name in new[] {"rowItems","colItems"})
        {
            var list=RequiredChild(root,name); Shape(list,"i","count"); var items=Items(list,"i",PivotCache.MaxRows+1);
            var count=name=="rowItems" ? output.RowItems.Length : column is null ? 0 : output.ColumnItems.Length;
            if (items.Length!=count+1) throw new NotSupportedException("透视显示项数量未建模。");
            for (var n=0;n<items.Length;n++)
            {
                Shape(items[n],"x","t"); var x=items[n].Elements(Main+"x").ToArray();
                if (x.Length!=1) throw new NotSupportedException("多层显示项尚未支持。"); Shape(x[0],"","v");
                var type=(string?)items[n].Attribute("t");
                if (n<count && (type is not (null or "data") || Int((string?)x[0].Attribute("v") ?? "0",PivotCache.MaxRows)!=n) ||
                    n==count && count>0 && type!="grand") throw new NotSupportedException("透视显示项排序/总计未建模。");
            }
        }
        if (root.Element(Main+"pivotTableStyleInfo") is { } style)
        {
            Shape(style,"","name showRowHeaders showColHeaders showRowStripes showColStripes showLastColumn");
            if (!string.IsNullOrEmpty((string?)style.Attribute("name"))) throw new NotSupportedException("命名透视样式尚未建模。");
        }
        for (var r=range.Top;r<=range.Bottom;r++)
        for (var c=range.Left;c<=range.Right;c++)
        {
            package.Budget.Check(); var cell=sheet.GetCell(r,c);
            if (cell?.Content.Kind==CellContentKind.Formula || cell?.SpillOwner is not null || (cell?.Value ?? CellValue.Blank)!=output.Cells.GetValueOrDefault(new(r,c)))
                throw new NotSupportedException("透视显示值不匹配已建模缓存，保留原包。");
        }
        return new(Decode(Attr(root,"name")),source,range,cache,row,column,field,aggregation,caption,grand,stale);
    }
}
