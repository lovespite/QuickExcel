using System.Collections.Immutable;

namespace QuickExcel.Domain;

public enum PivotAggregation { Sum, Count, Average, Min, Max }

/// <summary>Workbook-local source. Range includes headers; table sources resolve the current table extent.</summary>
public sealed record PivotSource(string Sheet, CellRectangle Range, string? TableName = null);

/// <summary>Typed immutable source records, distinct from the displayed report cells.</summary>
public sealed class PivotCache
{
    public const int MaxValues = 200_000;
    public const int MaxFields = 128;
    public const int MaxRows = 50_000;
    public PivotCache(IEnumerable<string> fields, IEnumerable<IEnumerable<CellValue>> records)
    {
        Fields = fields.Take(MaxFields + 1).ToImmutableArray();
        if (Fields.Length is < 1 or > MaxFields || Fields.Any(string.IsNullOrWhiteSpace) ||
            Fields.Any(f => f.Length > 255) || Fields.Distinct(StringComparer.OrdinalIgnoreCase).Count() != Fields.Length)
            throw new ArgumentException("透视源须有不重复的非空文本列名，最多 128 列，每个列名最多 255 字符。");
        var rows = ImmutableArray.CreateBuilder<ImmutableArray<CellValue>>(); var bytes = 1024L + Fields.Sum(f => 32L + f.Length * 2L);
        foreach (var record in records)
        {
            if (rows.Count >= MaxRows || (long)(rows.Count + 1) * Fields.Length > MaxValues) throw new ArgumentException("透视缓存超过 50,000 行或 200,000 个值。");
            var values = record.Take(Fields.Length + 1).ToImmutableArray();
            if (values.Length != Fields.Length) throw new ArgumentException("透视缓存记录列数不一致。");
            foreach (var value in values)
            {
                if (value.Kind == CellValueKind.Text) WorkbookLimits.ValidateText(value.TextValue);
                bytes += 32 + (value.Kind == CellValueKind.Text ? value.TextValue.Length * 2L : 0);
            }
            if (bytes > 16 * 1024 * 1024) throw new ArgumentException("透视缓存超过 16 MiB。");
            rows.Add(values);
        }
        Records = rows.ToImmutable(); EstimatedBytes = bytes;
    }
    public ImmutableArray<string> Fields { get; }
    public ImmutableArray<ImmutableArray<CellValue>> Records { get; }
    public long EstimatedBytes { get; }
    public bool ContentEquals(PivotCache other) => Fields.SequenceEqual(other.Fields) && Records.Length == other.Records.Length &&
        Records.Zip(other.Records).All(p => p.First.SequenceEqual(p.Second));
}

/// <summary>Basic tabular report: one row field, optional column field, one measure, both grand totals.</summary>
public sealed class WorksheetPivot
{
    public WorksheetPivot(string name, PivotSource source, CellRectangle range, PivotCache cache, int rowField, int? columnField,
        int dataField, PivotAggregation aggregation, string dataCaption, string grandCaption = "总计", bool stale = false)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || name.Any(char.IsWhiteSpace) || name.IndexOfAny(['[',']','!']) >= 0)
            throw new ArgumentException("透视表名称无效。");
        WorkbookLimits.ValidateSheetName(source.Sheet);
        foreach (var r in new[] { source.Range,range })
        {
            WorkbookLimits.ValidatePosition(r.Top,r.Left); WorkbookLimits.ValidatePosition(r.Bottom,r.Right);
            if (r.Top > r.Bottom || r.Left > r.Right) throw new ArgumentException("透视表或源区域无效。");
        }
        if (source.Range.Rows < 2 || range.Count > 20_000 ||
            rowField < 0 || rowField >= cache.Fields.Length || columnField is { } c && (c < 0 || c >= cache.Fields.Length || c == rowField) ||
            dataField < 0 || dataField >= cache.Fields.Length || !Enum.IsDefined(aggregation)) throw new ArgumentException("透视表字段、来源或输出大小无效。");
        if (string.IsNullOrWhiteSpace(dataCaption) || dataCaption.Length > 255 || string.IsNullOrWhiteSpace(grandCaption) || grandCaption.Length > 255)
            throw new ArgumentException("透视表汇总标题无效。");
        Name=name; Source=source; Range=range; Cache=cache; RowField=rowField; ColumnField=columnField;
        DataField=dataField; Aggregation=aggregation; DataCaption=dataCaption; GrandCaption=grandCaption; Stale=stale;
    }
    public string Name { get; }
    public PivotSource Source { get; }
    public CellRectangle Range { get; }
    public PivotCache Cache { get; }
    public int RowField { get; }
    public int? ColumnField { get; }
    public int DataField { get; }
    public PivotAggregation Aggregation { get; }
    public string DataCaption { get; }
    public string GrandCaption { get; }
    public bool Stale { get; }
    public long EstimatedBytes => 2048 + Cache.EstimatedBytes;
    public WorksheetPivot WithStale(bool stale) => stale == Stale ? this : new(Name,Source,Range,Cache,RowField,ColumnField,DataField,Aggregation,DataCaption,GrandCaption,stale);
    public bool ContentEquals(WorksheetPivot other) => Name == other.Name && Source == other.Source && Range == other.Range && RowField == other.RowField &&
        ColumnField == other.ColumnField && DataField == other.DataField && Aggregation == other.Aggregation && DataCaption == other.DataCaption &&
        GrandCaption == other.GrandCaption && Stale == other.Stale && Cache.ContentEquals(other.Cache);
}
