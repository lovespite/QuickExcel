using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Explicit, bounded local refresh. Never contacts an external source.</summary>
public static class PivotReport
{
    public static bool HasUnmodeledDefinitions(WorkbookSnapshot book) => book.FileState?.CompatibilityDiagnostics.Any(d=>
        d.Feature.Equals("worksheet/pivotTable",StringComparison.OrdinalIgnoreCase) ||
        d.Feature.StartsWith("relationship/pivot",StringComparison.OrdinalIgnoreCase) ||
        d.Feature.StartsWith("part/",StringComparison.OrdinalIgnoreCase) && d.Feature.Contains("pivot",StringComparison.OrdinalIgnoreCase))==true;
    public static (WorksheetSnapshot Sheet, CellRectangle Range) Resolve(WorkbookSnapshot book, PivotSource source)
    {
        var sheet = book.Sheets.FirstOrDefault(s => s.Name.Equals(source.Sheet,StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException("透视表来源工作表不存在。");
        if (source.TableName is null) return (sheet,source.Range);
        var table = sheet.Data.Tables.FirstOrDefault(t => t.Name.Equals(source.TableName,StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException("透视表来源表格不存在。");
        if (!table.HeaderRow) throw new NotSupportedException("透视源表格需要标题行。");
        return (sheet,table.Range with { Bottom = table.Range.Bottom - (table.TotalsRow ? 1 : 0) });
    }

    public static PivotCache Capture(WorkbookSnapshot book, PivotSource source, CancellationToken cancellationToken = default)
    {
        var (sheet,range) = Resolve(book,source);
        if (range.Rows < 2 || range.Rows - 1 > PivotCache.MaxRows || range.Columns > PivotCache.MaxFields || range.Count > PivotCache.MaxValues)
            throw new InvalidOperationException("透视源需要标题和数据行，最多 128 列、50,000 行、200,000 格。");
        if (sheet.MergedCells.Select(CellRectangle.Parse).Any(r => ValidationSafety.Overlaps(r,range)) ||
            sheet.Data.Pivots.Any(p => ValidationSafety.Overlaps(p.Range,range))) throw new NotSupportedException("透视源不能含合并区域或其他透视表结果。");
        CellValue Read(int row,int column)
        {
            cancellationToken.ThrowIfCancellationRequested(); var cell=sheet.GetCell(row,column);
            if ((cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null) &&
                (cell.Calculation?.InputRevision != book.Revision || cell.Calculation?.State is not (CalculationState.Fresh or CalculationState.Error)))
                throw new InvalidOperationException("透视源含旧公式结果，请先完成重算再刷新。");
            return cell?.Value ?? CellValue.Blank;
        }
        var fields=Enumerable.Range(range.Left,range.Columns).Select(c => Read(range.Top,c)).ToArray();
        if (fields.Any(f => f.Kind != CellValueKind.Text)) throw new InvalidOperationException("透视源标题须为非空文本。");
        return new(fields.Select(v => v.TextValue),Enumerable.Range(range.Top+1,range.Rows-1).Select(r =>
            Enumerable.Range(range.Left,range.Columns).Select(c => Read(r,c))));
    }

    public sealed record Result(CellRectangle Range, ImmutableDictionary<CellPosition,CellValue> Cells,
        ImmutableArray<CellValue> RowItems, ImmutableArray<CellValue> ColumnItems);

    public static Result Build(PivotCache cache, CellPosition position, int rowField, int? columnField, int dataField,
        PivotAggregation aggregation, string dataCaption, string grandCaption = "总计", CancellationToken cancellationToken = default)
    {
        if (rowField < 0 || rowField >= cache.Fields.Length || dataField < 0 || dataField >= cache.Fields.Length ||
            columnField is { } index && (index < 0 || index >= cache.Fields.Length || index == rowField) || !Enum.IsDefined(aggregation))
            throw new ArgumentException("透视字段配置无效。");
        var comparer = new KeyComparer();
        var rows = new HashSet<CellValue>(comparer); var columns = new HashSet<CellValue>(comparer);
        var sums = new Dictionary<(int Row,int Column),Aggregate>();
        foreach (var record in cache.Records)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (record[rowField].Kind == CellValueKind.Error || columnField is { } c && record[c].Kind == CellValueKind.Error || record[dataField].Kind == CellValueKind.Error)
                throw new NotSupportedException("含错误值的透视分组或汇总暂不支持刷新；原缓存保持不变。");
            rows.Add(record[rowField]); if (columnField is { } field) columns.Add(record[field]);
        }
        if (columnField is null) columns.Add(CellValue.Blank);
        var rowItems=rows.Order(comparer).ToImmutableArray(); var columnItems=columns.Order(comparer).ToImmutableArray();
        var rowIds=rowItems.Select((v,i)=>(v,i)).ToDictionary(p=>p.v,p=>p.i,comparer);
        var columnIds=columnItems.Select((v,i)=>(v,i)).ToDictionary(p=>p.v,p=>p.i,comparer);
        var headers=columnField is null ? 1 : 2;
        var width=1 + columnItems.Length + (columnField is null ? 0 : 1);
        var height=headers + rowItems.Length + 1;
        if ((long)width * height > 20_000) throw new InvalidOperationException("透视输出超过 20,000 格。");
        var range=new CellRectangle(position.Row,position.Column,checked(position.Row+height-1),checked(position.Column+width-1));
        WorkbookLimits.ValidatePosition(range.Top,range.Left); WorkbookLimits.ValidatePosition(range.Bottom,range.Right);
        void Add(int row,int column,CellValue value)
        {
            if (value.Kind == CellValueKind.Blank) return;
            var current=sums.GetValueOrDefault((row,column));
            var sum=current.Sum + (value.Kind == CellValueKind.Number ? value.NumberValue : 0);
            if (!double.IsFinite(sum)) throw new InvalidOperationException("透视求和超出有限数值范围。");
            sums[(row,column)]=new(sum,current.Count+1);
        }
        foreach (var record in cache.Records)
        {
            cancellationToken.ThrowIfCancellationRequested(); var row=rowIds[record[rowField]];
            var column=columnIds[columnField is { } c ? record[c] : CellValue.Blank]; var value=record[dataField];
            Add(row,column,value); Add(row,-1,value); Add(-1,column,value); Add(-1,-1,value);
        }
        var cells=ImmutableDictionary.CreateBuilder<CellPosition,CellValue>();
        void Put(int row,int column,CellValue value) { if (value.Kind != CellValueKind.Blank) cells[new(position.Row+row,position.Column+column)]=value; }
        CellValue Label(CellValue value) => value.Kind == CellValueKind.Blank ? CellValue.FromText("(空白)") : value;
        if (columnField is { } col)
        {
            Put(0,0,CellValue.FromText(dataCaption)); Put(0,1,CellValue.FromText(cache.Fields[col]));
            for (var i=0;i<columnItems.Length;i++) Put(1,i+1,Label(columnItems[i]));
            Put(1,width-1,CellValue.FromText(grandCaption));
        }
        else Put(0,1,CellValue.FromText(dataCaption));
        Put(headers-1,0,CellValue.FromText(cache.Fields[rowField]));
        for (var i=0;i<rowItems.Length;i++) Put(headers+i,0,Label(rowItems[i]));
        Put(height-1,0,CellValue.FromText(grandCaption));
        foreach (var (key,value) in sums)
        {
            if (columnField is null && key.Column == -1) continue;
            Put(key.Row == -1 ? height-1 : headers+key.Row,key.Column == -1 ? width-1 : key.Column+1,
                CellValue.FromNumber(aggregation == PivotAggregation.Sum ? value.Sum : value.Count));
        }
        return new(range,cells.ToImmutable(),rowItems,columnItems);
    }

    private readonly record struct Aggregate(double Sum,int Count);
    private sealed class KeyComparer : IEqualityComparer<CellValue>, IComparer<CellValue>
    {
        public bool Equals(CellValue x,CellValue y) => x.Kind == CellValueKind.Text && y.Kind == CellValueKind.Text
            ? StringComparer.OrdinalIgnoreCase.Equals(x.TextValue,y.TextValue) : x == y;
        public int GetHashCode(CellValue value) => value.Kind == CellValueKind.Text ? StringComparer.OrdinalIgnoreCase.GetHashCode(value.TextValue) : value.GetHashCode();
        public int Compare(CellValue x,CellValue y) => SortRangeCommand.Compare(x,y,false);
    }
}
