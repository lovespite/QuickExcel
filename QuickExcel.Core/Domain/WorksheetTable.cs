using System.Collections.Immutable;

namespace QuickExcel.Domain;

public sealed record TableColumn(int Id, string Name, string? TotalsFunction = null, string? TotalsLabel = null,
    string? CalculatedFormula = null, string? TotalsFormula = null)
{
    public TableStyleElement? DataStyle { get; init; }
    public TableStyleElement? HeaderStyle { get; init; }
    public TableStyleElement? TotalsStyle { get; init; }
}
public sealed record TableBorder(int Style, string? Color);
public sealed record TableStyleElement(string Type, int Size = 1, bool? Bold = null, string? FontColor = null,
    string? FillColor = null, TableBorder? Left = null, TableBorder? Right = null, TableBorder? Top = null,
    TableBorder? Bottom = null, TableBorder? Horizontal = null, TableBorder? Vertical = null,
    int? NumberFormatId = null, string? NumberFormatCode = null)
{
    internal void Validate()
    {
        if (Type is not ("wholeTable" or "headerRow" or "totalRow" or "firstColumn" or "lastColumn" or "firstRowStripe" or "secondRowStripe" or "firstColumnStripe" or "secondColumnStripe" or
            "firstHeaderCell" or "lastHeaderCell" or "firstTotalCell" or "lastTotalCell") || Size is < 1 or > 9 || NumberFormatId is < 0 or >= 164 || NumberFormatCode?.Length > 32767)
            throw new ArgumentException("表差异样式类型、条纹或数字格式无效。");
        var borders = new[] { Left, Right, Top, Bottom, Horizontal, Vertical };
        if (borders.Any(b => b is { Style: < 0 or > 7 }) || borders.Select(b => b?.Color).Concat([FontColor, FillColor]).OfType<string>().Any(c => c.Length != 8 || c.Any(v => !char.IsAsciiHexDigit(v))))
            throw new ArgumentException("表差异样式颜色或边框无效。");
    }
}
public sealed record TableStyleInfo(string? Name = null, bool FirstColumn = false, bool LastColumn = false,
    bool RowStripes = true, bool ColumnStripes = false)
{
    // Resolved display projection; the named style remains the persisted identity.
    public ImmutableArray<TableStyleElement> Elements { get; init; } = [];
    public bool ContentEquals(TableStyleInfo other) => Name == other.Name && FirstColumn == other.FirstColumn &&
        LastColumn == other.LastColumn && RowStripes == other.RowStripes && ColumnStripes == other.ColumnStripes && Elements.SequenceEqual(other.Elements);
}

/// <summary>Independent Excel table object. Its AutoFilter excludes the total row.</summary>
public sealed class WorksheetTable
{
    private static readonly HashSet<string> TotalsFunctions = ["none", "sum", "min", "max", "average", "count", "countNums", "stdDev", "var", "custom"];
    public WorksheetTable(int id, string name, CellRectangle range, IEnumerable<TableColumn> columns,
        bool headerRow = true, bool totalsRow = false, bool totalsRowShown = false, TableStyleInfo? style = null,
        WorksheetAutoFilter? filter = null, WorksheetSortState? sort = null)
    {
        WorkbookLimits.ValidatePosition(range.Top, range.Left); WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        ValidateName(name);
        Columns = columns.Take(WorkbookLimits.MaxColumns + 1).ToImmutableArray();
        if (id < 1 || range.Bottom < range.Top || range.Right < range.Left || Columns.Length != range.Right - range.Left + 1 ||
            range.Bottom - range.Top + 1 < (headerRow ? 1 : 0) + (totalsRow ? 1 : 0)) throw new ArgumentException("表编号、范围或列数无效。");
        var ids = new HashSet<int>(); var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in Columns)
        {
            if (column is null || column.Id < 1 || !ids.Add(column.Id) || string.IsNullOrEmpty(column.Name) || column.Name.Length > 255 || !names.Add(column.Name) ||
                column.TotalsFunction is not null && !TotalsFunctions.Contains(column.TotalsFunction)) throw new ArgumentException("表列编号、名称或汇总函数无效。");
            foreach (var text in new[] { column.TotalsLabel, column.CalculatedFormula, column.TotalsFormula }.OfType<string>()) WorkbookLimits.ValidateText(text);
            foreach (var format in new[] { column.DataStyle, column.HeaderStyle, column.TotalsStyle }.OfType<TableStyleElement>()) format.Validate();
        }
        if (filter is not null && (!headerRow || filter.Range != range with { Bottom = range.Bottom - (totalsRow ? 1 : 0) }))
            throw new ArgumentException("表筛选必须覆盖表头和数据区，不含汇总行。");
        if (sort is not null && (!range.Contains(sort.Range.Top, sort.Range.Left) || !range.Contains(sort.Range.Bottom, sort.Range.Right) ||
            sort.Range.Top < range.Top + (headerRow ? 1 : 0) || sort.Range.Bottom > range.Bottom - (totalsRow ? 1 : 0)))
            throw new ArgumentException("表排序超出数据区。");
        Id = id; Name = name; Range = range; HeaderRow = headerRow; TotalsRow = totalsRow; TotalsRowShown = totalsRowShown;
        Style = style ?? new(); Filter = filter; Sort = sort;
        if (Style.Name?.Length > 255 || Style.Elements.IsDefault || Style.Elements.Any(e => e is null) ||
            Style.Name is null && !Style.Elements.IsEmpty || Style.Elements.Length > 20 || Style.Elements.Select(e => e.Type).Distinct().Count() != Style.Elements.Length)
            throw new ArgumentException("表样式名称或元素数量无效。");
        foreach (var element in Style.Elements)
        { element.Validate(); if (element.NumberFormatId.HasValue) throw new ArgumentException("表区域差异数字格式未支持，请使用列数字格式。"); }
    }
    public static void ValidateName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || !System.Text.RegularExpressions.Regex.IsMatch(name, @"^[\p{L}_\\][\p{L}\p{N}_.\\]*$") ||
            QuickExcel.Formulas.FormulaSyntax.IsCellName(name) || name.Equals("R", StringComparison.OrdinalIgnoreCase) || name.Equals("C", StringComparison.OrdinalIgnoreCase) ||
            System.Text.RegularExpressions.Regex.IsMatch(name, @"^R[0-9]+C[0-9]+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant))
            throw new ArgumentException("表名无效，不能是单元格地址或包含空格。");
    }
    public int Id { get; }
    public string Name { get; }
    public CellRectangle Range { get; }
    public ImmutableArray<TableColumn> Columns { get; }
    public bool HeaderRow { get; }
    public bool TotalsRow { get; }
    public bool TotalsRowShown { get; }
    public TableStyleInfo Style { get; }
    public WorksheetAutoFilter? Filter { get; }
    public WorksheetSortState? Sort { get; }
    public bool HasFormulas => Columns.Any(c => c.CalculatedFormula is not null || c.TotalsFormula is not null);
    public bool ContentEquals(WorksheetTable other) => Id == other.Id && Name == other.Name && Range == other.Range && Columns.SequenceEqual(other.Columns) &&
        HeaderRow == other.HeaderRow && TotalsRow == other.TotalsRow && TotalsRowShown == other.TotalsRowShown && Style.ContentEquals(other.Style) &&
        (Filter is null ? other.Filter is null : other.Filter is not null && Filter.ContentEquals(other.Filter)) &&
        (Sort is null ? other.Sort is null : other.Sort is not null && Sort.ContentEquals(other.Sort));
    public WorksheetTable WithRange(CellRectangle range, WorksheetAutoFilter? filter, WorksheetSortState? sort = null) =>
        new(Id, Name, range, Columns, HeaderRow, TotalsRow, TotalsRowShown, Style, filter, sort);
    public WorksheetTable WithFilter(WorksheetAutoFilter? filter) => new(Id, Name, Range, Columns, HeaderRow, TotalsRow, TotalsRowShown, Style, filter, Sort);
    public WorksheetTable WithSort(WorksheetSortState? sort) => new(Id, Name, Range, Columns, HeaderRow, TotalsRow, TotalsRowShown, Style, Filter, sort);
}
