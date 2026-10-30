using System.IO;
using System.Text.Json;
using System.Collections.Immutable;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>View projection of persisted data state; isolated legacy views remain available for grid diagnostics.</summary>
internal sealed class SheetViewState
{
    private GridRange? _legacyRange;
    private WorksheetData? _persistedData;
    private bool _persistedFilter;
    public void Synchronize(WorkbookSnapshot snapshot, SheetId id)
    {
        var data = snapshot.GetSheet(id).Data;
        if (ReferenceEquals(_persistedData, data)) return;
        _persistedData = data; _persistedFilter = data.Filter is not null;
        _legacyRange = null; _evaluatedSnapshot = null; FilterError = null; FilterUsesSourceCaches = false; _hidden = [];
        Region = data.Filter is { } filter ? new(id.Value, id, new(filter.Range.Top, filter.Range.Left, filter.Range.Bottom, filter.Range.Right)) : null;
        Filters = data.Filter?.Columns.ToImmutableDictionary(p => p.Key, p => PersistentFiltering.Restore(p.Value)) ?? ImmutableDictionary<int, ColumnFilter>.Empty;
        SortKeys = data.Sort?.Keys.ToArray() ?? [];
    }

    private WorkbookSnapshot? _evaluatedSnapshot;
    private ImmutableHashSet<int> _hidden = [];
    public DataRegion? Region { get; private set; }
    public ImmutableDictionary<int, ColumnFilter> Filters { get; private set; } = ImmutableDictionary<int, ColumnFilter>.Empty;
    public IReadOnlyList<SortKey> SortKeys { get; set; } = [];
    public string? FilterError { get; private set; }
    public bool FilterUsesSourceCaches { get; private set; }
    public bool IsFilterButtonVisible(int column) => Filters.GetValueOrDefault(column)?.Stored?.ShowButton != false;
    public string HeaderMarker(int column)
    {
        var index = -1;
        for (var i = 0; i < SortKeys.Count; i++) if (SortKeys[i].Column == column) { index = i; break; }
        var sort = index < 0 ? "" : (SortKeys[index].Descending ? "↓" : "↑") + (index + 1);
        return (Filters.ContainsKey(column) ? "●" : "") + (sort.Length > 0 ? sort : "▾");
    }
    public GridRange? FilterRange
    {
        get => Region?.Range ?? _legacyRange;
        set { _legacyRange = value; Region = null; Filters = Filters.Clear(); SortKeys = []; _evaluatedSnapshot = null; FilterError = null; _hidden = []; }
    }
    public int FilterColumn { get; set; }
    public string FilterText { get; set; } = "";
    public int FrozenRows { get; set; }
    public int FrozenColumns { get; set; }
    public void SetRegion(DataRegion region)
    {
        region.Validate(); _persistedFilter = false; Region = region; _legacyRange = null; Filters = Filters.Clear(); SortKeys = [];
        _evaluatedSnapshot = null; FilterError = null; _hidden = [];
    }
    public void Apply(WorkbookSnapshot snapshot, int column, ColumnFilter? filter)
    {
        var region = Region ?? throw new InvalidOperationException("请先建立筛选区域。");
        if (column < region.Range.Left || column > region.Range.Right) throw new ArgumentException("筛选列不在区域内。");
        var next = filter is null ? Filters.Remove(column) : Filters.SetItem(column, filter);
        var hidden = DataFilterEngine.Evaluate(snapshot, region, next);
        Filters = next; _hidden = hidden; _evaluatedSnapshot = snapshot; FilterError = null;
    }
    public void ClearConditions() { Filters = Filters.Clear(); _hidden = []; _evaluatedSnapshot = null; FilterError = null; }
    public void RefreshFilter(WorkbookSnapshot snapshot)
    {
        if (Region is null || ReferenceEquals(_evaluatedSnapshot, snapshot)) return;
        try
        {
            var allowSourceCaches = WorksheetFilterEngine.CanUseSourceCaches(snapshot);
            _hidden = _persistedFilter ? WorksheetFilterEngine.Evaluate(snapshot, Region.SheetId, allowSourceCaches) : DataFilterEngine.Evaluate(snapshot, Region, Filters);
            FilterUsesSourceCaches = _persistedFilter && allowSourceCaches && snapshot.GetSheet(Region.SheetId).Cells.Any(p =>
                p.Key.Row > Region.Range.Top && p.Key.Row <= Region.Range.Bottom && Filters.ContainsKey(p.Key.Column) && p.Value.Content.Kind == CellContentKind.Formula &&
                (p.Value.Calculation?.InputRevision != snapshot.Revision || p.Value.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)));
            FilterError = null;
        }
        catch (InvalidOperationException ex) { _hidden = []; FilterError = ex.Message; FilterUsesSourceCaches = false; }
        _evaluatedSnapshot = snapshot;
    }
    public bool IsFilteredOut(WorkbookSnapshot snapshot, int sheetIndex, int row)
    {
        if (Region is not null)
        {
            if (snapshot.Sheets[sheetIndex].Id != Region.SheetId) return false;
            RefreshFilter(snapshot);
            return _hidden.Contains(row);
        }
        if (FilterRange is not { } range || row <= range.Top || row > range.Bottom) return false;
        var sheet = snapshot.Sheets[sheetIndex];
        var display = CellDisplayFormatter.Format(sheet.GetCell(row, FilterColumn)?.Value ?? CellValue.Blank,
            snapshot.Styles[sheet.GetStyleId(row, FilterColumn)], snapshot.DateSystem);
        return !display.Contains(FilterText, StringComparison.OrdinalIgnoreCase);
    }
    public void RequireVisibleTargets(WorkbookSnapshot snapshot, int sheetIndex, IEnumerable<int> rows)
    {
        RequireCurrentFilter(snapshot);
        if (rows.Any(row => IsFilteredOut(snapshot, sheetIndex, row)))
            throw new InvalidOperationException("目标含筛选隐藏行。请清除筛选后粘贴；本次没有写入。");
    }
    public void RequireCurrentFilter(WorkbookSnapshot snapshot)
    {
        RefreshFilter(snapshot);
        if (FilterError is not null) throw new InvalidOperationException(FilterError);
    }
    public Func<int, bool> CaptureRowFilter(WorkbookSnapshot snapshot, int sheetIndex)
    {
        RequireCurrentFilter(snapshot);
        if (Region is not null)
        {
            var hidden = _hidden; var matching = Region.SheetId == snapshot.Sheets[sheetIndex].Id;
            return row => matching && hidden.Contains(row);
        }
        var range = _legacyRange; var column = FilterColumn; var text = FilterText; var sheet = snapshot.Sheets[sheetIndex];
        return row => range is { } r && row > r.Top && row <= r.Bottom && !CellDisplayFormatter.Format(sheet.GetCell(row, column)?.Value ?? CellValue.Blank,
            snapshot.Styles[sheet.GetStyleId(row, column)], snapshot.DateSystem).Contains(text, StringComparison.OrdinalIgnoreCase);
    }
}

internal sealed class RecentFileStore(string path)
{
    public const int Limit = 10;
    public IReadOnlyList<string> Paths { get; private set; } = [];
    public string? Error { get; private set; }
    public void Load()
    {
        try
        {
            if (!File.Exists(path)) return;
            if (new FileInfo(path).Length > 64 * 1024) throw new InvalidDataException("最近文件配置过大。");
            Paths = Normalize(JsonSerializer.Deserialize<string[]>(File.ReadAllText(path)) ?? []);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or ArgumentException)
        { Paths = []; Error = ex.Message; }
    }
    private static string[] Normalize(IEnumerable<string> paths) => paths.Where(p => !string.IsNullOrWhiteSpace(p) && Path.IsPathFullyQualified(p)
        && Path.GetExtension(p).ToLowerInvariant() is ".xlsx" or ".xlsq").Select(Path.GetFullPath)
        .Distinct(StringComparer.OrdinalIgnoreCase).Take(Limit).ToArray();
    public void Remember(string file) { Paths = Normalize(new[] { file }.Concat(Paths)); Save(); }
    public void Clear() { Paths = []; Save(); }
    private void Save()
    {
        var temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            File.WriteAllText(temp, JsonSerializer.Serialize(Paths));
            File.Move(temp, path, true);
            Error = null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        { Error = ex.Message; }
        finally { try { if (File.Exists(temp)) File.Delete(temp); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }
}
