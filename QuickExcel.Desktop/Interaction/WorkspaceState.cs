using System.IO;
using System.Text.Json;
using System.Collections.Immutable;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>Session-only sheet view state. Addresses are always model coordinates, never filtered indices.</summary>
internal sealed class SheetViewState
{
    private GridRange? _legacyRange;
    private WorkbookSnapshot? _evaluatedSnapshot;
    private ImmutableHashSet<int> _hidden = [];
    public DataRegion? Region { get; private set; }
    public ImmutableDictionary<int, ColumnFilter> Filters { get; private set; } = ImmutableDictionary<int, ColumnFilter>.Empty;
    public IReadOnlyList<SortKey> SortKeys { get; set; } = [];
    public string? FilterError { get; private set; }
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
        region.Validate(); Region = region; _legacyRange = null; Filters = Filters.Clear(); SortKeys = [];
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
        try { _hidden = DataFilterEngine.Evaluate(snapshot, Region, Filters); FilterError = null; }
        catch (InvalidOperationException ex) { _hidden = []; FilterError = ex.Message; }
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
