using System.IO;
using System.Text.Json;
using System.Collections.Immutable;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

/// <summary>View projection of persisted data state; isolated legacy views remain available for grid diagnostics.</summary>
public sealed class SheetViewState
{
    private GridRange? _legacyRange;
    private WorksheetData? _persistedData;
    private bool _persistedFilter;
    private SheetId _sheetId;
    public int? TableId { get; private set; }
    public void Synchronize(WorkbookSnapshot snapshot, SheetId id, CellPosition? current = null)
    {
        var data = snapshot.GetSheet(id).Data;
        var table = current is { } position ? data.Tables.FirstOrDefault(t => t.Range.Contains(position.Row, position.Column)) : null;
        if (ReferenceEquals(_persistedData, data) && TableId == table?.Id) return;
        _sheetId = id; TableId = table?.Id;
        _persistedData = data; _persistedFilter = data.AllFilters.Any();
        _legacyRange = null; _evaluatedSnapshot = null; FilterError = null; FilterUsesSourceCaches = false; _hidden = [];
        var filter = table is null ? data.Filter : table.Filter;
        Region = filter is not null ? new(id.Value, id, new(filter.Range.Top, filter.Range.Left, filter.Range.Bottom, filter.Range.Right)) : null;
        Filters = filter?.Columns.ToImmutableDictionary(p => p.Key, p => PersistentFiltering.Restore(p.Value)) ?? ImmutableDictionary<int, ColumnFilter>.Empty;
        SortKeys = (table is null ? data.Sort : table.Sort)?.Keys.ToArray() ?? [];
    }

    public IEnumerable<(GridRange Range, WorksheetAutoFilter Filter, WorksheetSortState? Sort)> HeaderRegions()
    {
        if (_persistedData is null) yield break;
        if (_persistedData.Filter is { } filter) yield return (new(filter.Range.Top, filter.Range.Left, filter.Range.Bottom, filter.Range.Right), filter, _persistedData.Sort);
        foreach (var table in _persistedData.Tables)
            if (table.Filter is { } tableFilter) yield return (new(tableFilter.Range.Top, tableFilter.Range.Left, tableFilter.Range.Bottom, tableFilter.Range.Right), tableFilter, table.Sort);
    }
    public static string HeaderMarker(int column, WorksheetAutoFilter filter, WorksheetSortState? sort)
    {
        var key = sort?.Keys.FirstOrDefault(k => k.Column == column); var index = key is null ? -1 : sort!.Keys.IndexOf(key);
        return (filter.Columns.ContainsKey(column) ? "●" : "") + (key is null ? "▾" : (key.Descending ? "↓" : "↑") + (index + 1));
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
        var region = Region ?? throw new InvalidOperationException(UiText.Get("请先建立筛选区域。"));
        if (column < region.Range.Left || column > region.Range.Right) throw new ArgumentException(UiText.Get("筛选列不在区域内。"));
        var next = filter is null ? Filters.Remove(column) : Filters.SetItem(column, filter);
        var hidden = DataFilterEngine.Evaluate(snapshot, region, next);
        Filters = next; _hidden = hidden; _evaluatedSnapshot = snapshot; FilterError = null;
    }
    public void ClearConditions() { Filters = Filters.Clear(); _hidden = []; _evaluatedSnapshot = null; FilterError = null; }
    public void RefreshFilter(WorkbookSnapshot snapshot)
    {
        if (!_persistedFilter && Region is null || ReferenceEquals(_evaluatedSnapshot, snapshot)) return;
        try
        {
            var allowSourceCaches = WorksheetFilterEngine.CanUseSourceCaches(snapshot);
            _hidden = _persistedFilter ? WorksheetFilterEngine.Evaluate(snapshot, _sheetId, allowSourceCaches) : DataFilterEngine.Evaluate(snapshot, Region!, Filters);
            FilterUsesSourceCaches = _persistedFilter && allowSourceCaches && snapshot.GetSheet(_sheetId).Cells.Any(p =>
                _persistedData!.AllFilters.Any(f => p.Key.Row > f.Range.Top && p.Key.Row <= f.Range.Bottom && f.Columns.ContainsKey(p.Key.Column)) && p.Value.Content.Kind == CellContentKind.Formula &&
                (p.Value.Calculation?.InputRevision != snapshot.Revision || p.Value.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)));
            FilterError = null;
        }
        catch (InvalidOperationException ex) { _hidden = []; FilterError = ex.Message; FilterUsesSourceCaches = false; }
        _evaluatedSnapshot = snapshot;
    }
    public bool IsFilteredOut(WorkbookSnapshot snapshot, int sheetIndex, int row)
    {
        if (_persistedFilter) { if (snapshot.Sheets[sheetIndex].Id != _sheetId) return false; RefreshFilter(snapshot); return _hidden.Contains(row); }
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
            throw new InvalidOperationException(UiText.Get("目标含筛选隐藏行。请清除筛选后粘贴；本次没有写入。"));
    }
    public void RequireCurrentFilter(WorkbookSnapshot snapshot)
    {
        RefreshFilter(snapshot);
        if (FilterError is not null) throw new InvalidOperationException(FilterError);
    }
    public Func<int, bool> CaptureRowFilter(WorkbookSnapshot snapshot, int sheetIndex)
    {
        RequireCurrentFilter(snapshot);
        if (_persistedFilter || Region is not null)
        {
            var hidden = _hidden; var matching = (_persistedFilter ? _sheetId : Region!.SheetId) == snapshot.Sheets[sheetIndex].Id;
            return row => matching && hidden.Contains(row);
        }
        var range = _legacyRange; var column = FilterColumn; var text = FilterText; var sheet = snapshot.Sheets[sheetIndex];
        return row => range is { } r && row > r.Top && row <= r.Bottom && !CellDisplayFormatter.Format(sheet.GetCell(row, column)?.Value ?? CellValue.Blank,
            snapshot.Styles[sheet.GetStyleId(row, column)], snapshot.DateSystem).Contains(text, StringComparison.OrdinalIgnoreCase);
    }
}

public sealed record RecentFileEntry(string Path, bool IsPinned = false);

public sealed class RecentFileStore(string path)
{
    public const int DefaultLimit = 100;
    public const int MinimumLimit = 1;
    public const int MaximumLimit = 1000;
    private const int MaximumSettingsBytes = 8 * 1024 * 1024;
    private sealed record Settings(int Version, int Limit, RecentFileEntry[] Entries);
    private static readonly StringComparer PathComparer = OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
    public int Limit { get; private set; } = DefaultLimit;
    public IReadOnlyList<RecentFileEntry> Entries { get; private set; } = [];
    public IReadOnlyList<string> Paths { get; private set; } = [];
    public string? Error { get; private set; }
    public void Load()
    {
        try
        {
            if (!File.Exists(path)) { SetState([], DefaultLimit); Error = null; return; }
            if (new FileInfo(path).Length > MaximumSettingsBytes) throw new InvalidDataException(UiText.Get("最近文件配置过大。"));
            using var json = JsonDocument.Parse(File.ReadAllText(path));
            if (json.RootElement.ValueKind == JsonValueKind.Array)
                SetState(Normalize((json.Deserialize<string[]>() ?? []).Select(p => new RecentFileEntry(p)), DefaultLimit), DefaultLimit);
            else
            {
                var settings = json.Deserialize<Settings>();
                if (settings is null || settings.Version != 1 || settings.Limit is < MinimumLimit or > MaximumLimit || settings.Entries is null)
                    throw new InvalidDataException(UiText.Get("无效的历史文件设置。"));
                SetState(Normalize(settings.Entries, settings.Limit), settings.Limit);
            }
            Error = null;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or JsonException or ArgumentException)
        { SetState([], DefaultLimit); Error = ex.Message; }
    }
    private void SetState(RecentFileEntry[] entries, int limit)
    { Entries = entries; Paths = entries.Select(e => e.Path).ToArray(); Limit = limit; }
    private static RecentFileEntry[] Normalize(IEnumerable<RecentFileEntry> entries, int limit)
    {
        var unique = entries.Where(e => e is not null && !string.IsNullOrWhiteSpace(e.Path) && Path.IsPathFullyQualified(e.Path)
            && Path.GetExtension(e.Path).ToLowerInvariant() is ".xlsx" or ".xlsq" or ".csv")
            .Select(e => e with { Path = Path.GetFullPath(e.Path) }).OrderByDescending(e => e.IsPinned)
            .DistinctBy(e => e.Path, PathComparer).ToArray();
        return unique.Where(e => e.IsPinned).Concat(unique.Where(e => !e.IsPinned).Take(limit)).ToArray();
    }
    public void Remember(string file) => Save(current => current.Any(e => e.IsPinned && PathComparer.Equals(e.Path, file))
        ? current : new[] { new RecentFileEntry(file) }.Concat(current));
    public void SetPinned(string file, bool pinned) => Save(current => current.FirstOrDefault(e => PathComparer.Equals(e.Path, file)) is { } entry
        ? new[] { entry with { IsPinned = pinned } }.Concat(current.Where(e => !PathComparer.Equals(e.Path, file))) : current);
    public void ChangeLimit(int limit)
    {
        if (limit is < MinimumLimit or > MaximumLimit) throw new ArgumentOutOfRangeException(nameof(limit));
        Save(current => current, limit);
    }
    public void Remove(string file) => Save(current => current.Where(e => !PathComparer.Equals(e.Path, file)));
    public void Clear() => Save(_ => []);
    private void Save(Func<IReadOnlyList<RecentFileEntry>, IEnumerable<RecentFileEntry>> update, int? limit = null)
    {
        var temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            // Associated files can open in separate processes. Read the latest list
            // under the same lease as the replacement so a removed entry stays removed.
            using var lease = AcquireLease();
            Load();
            if (Error is not null) return;
            var nextLimit = limit ?? Limit;
            var entries = Normalize(update(Entries), nextLimit);
            var bytes = JsonSerializer.SerializeToUtf8Bytes(new Settings(1, nextLimit, entries));
            if (bytes.Length > MaximumSettingsBytes) throw new IOException(UiText.Get("最近文件配置过大。"));
            File.WriteAllBytes(temp, bytes);
            File.Move(temp, path, true);
            SetState(entries, nextLimit); Error = null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        { Error = ex.Message; }
        finally { try { if (File.Exists(temp)) File.Delete(temp); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }

    private FileStream AcquireLease()
    {
        for (var attempt = 0; ; attempt++)
        {
            try { return new FileStream(path + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) when (attempt < 10) { Thread.Sleep(10); }
        }
    }
}
