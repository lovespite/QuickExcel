using System.IO;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>Session-only sheet view state. Addresses are always model coordinates, never filtered indices.</summary>
internal sealed class SheetViewState
{
    public GridRange? FilterRange { get; set; }
    public int FilterColumn { get; set; }
    public string FilterText { get; set; } = "";
    public int FrozenRows { get; set; }
    public int FrozenColumns { get; set; }
    public bool IsFilteredOut(WorkbookSnapshot snapshot, int sheetIndex, int row)
    {
        if (FilterRange is not { } range || row <= range.Top || row > range.Bottom) return false;
        var sheet = snapshot.Sheets[sheetIndex];
        var display = CellDisplayFormatter.Format(sheet.GetCell(row, FilterColumn)?.Value ?? CellValue.Blank,
            snapshot.Styles[sheet.GetStyleId(row, FilterColumn)], snapshot.DateSystem);
        return !display.Contains(FilterText, StringComparison.OrdinalIgnoreCase);
    }
    public void RequireVisibleTargets(WorkbookSnapshot snapshot, int sheetIndex, IEnumerable<int> rows)
    {
        if (rows.Any(row => IsFilteredOut(snapshot, sheetIndex, row)))
            throw new InvalidOperationException("目标含筛选隐藏行。请清除筛选后粘贴；本次没有写入。");
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
