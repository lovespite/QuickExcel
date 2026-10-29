using System.IO;
using System.Text.Json;

namespace QuickExcel.Desktop;

internal enum SheetNavigationPlacement { Bottom, Left, Right }
internal sealed record DesktopPreferences(SheetNavigationPlacement Sheets = SheetNavigationPlacement.Bottom, int Zoom = 100);

internal sealed class DesktopSettings(string path)
{
    public DesktopPreferences Value { get; private set; } = new();
    public string? Error { get; private set; }
    public void Load()
    {
        try
        {
            if (!File.Exists(path)) return;
            if (new FileInfo(path).Length > 4096) throw new InvalidDataException("设置文件过大。");
            var value = JsonSerializer.Deserialize<DesktopPreferences>(File.ReadAllText(path));
            if (value is null || !Enum.IsDefined(value.Sheets) || value.Zoom is < 75 or > 175) throw new InvalidDataException("无效的界面设置。");
            Value = value; Error = null;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or JsonException)
        { Value = new(); Error = ex.Message; }
    }
    public void Save(DesktopPreferences value)
    {
        if (!Enum.IsDefined(value.Sheets) || value.Zoom is < 75 or > 175) throw new ArgumentException("无效的界面设置。");
        Value = value;
        var temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            File.WriteAllText(temp, JsonSerializer.Serialize(value));
            File.Move(temp, path, true); Error = null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Error = ex.Message; }
        finally { try { if (File.Exists(temp)) File.Delete(temp); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }
}
