using System.IO;
using System.Text.Json;

namespace QuickExcel.Presentation;

public enum SheetNavigationPlacement { Bottom, Left, Right }
public enum DesktopCalculationMode { Automatic, Partial, Manual }
public sealed record DesktopPreferences(SheetNavigationPlacement Sheets = SheetNavigationPlacement.Bottom, int Zoom = 100,
    DesktopCalculationMode Calculation = DesktopCalculationMode.Automatic, string Language = UiText.Chinese)
{
    internal bool IsValid => Enum.IsDefined(Sheets) && Zoom is >= 75 and <= 175 && Enum.IsDefined(Calculation);
}

public sealed class DesktopSettings(string path)
{
    public DesktopPreferences Value { get; private set; } = new();
    public string? Error { get; private set; }
    public void Load()
    {
        try
        {
            if (!File.Exists(path)) return;
            if (new FileInfo(path).Length > 4096) throw new InvalidDataException(UiText.Get("设置文件过大。"));
            var value = JsonSerializer.Deserialize<DesktopPreferences>(File.ReadAllText(path));
            if (value is null || !value.IsValid) throw new InvalidDataException(UiText.Get("无效的界面设置。"));
            Value = value with { Language = UiText.Resolve(value.Language).Tag };
            Error = null;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or JsonException)
        { Value = new(); Error = ex.Message; }
    }
    public void Save(DesktopPreferences value)
    {
        if (!value.IsValid) throw new ArgumentException(UiText.Get("无效的界面设置。"));
        var temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            File.WriteAllText(temp, JsonSerializer.Serialize(value));
            File.Move(temp, path, true); Value = value; Error = null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Error = ex.Message; }
        finally { try { if (File.Exists(temp)) File.Delete(temp); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }
}
