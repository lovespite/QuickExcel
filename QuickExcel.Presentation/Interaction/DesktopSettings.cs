using QuickExcel.Serialization;
using QuickExcel.Core;
using System.IO;
using System.Text.Json;

namespace QuickExcel.Presentation;

public enum SheetNavigationPlacement { Bottom, Left, Right }
public enum DesktopCalculationMode { Automatic, Partial, Manual }
public enum DesktopSourceMapMode { None, Compact, Full }
public sealed record DesktopPreferences(SheetNavigationPlacement Sheets = SheetNavigationPlacement.Bottom, int Zoom = 100,
    DesktopCalculationMode Calculation = DesktopCalculationMode.Automatic, string Language = UiText.Chinese,
    DesktopSourceMapMode ScriptSourceMap = DesktopSourceMapMode.Compact, bool DisableScripts = false)
{
    public ScriptNetworkOptions ScriptNetwork { get; init; } = new();
    internal bool IsValid => Enum.IsDefined(Sheets) && Zoom is >= 75 and <= 175 && Enum.IsDefined(Calculation) && Enum.IsDefined(ScriptSourceMap)
        && ScriptNetwork is not null && ScriptNetwork.IsValid;
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
            using var json = JsonDocument.Parse(File.ReadAllText(path));
            var value = ReadPreferences(json.RootElement);
            if (value is null || !value.IsValid) throw new InvalidDataException(UiText.Get("无效的界面设置。"));
            Value = value with { Language = UiText.Resolve(value.Language).Tag };
            Error = null;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or JsonException)
        { Value = new(); Error = ex.Message; }
    }
    internal static DesktopPreferences? ReadPreferences(JsonElement json)
    {
        var value = json.Deserialize(PresentationJsonContext.Default.DesktopPreferences);
        // Generated parameterized-record readers may supply null for an absent init-only
        // property. Apply the new default only when older files omit it; explicit null or
        // malformed network settings must still fail validation.
        return value is not null && json.ValueKind == JsonValueKind.Object && !json.TryGetProperty(nameof(DesktopPreferences.ScriptNetwork), out _)
            ? value with { ScriptNetwork = new() } : value;
    }
    public void Save(DesktopPreferences value)
    {
        if (!value.IsValid) throw new ArgumentException(UiText.Get("无效的界面设置。"));
        var temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            File.WriteAllText(temp, JsonSerializer.Serialize(value, PresentationJsonContext.Default.DesktopPreferences));
            File.Move(temp, path, true); Value = value; Error = null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Error = ex.Message; }
        finally { try { if (File.Exists(temp)) File.Delete(temp); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }
}
