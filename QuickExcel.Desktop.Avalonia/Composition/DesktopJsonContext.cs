using System.Text.Json;
using System.Text.Json.Serialization;
using QuickExcel.Desktop.Avalonia.Web;
using QuickExcel.Presentation.Web;

namespace QuickExcel.Desktop.Avalonia;

internal sealed record SelfCheckResult(bool passed, List<string> checks, string? error, string? webViewRuntime,
    bool softwareRenderingRequested, string operatingSystem, string framework);

[JsonSourceGenerationOptions(GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(SelfCheckResult))]
[JsonSerializable(typeof(ScriptEditorPage.Changed))]
[JsonSerializable(typeof(ScriptEditorPage.Flushed))]
[JsonSerializable(typeof(ScriptEditorPage.Navigation))]
internal partial class DesktopJsonContext : JsonSerializerContext
{
    internal static DesktopJsonContext Web { get; } = new(new JsonSerializerOptions(WebMessage.Json));
    internal static DesktopJsonContext Indented { get; } = new(new JsonSerializerOptions { WriteIndented = true });
}
