using System.Text.Json.Serialization;
using QuickExcel.Presentation;

namespace QuickExcel.Serialization;

[JsonSourceGenerationOptions(GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(DesktopPreferences))]
[JsonSerializable(typeof(string[]))]
[JsonSerializable(typeof(Dictionary<string, string>))]
[JsonSerializable(typeof(RecentFileStore.Settings), TypeInfoPropertyName = "RecentSettings")]
[JsonSerializable(typeof(QuickCellStyleStore.Library), TypeInfoPropertyName = "StyleLibrary")]
[JsonSerializable(typeof(DesktopSettingsMigration.MigrationMarker))]
internal partial class PresentationJsonContext : JsonSerializerContext;
