using System.Text.Json;
using System.Text.Json.Serialization;

namespace QuickExcel.Presentation.Web;

[JsonSourceGenerationOptions(JsonSerializerDefaults.Web, GenerationMode = JsonSourceGenerationMode.Metadata,
    MaxDepth = 24, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(WebMessage))]
[JsonSerializable(typeof(DraftChunk))]
public partial class WebJsonContext : JsonSerializerContext;
