using System.Text.Json;
using System.Text.Json.Serialization;
using QuickExcel.ScriptContracts;

[JsonSourceGenerationOptions(JsonSerializerDefaults.Web, MaxDepth = 24, PropertyNameCaseInsensitive = false,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow, RespectNullableAnnotations = true,
    RespectRequiredConstructorParameters = true)]
[JsonSerializable(typeof(CompilerBridgeRequest))]
[JsonSerializable(typeof(ModelBridgeReply))]
internal sealed partial class WorkerBridgeJsonContext : JsonSerializerContext;

internal sealed record CompilerBridgeRequest(CompilerModule[] Modules, string[] Entries, int ApiVersion, string SourceMapMode);
internal sealed record ModelBridgeReply(bool Ok, JsonElement Value, string? Error, string? Code);
