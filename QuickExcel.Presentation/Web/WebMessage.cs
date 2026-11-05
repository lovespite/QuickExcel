using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace QuickExcel.Presentation.Web;

/// <summary>UI-neutral, versioned local-page protocol. Independent of worker IPC.</summary>
public sealed record WebMessage(int Protocol, string Instance, string Document, long Generation, string Type, string RequestId, JsonElement Payload)
{
    public const int Version = 1, MaximumBytes = 256 * 1024;
    public static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web)
    { MaxDepth = 24, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow };
    public static WebMessage Parse(string body, string instance, string document, long generation, IReadOnlySet<string> allowed)
    {
        if (body.Length > MaximumBytes || Encoding.UTF8.GetByteCount(body) > MaximumBytes) throw new InvalidDataException(UiText.Get("Web 消息超过 256 KiB。"));
        using var parsed = JsonDocument.Parse(body, new() { MaxDepth = 24 });
        if (parsed.RootElement.ValueKind != JsonValueKind.Object || parsed.RootElement.EnumerateObject().Count() != 7)
            throw new InvalidDataException(UiText.Get("Web 消息缺少必需字段或含额外字段。"));
        RejectDuplicates(parsed.RootElement);
        var message = parsed.RootElement.Deserialize<WebMessage>(Json) ?? throw new InvalidDataException(UiText.Get("缺少 Web 消息。"));
        if (message.Protocol != Version || message.Instance != instance || message.Document != document || message.Generation != generation)
            throw new InvalidDataException(UiText.Get("Web 消息版本、文档或视图代次不匹配。"));
        if (message.Type is null || !allowed.Contains(message.Type) || !ValidId(message.RequestId)) throw new InvalidDataException(UiText.Get("Web 消息类型或请求标识无效。"));
        return message with { Payload = message.Payload.Clone() };
    }
    public static bool ValidId(string? id) => id is { Length: > 0 and <= 128 } && id.All(c => char.IsAsciiLetterOrDigit(c) || c is '-' or '_' or '.');
    private static void RejectDuplicates(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in value.EnumerateObject())
            { if (!names.Add(property.Name)) throw new InvalidDataException(UiText.Get("Web 消息含重复字段。")); RejectDuplicates(property.Value); }
        }
        else if (value.ValueKind == JsonValueKind.Array) foreach (var item in value.EnumerateArray()) RejectDuplicates(item);
    }
    public static string Create(string instance, string document, long generation, string type, string requestId, object payload)
    {
        var body = JsonSerializer.Serialize(new { protocol = Version, instance, document, generation, type, requestId, payload }, Json);
        if (Encoding.UTF8.GetByteCount(body) > MaximumBytes) throw new InvalidDataException(UiText.Get("Web 消息必须分块发送。"));
        return body;
    }
}
