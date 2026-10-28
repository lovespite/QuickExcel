using System.Buffers.Binary;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace QuickExcel.ScriptContracts;

// This assembly deliberately has no reference to Core, WPF, ExcelIO or native code.
public sealed record RequestIdentity(Guid SessionId, Guid WorkerInstanceId, Guid RequestId,
    long Revision, long? CalculationGeneration, long ScriptGeneration)
{
    public void Validate()
    {
        if (SessionId == Guid.Empty || WorkerInstanceId == Guid.Empty || RequestId == Guid.Empty ||
            Revision < 0 || CalculationGeneration is null or < 0 || ScriptGeneration < 0)
            throw new ProtocolException("Invalid request identity.");
    }
    public bool SameRound(RequestIdentity other) => this with { RequestId = other.RequestId } == other;
}

public sealed record WorkerEnvelope(int ProtocolVersion, RequestIdentity Identity, string Kind, JsonElement Payload);
public sealed record WorkerHello(string Engine, string Policy, int[] ApiVersions, int MaxFrameBytes);
public sealed record LoadScript(string Source, string Sha256);
public sealed record FunctionInfo(string Name, int ApiVersion, int MinArguments, int MaxArguments, string ReturnType, string Profile);
public sealed record LoadedScript(string Sha256, FunctionInfo[] Functions);
public sealed record FunctionCall(string Name, WireValue[] Arguments);
public sealed record InvokeBatch(FunctionCall[] Calls);
public sealed record InvocationResult(WireValue? Value, string? Error);
public sealed record BatchResult(InvocationResult[] Results);
public sealed record WorkerFailure(string Code, string Message);
public sealed record WorkerLogs(string[] Messages);
public sealed record EmptyPayload;
public sealed class ProtocolException(string message, Exception? inner = null) : IOException(message, inner);

public static class WorkerProtocol
{
    public const int Version = 2;
    public const string Engine = "QuickJS-NG/0.17.0";
    public const string Policy = "conservative-worker-v1";
    public const int MaxFrameBytes = 4 * 1024 * 1024;
    public const int MaxBatchCalls = 256;
    public const int MaxFunctions = 4096;
    public const int MaxLogs = 100;
    public const int MaxLogBytes = 4000;
    public static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web)
    {
        MaxDepth = 24,
        PropertyNameCaseInsensitive = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        RespectNullableAnnotations = true,
        RespectRequiredConstructorParameters = true
    };

    public static WorkerEnvelope Message<T>(RequestIdentity identity, string kind, T payload) =>
        new(Version, identity, kind, EncodePayload(payload));

    public static JsonElement EncodePayload<T>(T value)
    {
        using var buffer = new LimitedBuffer();
        JsonSerializer.Serialize(buffer, value, Json);
        buffer.Position = 0;
        using var document = JsonDocument.Parse(buffer, new() { MaxDepth = 24 });
        return document.RootElement.Clone();
    }

    public static T Payload<T>(WorkerEnvelope message)
    {
        try { return message.Payload.Deserialize<T>(Json) ?? throw new ProtocolException("Missing payload."); }
        catch (JsonException e) { throw new ProtocolException("Invalid payload.", e); }
    }

    public static void ValidateReply(WorkerEnvelope reply, RequestIdentity expected, string kind)
    {
        Validate(reply);
        if (reply.Identity != expected || reply.Kind != kind && reply.Kind != "failure")
            throw new ProtocolException("Worker reply identity or kind does not match the request.");
    }

    public static async Task WriteAsync(Stream stream, WorkerEnvelope message, CancellationToken cancellationToken = default)
    {
        Validate(message);
        // Serialize through a capped stream: oversized output never allocates an unbounded byte array.
        using var payload = new LimitedBuffer();
        await JsonSerializer.SerializeAsync(payload, message, Json, cancellationToken).ConfigureAwait(false);
        var header = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(header, checked((int)payload.Length));
        await stream.WriteAsync(header, cancellationToken).ConfigureAwait(false);
        payload.Position = 0;
        await payload.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
        await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    public static async Task<WorkerEnvelope> ReadAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        var header = new byte[4];
        await stream.ReadExactlyAsync(header, cancellationToken).ConfigureAwait(false);
        var length = BinaryPrimitives.ReadInt32LittleEndian(header);
        if (length is <= 0 or > MaxFrameBytes) throw new ProtocolException("IPC frame exceeds its byte budget.");
        var data = new byte[length];
        await stream.ReadExactlyAsync(data, cancellationToken).ConfigureAwait(false);
        try
        {
            using var document = JsonDocument.Parse(data, new() { MaxDepth = 24 });
            RejectDuplicateKeys(document.RootElement);
            var message = document.RootElement.Deserialize<WorkerEnvelope>(Json) ?? throw new ProtocolException("Missing message.");
            Validate(message);
            return message;
        }
        catch (JsonException e) { throw new ProtocolException("Malformed IPC JSON.", e); }
    }

    private static void Validate(WorkerEnvelope message)
    {
        if (message.ProtocolVersion != Version || message.Identity is null ||
            message.Kind is not ("hello" or "load" or "loadModules" or "invokeBatch" or "logs" or "shutdown" or "failure") ||
            message.Payload.ValueKind != JsonValueKind.Object)
            throw new ProtocolException("Unsupported protocol message.");
        message.Identity.Validate();
    }

    private static void RejectDuplicateKeys(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Object)
        {
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in value.EnumerateObject())
            {
                if (!keys.Add(property.Name)) throw new ProtocolException("Duplicate JSON property.");
                RejectDuplicateKeys(property.Value);
            }
        }
        else if (value.ValueKind == JsonValueKind.Array)
            foreach (var item in value.EnumerateArray()) RejectDuplicateKeys(item);
    }

    private sealed class LimitedBuffer : MemoryStream
    {
        private void Check(int count)
        { if (Position + count > MaxFrameBytes) throw new ProtocolException("IPC frame exceeds its byte budget."); }
        public override void Write(byte[] buffer, int offset, int count) { Check(count); base.Write(buffer, offset, count); }
        public override void Write(ReadOnlySpan<byte> buffer) { Check(buffer.Length); base.Write(buffer); }
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default)
        { token.ThrowIfCancellationRequested(); Write(buffer.Span); return ValueTask.CompletedTask; }
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token)
        { token.ThrowIfCancellationRequested(); Write(buffer, offset, count); return Task.CompletedTask; }
    }
}
