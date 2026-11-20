using System.Buffers.Binary;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;

namespace QuickExcel.ScriptContracts;

// This assembly deliberately has no reference to Core, WPF, ExcelIO or native code.
public sealed record RequestIdentity(Guid SessionId, Guid WorkerInstanceId, Guid RequestId,
    long Revision, long? CalculationGeneration, long ScriptGeneration)
{
    [JsonRequired] public Guid RoundId { get; init; } = Guid.NewGuid();
    public void Validate()
    {
        if (SessionId == Guid.Empty || WorkerInstanceId == Guid.Empty || RequestId == Guid.Empty || RoundId == Guid.Empty ||
            Revision < 0 || CalculationGeneration is < 0 || ScriptGeneration < 0)
            throw new ProtocolException("Invalid request identity.");
    }
    public bool SameRound(RequestIdentity other) => this with { RequestId = other.RequestId } == other;
}

public sealed record WorkerEnvelope(int ProtocolVersion, RequestIdentity Identity, string Kind, JsonElement Payload);
public sealed record WorkerHello(string Engine, string Policy, int[] ApiVersions, int MaxFrameBytes, string Role = "execute", string? Compiler = null);
public sealed record LoadScript(string Source, string Sha256);
public sealed record FunctionInfo(string Name, int ApiVersion, int MinArguments, int MaxArguments, string ReturnType, string Profile);
public sealed record LoadedScript(string Sha256, FunctionInfo[] Functions);
public sealed record LoadedWorkbook(string Sha256, FunctionInfo[] Functions, ActionInfo[] Actions);
public sealed record BeginWorkerRound;
public sealed record FunctionCall(string Name, WireValue[] Arguments);
public sealed record InvokeBatch(FunctionCall[] Calls);
public sealed record InvocationResult(WireValue? Value, string? Error, bool ValueTooLong = false);
public sealed record BatchResult(InvocationResult[] Results);
public sealed record WorkerFailure(string Code, string Message)
{
    public string[] Logs { get; init; } = [];
}
public sealed record WorkerLogs(string[] Messages);
public sealed record EmptyPayload;
public sealed class ProtocolException(string message, Exception? inner = null) : IOException(message, inner);

public static class WorkerProtocol
{
    public const int Version = 11;
    public const string Engine = "JavaScriptCore/2.54.2/b95c91b1/qx-sync-2";
    public const string Policy = "conservative-jsc-v1-memory-issue-open";
    public const int MaxFrameBytes = 16 * 1024 * 1024;
    public const int MaxBatchCalls = 256;
    public const int MaxCellTextLength = 32767;
    // A UTF-16 code unit can occupy six JSON bytes. Bound the worst-case reply,
    // including per-result metadata and the envelope, before executing user code.
    public const int MaxCalculationBatchCalls = (MaxFrameBytes - 1024) / (MaxCellTextLength * 6 + 256);
    public const int MaxFunctions = 4096;
    public const int MaxLogs = 100;
    public const int MaxLogBytes = 4000;
    public static readonly JsonSerializerOptions Json = WorkerJsonContext.Default.Options;

    public static JsonTypeInfo<T> TypeInfo<T>() =>
        WorkerJsonContext.Default.GetTypeInfo(typeof(T)) as JsonTypeInfo<T>
        ?? throw new NotSupportedException($"Unregistered worker JSON type: {typeof(T)}.");

    public static WorkerEnvelope Message<T>(RequestIdentity identity, string kind, T payload) =>
        new(Version, identity, kind, EncodePayload(payload));

    public static JsonElement EncodePayload<T>(T value)
    {
        using var buffer = new LimitedBuffer();
        JsonSerializer.Serialize(buffer, value, TypeInfo<T>());
        buffer.Position = 0;
        using var document = JsonDocument.Parse(buffer, new() { MaxDepth = 24 });
        return document.RootElement.Clone();
    }

    public static T Payload<T>(WorkerEnvelope message)
    {
        try { return message.Payload.Deserialize(TypeInfo<T>()) ?? throw new ProtocolException("Missing payload."); }
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
        await JsonSerializer.SerializeAsync(payload, message, TypeInfo<WorkerEnvelope>(), cancellationToken).ConfigureAwait(false);
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
            var message = document.RootElement.Deserialize(TypeInfo<WorkerEnvelope>()) ?? throw new ProtocolException("Missing message.");
            Validate(message);
            return message;
        }
        catch (JsonException e) { throw new ProtocolException("Malformed IPC JSON.", e); }
    }

    private static void Validate(WorkerEnvelope message)
    {
        if (message.ProtocolVersion != Version || message.Identity is null ||
            message.Kind is not ("hello" or "load" or "loadModules" or "loadActions" or "runAction" or "invokeBatch" or "logs" or "shutdown" or "failure"
                or "compile" or "model" or "loadWorkbook" or "beginRound" or "endRound"
                ) ||
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
