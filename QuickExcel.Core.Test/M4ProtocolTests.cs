using System.Buffers.Binary;
using System.Text;
using System.Text.Json;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Core.Test;

public sealed class M4ProtocolTests
{
    internal static RequestIdentity Identity() => new(Guid.NewGuid(), Guid.NewGuid(), Guid.NewGuid(), 1, 2, 3);

    [Fact]
    public async Task FragmentedFramesRoundtripUnicodeAndAllIdentityFields()
    {
        var message = WorkerProtocol.Message(Identity(), "load", new LoadScript("中文 🧮", "digest"));
        using var encoded = new MemoryStream();
        await WorkerProtocol.WriteAsync(encoded, message);
        using var fragmented = new FragmentedStream(encoded.ToArray());
        var decoded = await WorkerProtocol.ReadAsync(fragmented);
        WorkerProtocol.ValidateReply(decoded, message.Identity, message.Kind);
        Assert.Equal("中文 🧮", WorkerProtocol.Payload<LoadScript>(decoded).Source);
    }

    [Theory]
    [InlineData(0)] [InlineData(-1)] [InlineData(4 * 1024 * 1024 + 1)] [InlineData(int.MaxValue)]
    public async Task LengthIsRejectedBeforeReadingOrAllocatingBody(int length)
    {
        var bytes = new byte[4]; BinaryPrimitives.WriteInt32LittleEndian(bytes, length);
        using var stream = new MemoryStream(bytes);
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.ReadAsync(stream));
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(3)] [InlineData(5)]
    public async Task TruncatedFramesHaveControlledFailure(int count)
    {
        byte[] bytes = [8, 0, 0, 0, 123];
        using var stream = new MemoryStream(bytes.Take(count).ToArray());
        await Assert.ThrowsAsync<EndOfStreamException>(() => WorkerProtocol.ReadAsync(stream));
    }

    [Theory]
    [InlineData("session")] [InlineData("worker")] [InlineData("request")]
    [InlineData("revision")] [InlineData("calculation")] [InlineData("script")]
    [InlineData("kind")] [InlineData("version")]
    public void MismatchedReplyCannotBeAccepted(string field)
    {
        var expected = Identity();
        var identity = field switch
        {
            "session" => expected with { SessionId = Guid.NewGuid() },
            "worker" => expected with { WorkerInstanceId = Guid.NewGuid() },
            "request" => expected with { RequestId = Guid.NewGuid() },
            "revision" => expected with { Revision = 9 },
            "calculation" => expected with { CalculationGeneration = null },
            "script" => expected with { ScriptGeneration = 9 },
            _ => expected
        };
        var reply = WorkerProtocol.Message(identity, field == "kind" ? "load" : "logs", new EmptyPayload());
        if (field == "version") reply = reply with { ProtocolVersion = 99 };
        Assert.Throws<ProtocolException>(() => WorkerProtocol.ValidateReply(reply, expected, "logs"));
    }

    [Theory]
    [InlineData("{\"protocolVersion\":1,\"protocolVersion\":1}")]
    [InlineData("{\"protocolVersion\":1}")]
    [InlineData("null")]
    [InlineData("{broken")]
    public async Task MissingDuplicateAndMalformedFieldsFail(string json)
    {
        using var stream = Frame(Encoding.UTF8.GetBytes(json));
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.ReadAsync(stream));
    }

    [Fact]
    public async Task AlternatePropertyCasingCannotOverrideCanonicalIdentity()
    {
        var message = WorkerProtocol.Message(Identity(), "logs", new EmptyPayload());
        var json = JsonSerializer.Serialize(message, WorkerProtocol.Json).Replace($"\"protocolVersion\":{WorkerProtocol.Version}", $"\"ProtocolVersion\":99,\"protocolVersion\":{WorkerProtocol.Version}");
        using var stream = Frame(Encoding.UTF8.GetBytes(json));
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.ReadAsync(stream));
    }

    [Fact]
    public async Task Utf8ByteBudgetAndDepthAreEnforced()
    {
        using var stream = new MemoryStream();
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.WriteAsync(stream,
            WorkerProtocol.Message(Identity(), "load", new LoadScript(new string('中', 1_500_000), "hash"))));
        Assert.Equal(0, stream.Length); // A rejected outgoing frame writes no partial header or body.
        using var deep = Frame(Encoding.UTF8.GetBytes(new string('[', 30) + "0" + new string(']', 30)));
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.ReadAsync(deep));
        using var invalidUtf8 = Frame([0xFF, 0xFE]);
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.ReadAsync(invalidUtf8));
    }

    [Theory]
    [InlineData("number", "\"00123\"")]
    [InlineData("number", "1e999")]
    [InlineData("text", "123")]
    [InlineData("error", "\"#VALUE!\"")]
    [InlineData("blank", "\"\"")]
    [InlineData("array", "[[1]]")]
    [InlineData("range", "{\"kind\":\"range\",\"rows\":2,\"columns\":2,\"values\":[1]}")]
    [InlineData("range", "{\"kind\":\"range\",\"rows\":2147483647,\"columns\":2147483647,\"values\":[]}")]
    public void TypedValuesRejectAmbiguousOrMalformedPayloads(string kind, string json)
    {
        using var document = JsonDocument.Parse(json);
        Assert.Throws<ProtocolException>(() => new WireValue(kind, document.RootElement).Validate());
    }

    [Fact]
    public void TypedScalarsPreserveBlankEmptyTextBooleanAndErrorLookingText()
    {
        foreach (var value in new object?[] { null, "", "00123", "#VALUE!", 123d, true, false })
            Assert.Equal(value, WireValue.FromObject(value).ToObject(scalarOnly: true));
    }

    private static MemoryStream Frame(byte[] data)
    {
        var stream = new MemoryStream();
        var header = new byte[4]; BinaryPrimitives.WriteInt32LittleEndian(header, data.Length);
        stream.Write(header); stream.Write(data); stream.Position = 0;
        return stream;
    }
    private sealed class FragmentedStream(byte[] data) : MemoryStream(data)
    {
        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default) =>
            base.ReadAsync(buffer[..Math.Min(1, buffer.Length)], token);
    }
}
