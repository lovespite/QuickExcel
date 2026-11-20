using System.Text.Json;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Core.Test;

public sealed class ScriptNetworkProtocolTests
{
    [Fact]
    public async Task DirectModelFramePreservesWireBytesAndValidation()
    {
        var identity = M4ProtocolTests.Identity();
        using var request = JsonDocument.Parse("{\"method\":\"network.fetch\",\"args\":[{\"url\":\"https://example.com/中文\",\"bodyBase64\":\"-_8=\"}]}");
        var call = new ModelCall(Guid.NewGuid(), 1, request.RootElement);
        using var expected = new MemoryStream(); using var actual = new MemoryStream();
        await WorkerProtocol.WriteAsync(expected, WorkerProtocol.Message(identity, "model", call));
        await WorkerProtocol.WriteModelCallAsync(actual, identity, call);
        Assert.Equal(expected.ToArray(), actual.ToArray());
        using var direct = new MemoryStream();
        await WorkerProtocol.WriteNetworkCallAsync(direct, identity, call.OperationId, call.Sequence, request.RootElement.GetProperty("args")[0]);
        Assert.Equal(expected.ToArray(), direct.ToArray());
        actual.Position = 0;
        var decoded = await WorkerProtocol.ReadAsync(actual); WorkerProtocol.ValidateReply(decoded, identity, "model");
        var model = WorkerProtocol.Payload<ModelCall>(decoded);
        Assert.Equal(call.OperationId, model.OperationId); Assert.Equal(call.Sequence, model.Sequence);
        Assert.Equal(request.RootElement.GetProperty("args")[0].GetProperty("url").GetString(), model.Request.GetProperty("args")[0].GetProperty("url").GetString());
        using var invalid = new MemoryStream();
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.WriteModelCallAsync(invalid, identity, call with { OperationId = Guid.Empty }));
        Assert.Equal(0, invalid.Length);
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.WriteNetworkCallAsync(invalid, identity, call.OperationId, 0,
            request.RootElement.GetProperty("args")[0]));
        Assert.Equal(0, invalid.Length);
    }

    [Fact]
    public async Task DirectModelFrameRejectsOversizeBeforeWritingAnyBytes()
    {
        using var request = JsonDocument.Parse("{\"body\":\"" + new string('x', WorkerProtocol.MaxFrameBytes) + "\"}");
        using var stream = new MemoryStream();
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.WriteModelCallAsync(stream,
            M4ProtocolTests.Identity(), new(Guid.NewGuid(), 1, request.RootElement)));
        Assert.Equal(0, stream.Length);
        await Assert.ThrowsAsync<ProtocolException>(() => WorkerProtocol.WriteNetworkCallAsync(stream,
            M4ProtocolTests.Identity(), Guid.NewGuid(), 1, request.RootElement));
        Assert.Equal(0, stream.Length);
    }
}
