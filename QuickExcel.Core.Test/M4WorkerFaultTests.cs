using System.Diagnostics;
using QuickExcel.Runtime;

namespace QuickExcel.Core.Test;

public sealed class M4WorkerFaultTests
{
    private static ScriptWorkerRuntime Create(string mode) => (ScriptWorkerRuntime)new ScriptWorkerRuntimeFactory(
        Path.Combine(AppContext.BaseDirectory, "Fixtures", "WorkerFixture", "WorkerFixture.exe")).Create(new ScriptRuntimeRequest(
            mode, new(Guid.NewGuid(), 1, 2, 3), TimeSpan.FromSeconds(2), default));

    [Theory]
    [InlineData("session")] [InlineData("worker")] [InlineData("request")]
    [InlineData("revision")] [InlineData("calculation")] [InlineData("script")]
    [InlineData("duplicate")] [InlineData("oversize")] [InlineData("half")] [InlineData("stderr")]
    public void HostRejectsMisbehavingPeerBeforeAcceptingRegistry(string mode) =>
        Assert.Throws<ScriptRuntimeFailureException>(() => Create(mode));

    [Theory]
    [InlineData("bad-count")] [InlineData("bad-value")] [InlineData("bad-logs")]
    public void InvalidResultsAndLogsStopPeer(string mode)
    {
        using var runtime = Create(mode);
        Assert.Equal("Protocol", Assert.Throws<ScriptRuntimeFailureException>(() =>
        {
            if (mode == "bad-logs") runtime.GetLogs(); else runtime.Invoke("F", []);
        }).Code);
        Assert.Throws<ScriptRuntimeFailureException>(() => runtime.Invoke("F", []));
    }

    [Theory]
    [InlineData("hang")] [InlineData("hang-load")]
    public void WatchdogTerminatesPeerThatIgnoresAllRequests(string mode)
    {
        var clock = Stopwatch.StartNew();
        Assert.Throws<TimeoutException>(() => { using var runtime = Create(mode); runtime.Invoke("F", []); });
        Assert.InRange(clock.Elapsed.TotalSeconds, 1, 4);
    }

    [Fact]
    public void JobCommitBudgetDeniesAllocationOutsideJsEngine()
    {
        using var runtime = Create("memory");
        Assert.Equal(true, runtime.Invoke("F", []));
    }
}
