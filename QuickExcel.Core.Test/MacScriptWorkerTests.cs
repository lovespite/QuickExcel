using System.Diagnostics;
using System.Text;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Security;

namespace QuickExcel.Core.Test;

public sealed class MacScriptWorkerTests
{
    private static string? WorkerExecutable
    {
        get
        {
            if (!OperatingSystem.IsMacOS()) return null;
            var payload = Path.Combine(AppContext.BaseDirectory, "ScriptWorker");
            var executable = Path.Combine(payload, "QuickExcel.ScriptWorker");
            return File.Exists(executable) && File.Exists(Path.Combine(payload, "libcoreclr.dylib")) &&
                File.Exists(Path.Combine(payload, "libJavaScriptCore.dylib")) ? executable : null;
        }
    }

    private static bool Available => WorkerExecutable is not null;

    private static ScriptRuntimeRequest FormulaRequest(int apiVersion, string source, CancellationToken lifetime) => new(
        new ScriptModuleSet(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", source) },
            ["main"], scriptApiVersion: apiVersion, bindings: [new("main", "formula-functions")]),
        new(Guid.NewGuid(), 0, 0, 0), TimeSpan.FromSeconds(5), default) { RuntimeLifetime = lifetime };

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OwnershipChangesWaitForConcurrentCleanupAndRespectCancellation(bool cancel)
    {
        if (!OperatingSystem.IsMacOS()) return;
        MacWorkerLease.Recover();
        using var cancellation = new CancellationTokenSource();
        using var gate = new FileStream(Path.Combine(MacWorkerLease.Root, "root.lock"), FileMode.Open,
            FileAccess.ReadWrite, FileShare.None);
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var recovery = Task.Run(() =>
        {
            entered.SetResult();
            return OperatingSystem.IsMacOS() ? MacWorkerLease.Recover(cancellation.Token) : throw new PlatformNotSupportedException();
        });
        await entered.Task;
        await Task.Delay(100);
        Assert.False(recovery.IsCompleted);
        if (cancel)
        {
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => recovery.WaitAsync(TimeSpan.FromSeconds(2)));
        }
        else
        {
            gate.Dispose();
            Assert.True(await recovery.WaitAsync(TimeSpan.FromSeconds(5)) >= 0);
        }
    }

    [Theory]
    [InlineData(3)]
    [InlineData(4)]
    public async Task FormulaFactoryReusesContextAcrossIdleRoundsAndReclaimsItOnSessionClose(int apiVersion)
    {
        if (!OperatingSystem.IsMacOS()) return;
        Assert.NotNull(WorkerExecutable);
        using var session = new CancellationTokenSource();
        var request = FormulaRequest(apiVersion, apiVersion == 3
            ? "let n=0; host.functions.register('NEXT',()=>++n);"
            : "let n=0; export function NEXT(){return ++n;}", session.Token);
        var factory = new ScriptWorkerRuntimeFactory();
        var name = apiVersion == 3 ? "NEXT" : "JS.main.NEXT";
        int pid; string payload;
        using (var lease = Assert.IsType<ScriptWorkerLease>(factory.Create(request)))
        {
            pid = lease.Runtime.ProcessId; payload = lease.Runtime.SandboxPayloadDirectory;
            Assert.Equal(1d, lease.Invoke(name, []));
        }
        using var process = Process.GetProcessById(pid);
        await Task.Delay(5200);
        MacWorkerLease.Recover();
        Assert.False(process.HasExited);
        Assert.True(Directory.Exists(payload));
        using (var lease = Assert.IsType<ScriptWorkerLease>(factory.Create(request with { Stamp = request.Stamp with { CalculationGeneration = 1 } })))
        {
            Assert.Equal(pid, lease.Runtime.ProcessId);
            Assert.Equal(2d, lease.Invoke(name, []));
        }
        session.Cancel();
        Assert.True(process.WaitForExit(2000));
        Assert.True(SpinWait.SpinUntil(() => !Directory.Exists(payload), TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public void PersistentWorkerHasNoCumulativeCpuExpiryButStillStopsTimedOutOperations()
    {
        if (!OperatingSystem.IsMacOS()) return;
        Assert.NotNull(WorkerExecutable);
        using var session = new CancellationTokenSource();
        var request = FormulaRequest(4, """
            let n=0;
            export function SPIN(){const end=Date.now()+150; while(Date.now()<end){} return ++n;}
            export function HANG(){while(true){}}
            """, session.Token);
        var factory = new ScriptWorkerRuntimeFactory();
        int pid;
        using (var lease = Assert.IsType<ScriptWorkerLease>(factory.Create(request))) pid = lease.Runtime.ProcessId;
        using var process = Process.GetProcessById(pid);
        var timer = Stopwatch.StartNew();
        var calls = 0;
        // RLIMIT_CPU formerly killed execution at 30/31 accumulated CPU seconds,
        // even when every individual operation and round finished on time.
        while (process.TotalProcessorTime < TimeSpan.FromSeconds(32))
        {
            Assert.True(timer.Elapsed < TimeSpan.FromSeconds(120), "CPU-budget regression did not make progress.");
            using var lease = Assert.IsType<ScriptWorkerLease>(factory.Create(request));
            Assert.Equal(pid, lease.Runtime.ProcessId);
            Assert.Equal((double)++calls, lease.Invoke("JS.main.SPIN", []));
            process.Refresh();
        }
        using (var lease = Assert.IsType<ScriptWorkerLease>(factory.Create(request)))
            Assert.Throws<TimeoutException>(() => lease.Invoke("JS.main.HANG", []));
        Assert.True(process.WaitForExit(2000));
        using var replacement = Assert.IsType<ScriptWorkerLease>(factory.Create(request));
        Assert.NotEqual(pid, replacement.Runtime.ProcessId);
        Assert.Equal(1d, replacement.Invoke("JS.main.SPIN", []));
        session.Cancel();
    }

    [Fact]
    public void RealWorkerExecutesJavaScriptUnderSeatbelt()
    {
        if (!Available) return;
        using var runtime = new ScriptWorkerRuntimeFactory().Create("host.functions.register('ADD', value => value + 1);");
        Assert.Equal(42d, Convert.ToDouble(runtime.Invoke("ADD", [41d])));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FootprintWatchdogStopsRunawayEngineAllocation(bool persistent)
    {
        if (!Available) return;
        const string source = "host.functions.register('BIG', () => { const chunks=[]; for(let i=0;i<24;i++){const a=new Uint8Array(32*1024*1024); a.fill(1); chunks.push(a);} return chunks.length; });";
        using var session = new CancellationTokenSource();
        try
        {
            using var runtime = new ScriptWorkerRuntimeFactory().Create(FormulaRequest(3, source, persistent ? session.Token : default));
            Assert.ThrowsAny<Exception>(() => runtime.Invoke("BIG", []));
        }
        finally { session.Cancel(); }
    }

    [Fact]
    public void AbandonedSandboxRecoveryIsAvailable()
    {
        if (!OperatingSystem.IsMacOS()) return;
        Assert.True(ScriptWorkerRuntimeFactory.RecoverAbandonedSandboxes() >= 0);
    }

    [Fact]
    public void KeychainEitherProtectsOrFailsClosedWithoutApplicationIdentifier()
    {
        if (!OperatingSystem.IsMacOS()) return;
        var protection = TrustDataProtection.Create();
        var purpose = Encoding.UTF8.GetBytes("QuickExcel.Core.Test/trust");
        var data = Encoding.UTF8.GetBytes("trust-test");
        try
        {
            var encrypted = protection.Protect(data, purpose);
            Assert.NotEqual(data, encrypted);
            Assert.Equal(data, protection.Unprotect(encrypted, purpose));
        }
        catch (UserProtectionUnavailableException)
        {
            // An unsigned test host has no application identifier; the only
            // permitted behavior is to disable persistent trust, never plaintext.
        }
    }

    [Fact]
    public void OnlyDeadWorkerLeasesAreReclaimed()
    {
        if (!OperatingSystem.IsMacOS() || WorkerExecutable is not { } executable) return;
        using var process = MacWorkerProcess.Start(executable, CancellationToken.None);
        var directory = process.PayloadDirectory;
        MacWorkerLease.Recover();
        Assert.True(Directory.Exists(directory), "A live lease and its payload must be retained");
        process.Kill();
        Thread.Sleep(100);
        process.Dispose();
        Assert.False(Directory.Exists(directory), "A dead worker's leased payload must be removed");
    }
}
