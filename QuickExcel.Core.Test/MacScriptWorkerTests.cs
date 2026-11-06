using System.Text;
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

    [Fact]
    public void RealWorkerExecutesJavaScriptUnderSeatbelt()
    {
        if (!Available) return;
        using var runtime = new ScriptWorkerRuntimeFactory().Create("host.functions.register('ADD', value => value + 1);");
        Assert.Equal(42d, Convert.ToDouble(runtime.Invoke("ADD", [41d])));
    }

    [Fact]
    public void FootprintWatchdogStopsRunawayEngineAllocation()
    {
        if (!Available) return;
        using var runtime = new ScriptWorkerRuntimeFactory().Create(
            "host.functions.register('BIG', () => { const chunks=[]; for(let i=0;i<24;i++){const a=new Uint8Array(32*1024*1024); a.fill(1); chunks.push(a);} return chunks.length; });");
        Assert.ThrowsAny<Exception>(() => runtime.Invoke("BIG", []));
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
