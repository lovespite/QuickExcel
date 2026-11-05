using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Runtime;
using QuickExcel.ScriptContracts;

namespace QuickExcel.Core.Test;

public sealed class ScriptExecutionPolicyTests
{
    [Fact]
    public void RevocationCancelsPriorGenerationsEvenAfterReenable()
    {
        var policy = ScriptExecutionPolicy.Read([], _ => null);
        var prior = policy.CaptureLifetime();
        policy.UserSetting = true;
        Assert.True(prior.IsCancellationRequested);
        Assert.Throws<InvalidOperationException>(() => policy.CaptureLifetime());
        policy.UserSetting = false;
        var current = policy.CaptureLifetime();
        Assert.True(prior.IsCancellationRequested);
        Assert.False(current.IsCancellationRequested);
        Assert.NotEqual(prior, current);
    }

    [Fact]
    public void DisablingKillsAnIdlePersistentWorkerBetweenRounds()
    {
        if (!OperatingSystem.IsWindows()) return;
        var policy = ScriptExecutionPolicy.Current;
        var original = policy.UserSetting;
        using var session = new CancellationTokenSource();
        try
        {
            policy.RequireEnabled();
            var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule>
                { ["main"] = new("scripts/main.js", "export function VALUE(){return 42;}") }, ["main"], scriptApiVersion: 4,
                bindings: [new("main", "formula-functions")]);
            var executable = Path.Combine(AppContext.BaseDirectory, "ScriptWorker", "QuickExcel.ScriptWorker.exe");
            var request = new ScriptRuntimeRequest(modules, new(Guid.NewGuid(), 0, 0, 0), TimeSpan.FromSeconds(5), default) { RuntimeLifetime = session.Token };
            var lease = SharedScriptWorkers.Acquire(executable, request, actions: false);
            var runtime = lease.Runtime;
            using var process = System.Diagnostics.Process.GetProcessById(runtime.ProcessId);
            lease.Dispose(); // Retain the idle persistent context, with no active per-round cancellation callback.
            Assert.False(process.HasExited);
            policy.UserSetting = true;
            Assert.True(process.WaitForExit(5000));
            Assert.True(runtime.IsStopped);
        }
        finally { session.Cancel(); policy.UserSetting = original; }
    }

    [Fact]
    public void EverySourceIsIndependentAndAllReasonsRemainVisible()
    {
        for (var mask = 0; mask < 32; mask++)
        {
            var policy = ScriptExecutionPolicy.Read((mask & 2) != 0 ? ["book.xlsq", "--no-scripts"] : ["book.xlsq"],
                scope => (mask & (scope switch { EnvironmentVariableTarget.Process => 4, EnvironmentVariableTarget.User => 8, _ => 16 })) != 0 ? "1" : "0");
            policy.UserSetting = (mask & 1) != 0;
            var expected = new List<string>();
            if ((mask & 1) != 0) expected.Add("User setting");
            if ((mask & 2) != 0) expected.Add("Command line");
            if ((mask & 28) != 0) expected.Add("Env");
            Assert.Equal(mask != 0, policy.IsDisabled);
            Assert.Equal(mask == 0 ? "" : "No scripts - " + string.Join("/", expected), policy.Status);
            if (mask != 0) Assert.Contains(policy.Status, Assert.Throws<InvalidOperationException>(policy.RequireEnabled).Message);
            else policy.RequireEnabled();
            policy.UserSetting = false;
            Assert.Equal((mask & 30) != 0, policy.IsDisabled);
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("0")]
    [InlineData("true")]
    [InlineData(" 1 ")]
    public void OnlyTheDocumentedValueAndExactSwitchDisableScripts(string? value)
    {
        var policy = ScriptExecutionPolicy.Read(["--no-scripts=false", "book--no-scripts.xlsq"], _ => value);
        Assert.False(policy.IsDisabled);
    }

    [Fact]
    public async Task DisabledPolicyRejectsRuntimeAndCompilerBeforeAccessingWorkerPayload()
    {
        var policy = ScriptExecutionPolicy.Current;
        var original = policy.UserSetting;
        try
        {
            policy.UserSetting = true;
            // A missing executable would produce WorkerMissing if either guard were bypassed.
            var missing = Path.Combine(AppContext.BaseDirectory, "missing-worker.exe");
            var runtime = new ScriptWorkerRuntimeFactory(missing);
            Assert.Contains("No scripts", Assert.Throws<InvalidOperationException>(() => runtime.Create("throw Error('executed');")).Message);
            var request = new CompileRequest([new("main", "typescript", "const value: number = 1;")], ["main"], "") { ApiVersion = 4 };
            var compiler = new ScriptCompiler(missing);
            Assert.Contains("No scripts", (await Assert.ThrowsAsync<InvalidOperationException>(() => compiler.CompileAsync(
                request with { Sha256 = request.ComputeDigest() }, new(Guid.NewGuid(), 0, 0, 0)))).Message);
            Assert.Equal(0, compiler.ProcessLaunchCount);
            await Assert.ThrowsAsync<InvalidOperationException>(() => ScriptCompilation.PrepareAsync(ScriptModuleSet.Single(""), new(Guid.NewGuid(), 0, 0, 0), default));
        }
        finally { policy.UserSetting = original; }
    }

    [Fact]
    public void PreferenceRoundTripsAndOldSettingsDefaultToEnabled()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "test-data", "no-scripts-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "settings.json");
            File.WriteAllText(path, "{\"Zoom\":125}");
            var settings = new DesktopSettings(path); settings.Load();
            Assert.False(settings.Value.DisableScripts);
            settings.Save(settings.Value with { DisableScripts = true });
            Assert.Null(settings.Error);
            var reloaded = new DesktopSettings(path); reloaded.Load();
            Assert.True(reloaded.Value.DisableScripts);
            Assert.Equal(125, reloaded.Value.Zoom);
        }
        finally { Directory.Delete(directory, true); }
    }
}
