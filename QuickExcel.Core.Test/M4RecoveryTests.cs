using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;
using QuickExcel.Runtime;

namespace QuickExcel.Core.Test;

#pragma warning disable CA1416
public sealed class M4RecoveryTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "QuickExcel.RecoveryTest." + Guid.NewGuid().ToString("N"));
    private static string Worker => Path.Combine(AppContext.BaseDirectory, "ScriptWorker", "QuickExcel.ScriptWorker.exe");

    [Theory]
    [InlineData("record")]
    [InlineData("profile")]
    [InlineData("partial")]
    [InlineData("staged")]
    [InlineData("running")]
    [InlineData("busy")]
    public async Task HostDeathAtEachCreationStageIsRecoveredAndRetryIsIdempotent(string stage)
    {
        var (host, record) = await StartHost(stage);
        using (host)
        using (var worker = record.ProcessId == 0 ? null : Process.GetProcessById(record.ProcessId))
        {
            try
            {
                Assert.Equal(0, WorkerSandboxLease.Recover(root: _root)); // Must not steal another host's live lease.
                Assert.True(Directory.Exists(Path.GetDirectoryName(record.Payload)));
                if (stage != "record") Assert.True(Directory.Exists(ProfilePath(record.Profile)));
                host.Kill(); // Deliberately do not kill the tree: containment must terminate the worker.
                await host.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
                await WaitForLeaseRelease(record);
                if (worker is not null) await worker.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
                Assert.Equal(1, WorkerSandboxLease.Recover(root: _root));
                Assert.False(Directory.Exists(record.Payload));
                Assert.False(Directory.Exists(ProfilePath(record.Profile)));
                Assert.Empty(Directory.EnumerateDirectories(_root));
                Assert.Equal(0, WorkerSandboxLease.Recover(root: _root));
                using var next = WorkerSandbox.Create(Worker, default, _root);
                Assert.NotEqual(record.Profile, next.ProfileName);
            }
            finally { if (!host.HasExited) { host.Kill(); await host.WaitForExitAsync(); } }
        }
    }

    [Fact]
    public async Task RecoveryCannotRemoveResourcesWhileTheJobStillHasAWorker()
    {
        var (host, record) = await StartHost("busy");
        using (host)
        using (var worker = Process.GetProcessById(record.ProcessId))
        using (var retainedJob = OpenJobObject(4, false, @"Global\" + record.Profile))
        {
            try
            {
                Assert.False(retainedJob.IsInvalid);
                host.Kill(); await host.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
                await WaitForLeaseRelease(record);
                Assert.False(worker.HasExited);
                var files = Directory.GetFiles(record.Payload);
                var error = Assert.Throws<IOException>(() => WorkerSandboxLease.Recover(root: _root));
                Assert.Contains("Job still has active processes", error.ToString());
                Assert.Equal(files, Directory.GetFiles(record.Payload));
                Assert.True(Directory.Exists(record.Payload));
                Assert.True(Directory.Exists(ProfilePath(record.Profile)));
                retainedJob.Dispose();
                await worker.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
                Assert.Equal(1, WorkerSandboxLease.Recover(root: _root));
            }
            finally { if (!host.HasExited) { host.Kill(); await host.WaitForExitAsync(); } }
        }
    }

    [Fact]
    public async Task FailedDeletionKeepsRecordAndConcurrentCollectorsRetryExactlyOnce()
    {
        var (host, record) = await StartHost("partial");
        using (host)
        {
            host.Kill(); await host.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            await WaitForLeaseRelease(record);
        }
        using (var locked = new FileStream(Path.Combine(record.Payload, "partial.dll"), FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            Assert.Throws<IOException>(() => WorkerSandboxLease.Recover(root: _root));
            Assert.True(File.Exists(Path.Combine(Path.GetDirectoryName(record.Payload)!, "lease")));
            Assert.True(Directory.Exists(ProfilePath(record.Profile)));
        }
        var results = await Task.WhenAll(Enumerable.Range(0, 6).Select(async _ =>
        {
            using var collector = Process.Start(HostStartInfo("recover", Worker))!;
            var result = await collector.StandardOutput.ReadToEndAsync().WaitAsync(TimeSpan.FromSeconds(10));
            await collector.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Equal(0, collector.ExitCode);
            return int.Parse(result.Trim());
        }));
        Assert.Equal(1, results.Sum());
        Assert.False(Directory.Exists(ProfilePath(record.Profile)));
    }

    [Fact]
    public async Task NextWorkerLaunchRecoversAnAbandonedRoundAutomatically()
    {
        var (host, record) = await StartHost("partial");
        using (host) { host.Kill(); await host.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5)); }
        await WaitForLeaseRelease(record);
        using var next = WorkerSandbox.Create(Worker, default, _root);
        Assert.False(Directory.Exists(record.Payload));
        Assert.False(Directory.Exists(ProfilePath(record.Profile)));
        Assert.Single(Directory.GetDirectories(_root));
    }

    [Fact]
    public async Task FixtureStartupFailureWritesDiagnosticAndExitsInsteadOfOpeningErrorReporting()
    {
        using var host = Process.Start(HostStartInfo("staged", Path.Combine(_root, "missing", "worker.exe")))!;
        var errors = host.StandardError.ReadToEndAsync();
        await host.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
        Assert.Equal(1, host.ExitCode);
        Assert.Contains("Exception", await errors);
        Assert.Empty(Directory.GetDirectories(_root));
    }

    [Fact]
    public void CleanupFailureReleasesLeaseForRecoveryAndDoesNotHideError()
    {
        var sandbox = WorkerSandbox.Create(Worker, default, _root);
        var profile = sandbox.ProfileName;
        using (var locked = new FileStream(sandbox.Executable, FileMode.Open, FileAccess.Read, FileShare.Read))
        { Assert.Throws<IOException>(() => sandbox.Dispose()); }
        Assert.Equal(1, WorkerSandboxLease.Recover(root: _root));
        Assert.False(Directory.Exists(ProfilePath(profile)));
    }

    [Fact]
    public void PartialRecordAndFailedStagingAreRecoverableButUnknownNamesAreUntouched()
    {
        WorkerSandboxLease.CreatePrivateDirectory(_root);
        var interrupted = Path.Combine(_root, "QuickExcel.Worker." + Guid.NewGuid().ToString("N"));
        WorkerSandboxLease.CreatePrivateDirectory(interrupted); // Crash before opening lease.
        var unrelated = Path.Combine(_root, "QuickExcel.Worker.not-owned");
        Directory.CreateDirectory(unrelated);
        try
        {
            Assert.Equal(1, WorkerSandboxLease.Recover(root: _root));
            Assert.True(Directory.Exists(unrelated));
            Assert.ThrowsAny<IOException>(() => WorkerSandbox.Create(Path.Combine(_root, "missing", "missing.exe"), default, _root));
            Assert.Equal([unrelated], Directory.GetDirectories(_root));
            using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => WorkerSandboxLease.Create(cancellation.Token, _root));
        }
        finally { Directory.Delete(unrelated); }
    }

    [Theory]
    [InlineData("entry")]
    [InlineData("payload")]
    [InlineData("root")]
    public void RecoveryRefusesJunctionsWithoutTouchingTheirTargets(string location)
    {
        var outside = Path.Combine(Path.GetTempPath(), "QuickExcel.RecoveryCanary." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        var canary = Path.Combine(outside, "canary.txt"); File.WriteAllText(canary, "preserve");
        string link;
        if (location == "root") link = _root;
        else
        {
            WorkerSandboxLease.CreatePrivateDirectory(_root);
            var entry = Path.Combine(_root, "QuickExcel.Worker." + Guid.NewGuid().ToString("N"));
            if (location == "entry") link = entry;
            else { WorkerSandboxLease.CreatePrivateDirectory(entry); link = Path.Combine(entry, "payload"); }
        }
        try
        {
            var start = new ProcessStartInfo("cmd.exe") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            start.ArgumentList.Add("/c"); start.ArgumentList.Add("mklink"); start.ArgumentList.Add("/J");
            start.ArgumentList.Add(link); start.ArgumentList.Add(outside);
            using var command = Process.Start(start)!; Assert.True(command.WaitForExit(5000)); Assert.Equal(0, command.ExitCode);
            Assert.Throws<IOException>(() => WorkerSandboxLease.Recover(root: _root));
            Assert.Equal("preserve", File.ReadAllText(canary));
        }
        finally
        {
            if (Directory.Exists(link)) Directory.Delete(link); // Remove only the junction, never its target.
            File.Delete(canary); Directory.Delete(outside);
        }
    }

    private async Task<(Process Host, HostRecord Record)> StartHost(string stage)
    {
        var host = Process.Start(HostStartInfo(stage, Worker))!;
        try
        {
            var line = await host.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(15));
            if (line is null) Assert.Fail(await host.StandardError.ReadToEndAsync().WaitAsync(TimeSpan.FromSeconds(1)));
            return (host, JsonSerializer.Deserialize<HostRecord>(line!)!);
        }
        catch
        {
            if (!host.HasExited) { host.Kill(); await host.WaitForExitAsync(); }
            host.Dispose(); throw;
        }
    }

    private ProcessStartInfo HostStartInfo(string stage, string worker)
    {
        var start = new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory, "Fixtures", "SandboxHostFixture", "SandboxHostFixture.exe"))
        { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        start.ArgumentList.Add(_root); start.ArgumentList.Add(stage); start.ArgumentList.Add(worker);
        return start;
    }

    private static string ProfilePath(string name) => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Packages", name);
    private static async Task WaitForLeaseRelease(HostRecord record)
    {
        // The process object can signal before Windows finishes closing its file handles.
        var timer = Stopwatch.StartNew();
        while (true)
        {
            try
            {
                using var lease = new FileStream(Path.Combine(Path.GetDirectoryName(record.Payload)!, "lease"), FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return;
            }
            catch (IOException e) when ((e.HResult & 0xffff) == 32 && timer.Elapsed < TimeSpan.FromSeconds(5)) { await Task.Delay(20); }
        }
    }
    private sealed record HostRecord(string Profile, string Payload, int ProcessId);
    public void Dispose()
    {
        if (!Directory.Exists(_root)) return;
        WorkerSandboxLease.Recover(root: _root);
        File.Delete(Path.Combine(_root, ".gate"));
        Directory.Delete(_root);
    }
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "OpenJobObjectW")]
    private static extern SafeFileHandle OpenJobObject(uint access, bool inherit, string name);
}
