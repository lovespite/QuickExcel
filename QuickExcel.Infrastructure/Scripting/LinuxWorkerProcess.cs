using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace QuickExcel.Runtime;

/// <summary>Fail-closed Linux x64 launcher: private payload, cgroup barrier, Landlock and seccomp.</summary>
[SupportedOSPlatform("linux")]
internal sealed class LinuxWorkerProcess : IWorkerProcess
{
    private LinuxWorkerLease? _lease;
    private Process? _process;
    public Stream Input => _process!.StandardInput.BaseStream;
    public Stream Output => _process!.StandardOutput.BaseStream;
    public Stream Error => _process!.StandardError.BaseStream;
    public int Id => _process!.Id;
    public string ProfileName => _lease!.ScopeName;
    public string PayloadDirectory => _lease!.PayloadDirectory;
    internal bool WaitForExit(int milliseconds) => _process!.WaitForExit(milliseconds);
    internal int ExitCode => _process!.ExitCode;
    [DllImport("libc")] private static extern uint geteuid();

    internal static LinuxWorkerProcess Start(string executable, CancellationToken cancellation)
    {
        if (geteuid() == 0) throw new PlatformNotSupportedException("Linux worker 必须在普通用户会话中运行。");
        var child = new LinuxWorkerProcess();
        try
        {
            cancellation.ThrowIfCancellationRequested();
            child._lease = LinuxWorkerLease.Create(executable, cancellation);
            var start = new ProcessStartInfo(Path.Combine(child.PayloadDirectory, "quickexcel-sandbox"))
            { UseShellExecute = false, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, WorkingDirectory = child.PayloadDirectory };
            start.Environment.Clear();
            start.ArgumentList.Add(child.PayloadDirectory); start.ArgumentList.Add(Environment.ProcessId.ToString(CultureInfo.InvariantCulture));
            child._process = Process.Start(start) ?? throw new IOException("Cannot start Linux sandbox launcher.");
            child._lease.Attach(child._process);
            using var registration = cancellation.Register(child.Kill);
            child.AttachScope(cancellation);
            cancellation.ThrowIfCancellationRequested();
            child.Input.WriteByte((byte)'S'); child.Input.Flush();
            return child;
        }
        catch { child.Dispose(); throw; }
    }

    private void AttachScope(CancellationToken cancellation)
    {
        var bus = $"/run/user/{geteuid()}/bus";
        if (!File.Exists("/usr/bin/busctl") || !File.Exists(bus)) throw new PlatformNotSupportedException("Linux 脚本需要 systemd 用户会话与 cgroup v2 内存/进程限额；当前环境不具备，执行保持关闭。");
        var start = new ProcessStartInfo("/usr/bin/busctl") { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true };
        start.Environment.Clear(); start.Environment["LANG"] = "C.UTF-8";
        string[] arguments = [$"--address=unix:path={bus}", "--timeout=3s", "call", "org.freedesktop.systemd1", "/org/freedesktop/systemd1", "org.freedesktop.systemd1.Manager", "StartTransientUnit", "ssa(sv)a(sa(sv))",
            ProfileName, "fail", "8", "PIDs", "au", "1", Id.ToString(CultureInfo.InvariantCulture), "MemoryMax", "t", "268435456", "MemorySwapMax", "t", "0", "TasksMax", "t", "32",
            "RuntimeMaxUSec", "t", "5000000", "TimeoutStopUSec", "t", "100000", "KillMode", "s", "control-group", "CollectMode", "s", "inactive-or-failed", "0"];
        foreach (var argument in arguments) start.ArgumentList.Add(argument);
        using var busctl = Process.Start(start) ?? throw new IOException("Cannot contact systemd user manager.");
        var error = busctl.StandardError.ReadToEndAsync(cancellation);
        try
        {
            busctl.WaitForExitAsync(cancellation).WaitAsync(TimeSpan.FromSeconds(4), cancellation).GetAwaiter().GetResult();
            if (busctl.ExitCode != 0) throw new IOException("Linux worker cgroup 创建失败：" + error.GetAwaiter().GetResult());
        }
        finally { if (!busctl.HasExited) { busctl.Kill(); busctl.WaitForExit(1000); } }
        // Inspect kernel state, not merely a successful D-Bus reply. Worker has
        // received no byte and cannot execute managed or workbook code yet.
        string? membership = null;
        var timer = Stopwatch.StartNew();
        do
        {
            cancellation.ThrowIfCancellationRequested();
            membership = File.ReadAllLines($"/proc/{Id}/cgroup").SingleOrDefault(line => line.StartsWith("0::", StringComparison.Ordinal));
            if (membership?.EndsWith("/" + ProfileName, StringComparison.Ordinal) == true) break;
            if (cancellation.WaitHandle.WaitOne(10)) cancellation.ThrowIfCancellationRequested();
        } while (timer.Elapsed < TimeSpan.FromSeconds(1));
        if (membership is null || !membership.EndsWith("/" + ProfileName, StringComparison.Ordinal)) throw new IOException("Worker cgroup membership mismatch.");
        var group = Path.GetFullPath("/sys/fs/cgroup" + membership[3..]);
        if (!group.StartsWith("/sys/fs/cgroup/", StringComparison.Ordinal)) throw new IOException("Invalid worker cgroup path.");
        foreach (var (name, expected) in new[] { ("memory.max", "268435456"), ("memory.swap.max", "0"), ("pids.max", "32") })
            if (File.ReadAllText(Path.Combine(group, name)).Trim() != expected) throw new IOException("Worker cgroup limit not enforced: " + name);
        if (File.ReadAllLines(Path.Combine(group, "cgroup.procs")).Any(pid => pid != Id.ToString(CultureInfo.InvariantCulture))) throw new IOException("Unexpected process in private worker cgroup.");
    }
    internal void VerifyBeforeSource()
    {
        var status = File.ReadAllLines($"/proc/{Id}/status").Select(line => line.Split(':', 2)).Where(p => p.Length == 2).ToDictionary(p => p[0], p => p[1].Trim(), StringComparer.Ordinal);
        if (status["NoNewPrivs"] != "1" || status["Seccomp"] != "2" || status["CapEff"] != "0000000000000000" || status["CapPrm"] != "0000000000000000")
            throw new IOException("Linux worker protection state differs from required policy.");
    }
    public void Kill()
    {
        try { if (_process is { HasExited: false }) { _process.Kill(); _process.WaitForExit(2000); } }
        catch (InvalidOperationException) { } catch (Win32Exception) when (_process?.HasExited != false) { }
    }
    public void Dispose()
    {
        Kill(); _process?.Dispose(); _process = null;
        _lease?.Dispose(); _lease = null;
    }
}
