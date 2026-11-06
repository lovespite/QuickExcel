using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace QuickExcel.Runtime;

/// <summary>Fail-closed macOS arm64 launcher: private payload, Seatbelt profile, verified denials and watchdog.</summary>
[SupportedOSPlatform("macos")]
internal sealed class MacWorkerProcess : IWorkerProcess
{
    private MacWorkerLease? _lease;
    private Process? _process;
    public Stream Input => _process!.StandardInput.BaseStream;
    public Stream Output => _process!.StandardOutput.BaseStream;
    public Stream Error => _process!.StandardError.BaseStream;
    public int Id => _process!.Id;
    public string ProfileName => _lease!.ProfileName;
    public string PayloadDirectory => _lease!.PayloadDirectory;
    internal bool WaitForExit(int milliseconds) => _process!.WaitForExit(milliseconds);
    internal int ExitCode => _process!.ExitCode;
    [DllImport("libc")] private static extern uint geteuid();
    [DllImport("quickexcel-sandbox", EntryPoint = "qx_sandbox_check", CharSet = CharSet.Ansi)]
    private static extern int SandboxCheck(int pid, string operation);
    [DllImport("quickexcel-sandbox", EntryPoint = "qx_sandbox_check_path", CharSet = CharSet.Ansi)]
    private static extern int SandboxCheckPath(int pid, string operation, string path);
    [DllImport("/usr/lib/libproc.dylib", EntryPoint = "proc_pidpath")]
    private static extern int ProcessPath(int pid, byte[] buffer, uint size);

    internal static MacWorkerProcess Start(string executable, CancellationToken cancellation, bool compilation = false)
    {
        ScriptExecutionPolicy.Current.RequireEnabled();
        if (geteuid() == 0) throw new PlatformNotSupportedException("macOS worker 必须在普通用户会话中运行。");
        var child = new MacWorkerProcess();
        try
        {
            cancellation.ThrowIfCancellationRequested();
            child._lease = MacWorkerLease.Create(executable, cancellation);
            var start = new ProcessStartInfo(Path.Combine(child.PayloadDirectory, "quickexcel-sandbox"))
            { UseShellExecute = false, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true, WorkingDirectory = child.PayloadDirectory };
            start.Environment.Clear();
            start.ArgumentList.Add(child.PayloadDirectory);
            start.ArgumentList.Add(Environment.ProcessId.ToString(CultureInfo.InvariantCulture));
            if (compilation) start.ArgumentList.Add("compile");
            cancellation.ThrowIfCancellationRequested();
            ScriptExecutionPolicy.Current.RequireEnabled();
            child._process = Process.Start(start) ?? throw new IOException("Cannot start macOS sandbox launcher.");
            child._lease.Attach(child._process);
            using var registration = cancellation.Register(child.Kill);
            cancellation.ThrowIfCancellationRequested();
            child.Input.WriteByte((byte)'S'); child.Input.Flush();
            return child;
        }
        catch { child.Dispose(); throw; }
    }

    /// <summary>The host asks the kernel about the launched pid; no worker-supplied status is trusted.</summary>
    internal void VerifyBeforeSource()
    {
        var path = new byte[4096];
        if (ProcessPath(Id, path, (uint)path.Length) <= 0) throw new IOException("macOS worker executable path is unavailable.");
        var executable = Path.GetFullPath(Path.Combine(PayloadDirectory, "QuickExcel.ScriptWorker"));
        var actual = System.Text.Encoding.UTF8.GetString(path, 0, Array.IndexOf(path, (byte)0) is var end and > 0 ? end : 0);
        if (!string.Equals(Path.GetFullPath(actual), executable, StringComparison.Ordinal)) throw new IOException("macOS worker executable differs from its private payload.");
        var read = SandboxCheckPath(Id, "file-read-data", "/etc/passwd");
        var write = SandboxCheckPath(Id, "file-write-create", Path.Combine(PayloadDirectory, ".qx-host-probe"));
        var network = SandboxCheckPath(Id, "network-outbound", "127.0.0.1:9");
        var fork = SandboxCheck(Id, "process-fork");
        if (read != 1 || write != 1 || network != 1 || fork != 1)
            throw new IOException($"macOS worker sandbox state differs from required policy (read={read}, write={write}, network={network}, fork={fork}).");
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
