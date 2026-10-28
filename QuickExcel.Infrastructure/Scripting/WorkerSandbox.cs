using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;

namespace QuickExcel.Runtime;

// ScriptWorkerRuntime.Start checks Windows x64 before entering this internal launch path.
#pragma warning disable CA1416

/// <summary>A per-round identity and private copy of the executable payload. Installation ACLs are never changed.</summary>
internal sealed class WorkerSandbox : IDisposable
{
    private readonly WorkerSandboxLease _lease;
    private string _profile => _lease.ProfileName;
    private string _directory => _lease.PayloadDirectory;
    private bool _disposed;
    private WorkerSandbox(WorkerSandboxLease lease) => _lease = lease;
    internal IntPtr Sid { get; private set; }
    internal string Executable { get; private set; } = "";
    internal string ProfileName => _profile;
    internal string PayloadDirectory => _directory;
    internal string JobName => _lease.JobName;

    internal static WorkerSandbox Create(string executable, CancellationToken cancellation, string? ownershipRoot = null)
    {
        var sandbox = new WorkerSandbox(WorkerSandboxLease.Create(cancellation, ownershipRoot));
        try
        {
            cancellation.ThrowIfCancellationRequested();
            var result = CreateAppContainerProfile(sandbox._profile, sandbox._profile, "QuickExcel script worker", IntPtr.Zero, 0, out var sid);
            if (result < 0) Marshal.ThrowExceptionForHR(result);
            sandbox.Sid = sid;
            var acl = new DirectorySecurity();
            acl.SetAccessRuleProtection(true, false);
            using var identity = WindowsIdentity.GetCurrent();
            acl.AddAccessRule(new(identity.User!, FileSystemRights.FullControl,
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
            acl.AddAccessRule(new(new SecurityIdentifier(sid), FileSystemRights.ReadAndExecute,
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
            // The unpredictable directory is created with its final DACL, before any binaries are copied.
            if (Directory.Exists(sandbox._directory)) throw new IOException("Worker staging directory already exists.");
            new DirectoryInfo(sandbox._directory).Create(acl);
            var source = Path.GetDirectoryName(executable)!;
            RejectReparsePoints(source);
            foreach (var file in Directory.EnumerateFiles(source))
            {
                cancellation.ThrowIfCancellationRequested();
                if ((File.GetAttributes(file) & FileAttributes.ReparsePoint) != 0) throw new IOException("Worker payload cannot contain links.");
                if (Path.GetExtension(file).Equals(".pdb", StringComparison.OrdinalIgnoreCase)) continue;
                File.Copy(file, Path.Combine(sandbox._directory, Path.GetFileName(file)), overwrite: false);
            }
            sandbox.Executable = Path.Combine(sandbox._directory, Path.GetFileName(executable));
            if (!File.Exists(sandbox.Executable)) throw new FileNotFoundException("Worker executable was not staged.");
            return sandbox;
        }
        catch
        {
            sandbox.Dispose();
            throw;
        }
    }

    internal static void RejectReparsePoints(string path)
    {
        for (var item = new DirectoryInfo(path); item is not null; item = item.Parent)
            if ((item.Attributes & FileAttributes.ReparsePoint) != 0) throw new IOException("Worker path cannot traverse links.");
    }

    internal string EnvironmentBlock()
    {
        var sid = new SecurityIdentifier(Sid).Value;
        var result = GetAppContainerFolderPath(sid, out var path);
        if (result < 0) Marshal.ThrowExceptionForHR(result);
        string profile;
        try { profile = Marshal.PtrToStringUni(path) ?? throw new IOException("Missing AppContainer profile."); }
        finally { Marshal.FreeCoTaskMem(path); }
        // Do not forward secrets, proxy settings, PATH, .NET startup hooks or profiler variables.
        var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        var values = new SortedDictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["SystemRoot"] = windows, ["WINDIR"] = windows,
            ["LOCALAPPDATA"] = profile, ["TEMP"] = Path.Combine(profile, "Temp"), ["TMP"] = Path.Combine(profile, "Temp"),
            ["DOTNET_EnableDiagnostics"] = "0", ["DOTNET_CLI_TELEMETRY_OPTOUT"] = "1"
        };
        return string.Join('\0', values.Select(p => p.Key + "=" + p.Value)) + "\0\0";
    }

    public void Dispose()
    {
        if (_disposed) return;
        try { _lease.Cleanup(); }
        finally
        {
            if (Sid != IntPtr.Zero) { FreeSid(Sid); Sid = IntPtr.Zero; }
            _disposed = true; // Failed cleanup is retried through the durable ownership record.
        }
    }

    [DllImport("userenv.dll", CharSet = CharSet.Unicode)] private static extern int CreateAppContainerProfile(string name, string display, string description, IntPtr capabilities, uint count, out IntPtr sid);
    [DllImport("userenv.dll", CharSet = CharSet.Unicode)] private static extern int GetAppContainerFolderPath(string sid, out IntPtr path);
    [DllImport("advapi32.dll")] private static extern IntPtr FreeSid(IntPtr sid);
}
