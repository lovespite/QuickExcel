using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;

namespace QuickExcel.Runtime;

#pragma warning disable CA1416 // Windows-only worker launch path.

/// <summary>A durable ownership record, written before any OS resource is created.</summary>
internal sealed class WorkerSandboxLease
{
    internal static string Root => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "QuickExcel.WorkerSandboxes.v1");
    private readonly string _root;
    private FileStream? _lease;
    internal string DirectoryPath { get; }
    internal string ProfileName => Path.GetFileName(DirectoryPath);
    internal string PayloadDirectory => Path.Combine(DirectoryPath, "payload");
    internal string JobName => @"Global\" + ProfileName;

    private WorkerSandboxLease(string root, string directory, FileStream lease)
    { _root = root; DirectoryPath = directory; _lease = lease; }

    internal static WorkerSandboxLease Create(CancellationToken cancellation, string? root = null)
    {
        root ??= Root;
        using var gate = LockRoot(root, cancellation);
        RecoverLocked(root, cancellation);
        var directory = Path.Combine(root, "QuickExcel.Worker." + Guid.NewGuid().ToString("N"));
        CreatePrivateDirectory(directory);
        // A crash even before lease creation is recoverable: all resource names derive from this directory.
        return new(root, directory, OpenLease(directory));
    }

    internal static int Recover(CancellationToken cancellation = default, string? root = null)
    {
        root ??= Root;
        using var gate = LockRoot(root, cancellation);
        return RecoverLocked(root, cancellation);
    }

    private static int RecoverLocked(string root, CancellationToken cancellation)
    {
        var recovered = 0;
        var errors = new List<Exception>();
        foreach (var directory in Directory.EnumerateDirectories(root))
        {
            cancellation.ThrowIfCancellationRequested();
            var name = Path.GetFileName(directory);
            const string prefix = "QuickExcel.Worker.";
            // Never sweep legacy temporary directories or infer ownership from a loose prefix match.
            if (!name.StartsWith(prefix, StringComparison.Ordinal) ||
                !Guid.TryParseExact(name[prefix.Length..], "N", out var id) || name != prefix + id.ToString("N")) continue;
            try
            {
                ValidatePrivateDirectory(directory);
                FileStream lease;
                try { lease = OpenLease(directory); }
                catch (IOException e) when ((e.HResult & 0xffff) == 32) { continue; } // Another live host owns this round.
                using (lease) RemoveResources(directory);
                File.Delete(Path.Combine(directory, "lease"));
                Directory.Delete(directory);
                recovered++;
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException or System.ComponentModel.Win32Exception or COMException)
            { errors.Add(new IOException($"Cannot recover worker sandbox {name}; ownership record retained for retry.", e)); }
        }
        if (errors.Count != 0) throw new IOException("Worker sandbox recovery failed.", new AggregateException(errors));
        return recovered;
    }

    internal void Cleanup()
    {
        if (_lease is null) return;
        try
        {
            using var gate = LockRoot(_root, default);
            ValidatePrivateDirectory(DirectoryPath);
            RemoveResources(DirectoryPath);
            _lease.Dispose();
            File.Delete(Path.Combine(DirectoryPath, "lease"));
            Directory.Delete(DirectoryPath);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or System.ComponentModel.Win32Exception or COMException)
        { throw new IOException("Worker sandbox cleanup failed; ownership record retained for retry.", e); }
        finally
        {
            // On failure the durable record remains, but must not keep a dead round leased forever.
            _lease.Dispose();
            _lease = null;
        }
    }

    private static void RemoveResources(string directory)
    {
        WorkerJob.EnsureEmpty(@"Global\" + Path.GetFileName(directory));
        foreach (var item in Directory.EnumerateFileSystemEntries(directory))
            if (Path.GetFileName(item) is not ("lease" or "payload")) throw new IOException("Unexpected sandbox ownership record content.");
        var payload = Path.Combine(directory, "payload");
        var timer = Stopwatch.StartNew();
        while (true)
        {
            try { RemovePayload(payload); break; }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException &&
                (e.HResult & 0xffff) is 5 or 32 or 33 && timer.Elapsed < TimeSpan.FromSeconds(1))
            { Thread.Sleep(20); } // Image/AV handles can outlive the process's signaled state briefly.
        }
        var result = DeleteAppContainerProfile(Path.GetFileName(directory));
        if (result < 0) Marshal.ThrowExceptionForHR(result);
    }

    private static void RemovePayload(string payload)
    {
        if (Directory.Exists(payload))
        {
            WorkerSandbox.RejectReparsePoints(payload);
            var files = Directory.GetFileSystemEntries(payload);
            // Validate the complete flat payload before deleting anything. Never traverse a link or subdirectory.
            foreach (var file in files)
                if ((File.GetAttributes(file) & (FileAttributes.ReparsePoint | FileAttributes.Directory)) != 0)
                    throw new IOException("Worker payload contains a link or subdirectory.");
            foreach (var file in files)
            {
                File.SetAttributes(file, File.GetAttributes(file) & ~FileAttributes.ReadOnly);
                File.Delete(file);
            }
            Directory.Delete(payload);
        }
    }

    internal static void CreatePrivateDirectory(string directory)
    {
        WorkerSandbox.RejectReparsePoints(Path.GetDirectoryName(directory)!);
        var acl = new DirectorySecurity();
        using var identity = WindowsIdentity.GetCurrent();
        acl.SetOwner(identity.User!);
        acl.SetAccessRuleProtection(true, false);
        acl.AddAccessRule(new(identity.User!, FileSystemRights.FullControl,
            InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
        new DirectoryInfo(directory).Create(acl);
        ValidatePrivateDirectory(directory);
    }

    private static void ValidatePrivateDirectory(string directory)
    {
        WorkerSandbox.RejectReparsePoints(directory);
        using var identity = WindowsIdentity.GetCurrent();
        var acl = new DirectoryInfo(directory).GetAccessControl(AccessControlSections.Access | AccessControlSections.Owner);
        if (!acl.AreAccessRulesProtected || !identity.User!.Equals(acl.GetOwner(typeof(SecurityIdentifier))))
            throw new IOException("Worker ownership directory must be private and owned by the current user.");
        foreach (FileSystemAccessRule rule in acl.GetAccessRules(true, true, typeof(SecurityIdentifier)))
            if (rule.AccessControlType == AccessControlType.Allow && !identity.User.Equals(rule.IdentityReference))
                throw new IOException("Worker ownership directory grants access to another identity.");
    }

    private static FileStream OpenLease(string directory)
    {
        var path = Path.Combine(directory, "lease");
        RejectFileLink(path);
        return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    }

    private static FileStream LockRoot(string root, CancellationToken cancellation)
    {
        cancellation.ThrowIfCancellationRequested();
        if (!Path.IsPathFullyQualified(root)) throw new IOException("Worker ownership root must be absolute.");
        if (!Directory.Exists(root)) CreatePrivateDirectory(root);
        ValidatePrivateDirectory(root);
        var path = Path.Combine(root, ".gate");
        RejectFileLink(path);
        var timer = Stopwatch.StartNew();
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            try { return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException e) when ((e.HResult & 0xffff) == 32 && timer.Elapsed < TimeSpan.FromSeconds(5))
            { cancellation.WaitHandle.WaitOne(20); }
        }
    }

    private static void RejectFileLink(string path)
    {
        try
        {
            if ((File.GetAttributes(path) & (FileAttributes.ReparsePoint | FileAttributes.Directory)) != 0)
                throw new IOException("Worker ownership file cannot be a link or directory.");
        }
        catch (FileNotFoundException) { }
    }

    [DllImport("userenv.dll", CharSet = CharSet.Unicode)] private static extern int DeleteAppContainerProfile(string name);
}
