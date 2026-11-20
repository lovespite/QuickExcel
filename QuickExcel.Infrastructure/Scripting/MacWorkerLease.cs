using QuickExcel.Serialization;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text.Json;

namespace QuickExcel.Runtime;

/// <summary>Only journaled, exclusively leased macOS workers may be reclaimed.</summary>
[SupportedOSPlatform("macos")]
internal sealed class MacWorkerLease : IDisposable
{
    private const UnixFileMode PrivateDirectory = UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute;
    private const UnixFileMode PayloadFile = UnixFileMode.UserRead | UnixFileMode.UserExecute;
    private const UnixFileMode PayloadData = UnixFileMode.UserRead;
    private const int ProcPidTbsdInfo = 3;
    private const int ProcPidTbsdInfoSize = 136;
    private const int StartTimeSecondsOffset = 120;
    private const int StartTimeMicrosecondsOffset = 128;
    internal static string Root
    {
        get
        {
            var parent = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData, Environment.SpecialFolderOption.DoNotVerify);
            if (string.IsNullOrWhiteSpace(parent) || !Path.IsPathFullyQualified(parent))
                throw new IOException("macOS worker records require an absolute per-user data directory.");
            return Path.Combine(parent, "QuickExcel.MacWorkers.v1");
        }
    }
    internal string DirectoryPath { get; }
    internal string PayloadDirectory => Path.Combine(DirectoryPath, "payload");
    internal string ProfileName { get; }
    private readonly FileStream _lease;
    private Record _record;
    internal sealed record Record(int SchemaVersion, string Id, string Profile, int ProcessId, ulong StartTime, Dictionary<string, string> Files);
    private MacWorkerLease(string directory, string id, FileStream lease)
    {
        DirectoryPath = directory; ProfileName = "macos-seatbelt-" + id; _lease = lease;
        _record = new(1, id, ProfileName, 0, 0, new(StringComparer.Ordinal));
    }

    internal static MacWorkerLease Create(string executable, CancellationToken cancellation)
    {
        using var gate = LockRoot(cancellation);
        RecoverLocked(cancellation);
        var id = Guid.NewGuid().ToString("N"); var directory = Path.Combine(Root, id);
        Directory.CreateDirectory(directory, PrivateDirectory);
        var instance = new MacWorkerLease(directory, id, new FileStream(Path.Combine(directory, "lease"), FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None));
        try
        {
            instance.Save(); // Ownership precedes copying and process creation.
            Directory.CreateDirectory(instance.PayloadDirectory, PrivateDirectory);
            var source = Path.GetDirectoryName(Path.GetFullPath(executable))!;
            NoLinks(source);
            long total = 0; var count = 0;
            foreach (var file in Directory.EnumerateFileSystemEntries(source))
            {
                cancellation.ThrowIfCancellationRequested(); NoLinks(file);
                if (Directory.Exists(file)) throw new IOException("Worker payload must contain flat regular files only.");
                var name = Path.GetFileName(file); var destination = Path.Combine(instance.PayloadDirectory, name);
                using var input = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read);
                total = checked(total + input.Length);
                if (++count > 512 || total > 256L * 1024 * 1024 || input.Length > 128 * 1024 * 1024) throw new IOException("Worker payload file is oversized.");
                using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { input.CopyTo(output); output.Flush(true); }
                using var copied = File.OpenRead(destination);
                instance._record.Files.Add(name, Convert.ToHexString(SHA256.HashData(copied)));
                File.SetUnixFileMode(destination, name is "QuickExcel.ScriptWorker" or "quickexcel-sandbox" || name.EndsWith(".dylib", StringComparison.Ordinal)
                    ? PayloadFile : PayloadData);
            }
            foreach (var name in new[] { "QuickExcel.ScriptWorker", "quickexcel-sandbox", "libquickexcel-sandbox.dylib", "libcoreclr.dylib",
                "libJavaScriptCore.dylib", "typescript.bundle.json", "JSC-BUILD-IDENTITY.txt"
            })
                if (!instance._record.Files.ContainsKey(name)) throw new IOException("macOS worker payload missing: " + name);
            instance.Save(); File.SetUnixFileMode(instance.PayloadDirectory, UnixFileMode.UserRead | UnixFileMode.UserExecute);
            return instance;
        }
        catch { instance._lease.Dispose(); throw; } // Durable record retained for verified recovery.
    }

    internal void Attach(Process process)
    {
        _record = _record with { ProcessId = process.Id, StartTime = ReadStartTime(process.Id) };
        Save();
    }
    private void Save()
    {
        var temporary = Path.Combine(DirectoryPath, "record.new");
        using (var file = new FileStream(temporary, FileMode.Create, FileAccess.Write, FileShare.None)) { JsonSerializer.Serialize(file, _record, InfrastructureJsonContext.Default.MacLeaseRecord); file.Flush(true); }
        File.Move(temporary, Path.Combine(DirectoryPath, "record.json"), true);
    }
    internal static ulong ReadStartTime(int pid)
    {
        var buffer = new byte[ProcPidTbsdInfoSize];
        var read = proc_pidinfo(pid, ProcPidTbsdInfo, 0, buffer, buffer.Length);
        if (read != buffer.Length) throw new IOException("Cannot read the macOS worker process identity.");
        return BitConverter.ToUInt64(buffer, StartTimeSecondsOffset) * 1_000_000UL + BitConverter.ToUInt64(buffer, StartTimeMicrosecondsOffset);
    }
    internal static int Recover(CancellationToken cancellation = default)
    { using var gate = LockRoot(cancellation); return RecoverLocked(cancellation); }
    private static FileStream LockRoot(CancellationToken cancellation = default)
    {
        cancellation.ThrowIfCancellationRequested();
        NoLinks(Path.GetDirectoryName(Root)!);
        Directory.CreateDirectory(Root, PrivateDirectory); NoLinks(Root);
        if ((File.GetUnixFileMode(Root) & (UnixFileMode.GroupRead | UnixFileMode.GroupWrite | UnixFileMode.GroupExecute | UnixFileMode.OtherRead | UnixFileMode.OtherWrite | UnixFileMode.OtherExecute)) != 0)
            throw new IOException("macOS worker records require a private user directory.");
        var path = Path.Combine(Root, "root.lock"); NoLinks(path);
        // Shared workers are cleaned up asynchronously on session invalidation.
        // A new round (or the other role) may start while that cleanup holds the
        // root lock. Serialize ownership changes, as on Windows, without making
        // transient lock contention a script failure or bypassing the lease.
        var timer = Stopwatch.StartNew();
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            NoLinks(path);
            try { return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
            // Darwin reports flock contention as EWOULDBLOCK (35).
            catch (IOException error) when ((error.HResult & 0xffff) == 35 && timer.Elapsed < TimeSpan.FromSeconds(5))
            { cancellation.WaitHandle.WaitOne(20); }
        }
    }
    private static int RecoverLocked(CancellationToken cancellation)
    {
        var recovered = 0;
        foreach (var directory in Directory.EnumerateDirectories(Root))
        {
            cancellation.ThrowIfCancellationRequested();
            var id = Path.GetFileName(directory);
            if (!Guid.TryParseExact(id, "N", out var guid) || id != guid.ToString("N")) continue;
            NoLinks(directory);
            var recordPath = Path.Combine(directory, "record.json");
            if (!File.Exists(recordPath)) continue; // Never infer ownership from a name alone.
            NoLinks(recordPath); NoLinks(Path.Combine(directory, "lease"));
            FileStream lease;
            try { lease = new FileStream(Path.Combine(directory, "lease"), FileMode.Open, FileAccess.ReadWrite, FileShare.None); }
            catch (IOException) { continue; }
            using (lease)
            {
                if (new FileInfo(recordPath).Length > 128 * 1024) throw new IOException("Invalid macOS worker record size.");
                var record = JsonSerializer.Deserialize(File.ReadAllText(recordPath), InfrastructureJsonContext.Default.MacLeaseRecord) ?? throw new IOException("Missing macOS ownership record.");
                if (record.SchemaVersion != 1 || record.Id != id || record.Profile != "macos-seatbelt-" + id) throw new IOException("Invalid macOS ownership record.");
                if (record.ProcessId > 0)
                {
                    try { if (ReadStartTime(record.ProcessId) == record.StartTime) continue; }
                    catch (IOException) { if (ProcessExists(record.ProcessId)) continue; }
                }
                RemoveOwnedFiles(directory);
            }
            File.Delete(Path.Combine(directory, "lease")); Directory.Delete(directory); recovered++;
        }
        return recovered;
    }
    private static void RemoveOwnedFiles(string directory)
    {
        // A round contains only a flat payload and these exact journal files.
        foreach (var entry in Directory.EnumerateFileSystemEntries(directory))
        {
            NoLinks(entry);
            if (Path.GetFileName(entry) is not ("payload" or "record.json" or "record.new" or "lease")) throw new IOException("Unknown macOS worker record content; retained.");
        }
        var payload = Path.Combine(directory, "payload");
        if (Directory.Exists(payload))
        {
            var files = Directory.GetFileSystemEntries(payload);
            foreach (var file in files) { NoLinks(file); if (Directory.Exists(file)) throw new IOException("Nested worker payload; retained."); }
            File.SetUnixFileMode(payload, PrivateDirectory);
            foreach (var file in files) File.Delete(file);
            Directory.Delete(payload);
        }
        File.Delete(Path.Combine(directory, "record.new")); File.Delete(Path.Combine(directory, "record.json"));
    }
    internal static void NoLinks(string path)
    {
        for (var current = Path.GetFullPath(path); current != Path.GetPathRoot(current); current = Path.GetDirectoryName(current)!)
            if (new FileInfo(current).LinkTarget is not null ||
                (File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("Linked macOS worker path is not allowed: " + current);
    }
    public void Dispose()
    {
        try
        {
            using var gate = LockRoot();
            if (_record.ProcessId > 0)
            {
                var live = false;
                try { live = ReadStartTime(_record.ProcessId) == _record.StartTime; }
                catch (IOException) { live = ProcessExists(_record.ProcessId); }
                if (live) throw new IOException("Live macOS worker payload cannot be removed.");
            }
            RemoveOwnedFiles(DirectoryPath);
            _lease.Dispose(); File.Delete(Path.Combine(DirectoryPath, "lease")); Directory.Delete(DirectoryPath);
        }
        finally { _lease.Dispose(); }
    }

    [DllImport("/usr/lib/libproc.dylib")]
    private static extern int proc_pidinfo(int pid, int flavor, ulong argument, byte[] buffer, int size);
    [DllImport("libc", SetLastError = true)]
    private static extern int kill(int pid, int signal);
    // If the identity cannot be read, keep a process that still exists. Never
    // infer that an unreadable round is dead merely because proc_pidinfo failed.
    private static bool ProcessExists(int pid) => kill(pid, 0) == 0 || Marshal.GetLastWin32Error() == 1;
}
