using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace QuickExcel.Runtime;

internal sealed class UserFileError(string code, string message) : Exception(message)
{
    public string Code { get; } = code;
    public static UserFileError Invalid() => new("InvalidArgument", "Invalid file operation or options.");
    public static UserFileError Limit() => new("BudgetExceeded", "File operation exceeds the Action budget.");
    public static UserFileError Conflict() => new("Conflict", "The selected item changed.");
    public static UserFileError Unsupported() => new("UnsupportedApi", "The selected storage does not support this operation.");
    public static UserFileError Denied() => new("PermissionDenied", "The selected access does not permit this operation.");
}

/// <summary>Hold the physical directory chain without write/delete sharing throughout each operation.
/// OPEN_REPARSE_POINT inspects each successive component before using it as a parent.</summary>
internal static class SelectedFileSystem
{
    internal readonly record struct Identity(uint Volume, ulong Index);
    internal sealed record Stamp(Identity Id, bool Directory, long Length, long Created, long Written, long Changed, uint Links);
    internal sealed record Target(string Path, Identity[] Parents);
    internal const uint Read = 0x80000000, Write = 0x40000000, Delete = 0x10000;
    private const uint Inspect = 0x80, Flags = 0x02200000; // BACKUP_SEMANTICS | OPEN_REPARSE_POINT

    [StructLayout(LayoutKind.Sequential)]
    private struct Info
    {
        internal uint Attributes, CreatedLow, CreatedHigh, AccessLow, AccessHigh, WrittenLow, WrittenHigh;
        internal uint Volume, LengthHigh, LengthLow, Links, IndexHigh, IndexLow;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct Basic { internal long Created, Access, Written, Changed; internal uint Attributes; }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
    private static extern SafeFileHandle CreateFileW(string path, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
    private static extern bool GetVolumeNameForVolumeMountPointW(string root, StringBuilder volume, uint capacity);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetFileInformationByHandle(SafeFileHandle handle, out Info info);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetFileInformationByHandleEx(SafeFileHandle handle, int kind, out Basic info, uint length);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetFileInformationByHandle(SafeFileHandle handle, int kind, ref byte value, uint length);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
    private static extern bool CreateDirectoryW(string path, IntPtr attributes);
    [DllImport("ktmw32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern SafeFileHandle CreateTransaction(IntPtr attributes, IntPtr uow, uint options, uint isolation, uint flags, uint timeout, string? description);
    [DllImport("ktmw32.dll", SetLastError = true)]
    private static extern bool CommitTransaction(SafeFileHandle transaction);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
    private static extern SafeFileHandle CreateFileTransactedW(string path, uint access, uint share, IntPtr security, uint disposition,
        uint flags, IntPtr template, SafeFileHandle transaction, IntPtr version, IntPtr reserved);

    private static Exception LastError() => Marshal.GetLastWin32Error() switch
    {
        2 or 3 => new UserFileError("ItemNotFound", "The selected item does not exist."),
        5 => UserFileError.Denied(),
        32 or 33 => new UserFileError("Locked", "The selected item is in use."),
        80 or 183 or 6800 or 145 => UserFileError.Conflict(),
        1 or 50 or 6704 or 6801 or 6805 or 6831 => UserFileError.Unsupported(),
        var code => new IOException("Local file operation failed.", new Win32Exception(code))
    };

    internal static string Name(string name)
    {
        if (name.Length is 0 or > 255 || name is "." or ".." || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
            name.Contains('\\') || name.Contains('/') || name.Contains(':') || name.EndsWith('.') || name.EndsWith(' ') ||
            name.Any(c => char.IsControl(c) || char.GetUnicodeCategory(c) == System.Globalization.UnicodeCategory.Format))
            throw UserFileError.Invalid();
        var stem = name.Split('.')[0].TrimEnd(' ').ToUpperInvariant();
        if (stem is "CON" or "AUX" or "PRN" or "NUL" or "CONIN$" or "CONOUT$" or "CLOCK$" ||
            stem.Length == 4 && (stem.StartsWith("COM") || stem.StartsWith("LPT")) && "0123456789¹²³".Contains(stem[3]))
            throw UserFileError.Invalid();
        return name;
    }

    private static SafeFileHandle Open(string path, uint access, uint share = 1)
    {
        var handle = CreateFileW(path, access | Inspect, share, IntPtr.Zero, 3, Flags, IntPtr.Zero);
        if (handle.IsInvalid) { var error = LastError(); handle.Dispose(); throw error; }
        return handle;
    }

    internal static Stamp Stat(SafeFileHandle handle)
    {
        if (!GetFileInformationByHandle(handle, out var info) ||
            !GetFileInformationByHandleEx(handle, 0, out var basic, (uint)Marshal.SizeOf<Basic>())) throw LastError();
        if ((info.Attributes & (uint)(FileAttributes.ReparsePoint | FileAttributes.Device)) != 0) throw UserFileError.Unsupported();
        return new(new(info.Volume, ((ulong)info.IndexHigh << 32) | info.IndexLow),
            (info.Attributes & (uint)FileAttributes.Directory) != 0, ((long)info.LengthHigh << 32) | info.LengthLow,
            basic.Created, basic.Written, basic.Changed, info.Links);
    }

    // The trusted workspace path may itself be a link or a network share. Identify
    // the actual open workbook without granting a script access to that path.
    internal static Identity? WorkbookIdentity(string path)
    {
        using var file = CreateFileW(path, Inspect, 7, IntPtr.Zero, 3, 0, IntPtr.Zero);
        if (!file.IsInvalid) return Stat(file).Id;
        var error = LastError();
        if (error is UserFileError { Code: "ItemNotFound" }) return null;
        throw error;
    }

    internal sealed class Chain : IDisposable
    {
        private readonly List<SafeFileHandle> _handles = [];
        internal Target Target { get; }
        internal string PhysicalPath { get; }

        internal Chain(string path, Identity[]? expected = null)
        {
            if (!OperatingSystem.IsWindows()) throw UserFileError.Unsupported();
            path = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
            var root = Path.GetPathRoot(path)!;
            if (root.Length != 3 || !char.IsAsciiLetter(root[0]) || root[1] != ':' ||
                new DriveInfo(root).DriveType is not (DriveType.Fixed or DriveType.Removable)) throw UserFileError.Unsupported();
            var parts = path[root.Length..].Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 128) throw UserFileError.Limit();
            var identities = new List<Identity>();
            // Drive letters can be reassigned while an operation is running. Use
            // the volume GUID throughout this chain, including the final syscall.
            var volume = new StringBuilder(64);
            if (!GetVolumeNameForVolumeMountPointW(root, volume, (uint)volume.Capacity)) throw UserFileError.Unsupported();
            var current = volume.ToString();
            try
            {
                for (int i = 0; i < parts.Length; i++)
                {
                    Name(parts[i]);
                    var handle = Open(current, 1);
                    _handles.Add(handle);
                    var stamp = Stat(handle);
                    if (!stamp.Directory) throw UserFileError.Conflict();
                    identities.Add(stamp.Id);
                    current = Path.Combine(current, parts[i]);
                }
                if (expected is not null && !identities.SequenceEqual(expected)) throw UserFileError.Conflict();
                Target = new(path, identities.ToArray());
                PhysicalPath = current;
            }
            catch { Dispose(); throw; }
        }

        internal SafeFileHandle Pin(uint access = 0) => Open(PhysicalPath, access);
        internal Stamp? Observe()
        {
            try { using var handle = Pin(); return Stat(handle); }
            catch (UserFileError e) when (e.Code == "ItemNotFound") { return null; }
        }
        public void Dispose() { for (int i = _handles.Count - 1; i >= 0; i--) _handles[i].Dispose(); }
    }

    internal static void MakeDirectory(string path)
    {
        if (!CreateDirectoryW(path, IntPtr.Zero)) throw LastError();
    }

    internal static void Remove(SafeFileHandle handle)
    {
        byte delete = 1;
        if (!SetFileInformationByHandle(handle, 4, ref delete, 1)) throw LastError();
    }

    // A checked overwrite must validate and modify the SAME object atomically. A path-based
    // check followed by File.Move(overwrite:true) would overwrite a concurrently replaced file.
    internal static Stamp Publish(Chain chain, Stamp? expected, MemoryStream draft, CancellationToken token)
    {
        using var transaction = CreateTransaction(IntPtr.Zero, IntPtr.Zero, 0, 0, 0, 0, null);
        if (transaction.IsInvalid) throw LastError();
        Identity published;
        using (var handle = CreateFileTransactedW(chain.PhysicalPath, Write | Inspect, 1, IntPtr.Zero,
            expected is null ? 1u : 3u, Flags, IntPtr.Zero, transaction, IntPtr.Zero, IntPtr.Zero))
        {
            if (handle.IsInvalid) throw LastError();
            var stamp = Stat(handle);
            if (stamp.Directory || stamp.Links > 1) throw UserFileError.Denied();
            if (expected is not null && stamp != expected) throw UserFileError.Conflict();
            published = stamp.Id;
            using var target = new FileStream(handle, FileAccess.Write);
            target.SetLength(draft.Length);
            draft.Position = 0;
            draft.CopyTo(target);
            target.Flush();
        }
        token.ThrowIfCancellationRequested();
        if (!CommitTransaction(transaction)) throw LastError();
        var result = chain.Observe();
        return result?.Id == published ? result : throw UserFileError.Conflict();
    }
}

