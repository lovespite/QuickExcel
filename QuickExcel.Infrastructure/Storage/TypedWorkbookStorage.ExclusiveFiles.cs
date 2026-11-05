using System.Security.Cryptography;
using System.Text;
using QuickExcel.Core;

namespace QuickExcel.Storage;

public sealed partial class TypedWorkbookStorage
{
    private readonly Dictionary<string, ExclusiveFile> _exclusiveFiles = new(
        OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);

    /// <summary>Own an XLSQ file until the workspace closes, detaches or changes its file binding.</summary>
    public IDisposable? AcquireExclusiveFile(string path, bool allowMissing = false)
    {
        if (!Path.GetExtension(path).Equals(".xlsq", StringComparison.OrdinalIgnoreCase)) return null;
        return AcquireFile(path, allowMissing, readOnly: false);
    }

    /// <summary>A shared read handle needs no writes or companion file in the source directory.</summary>
    public IDisposable AcquireReadOnlyFile(string path) => AcquireFile(path, allowMissing: false, readOnly: true);

    // XLSX/CSV retain shared-open semantics, but the desktop still detects read-only sources.
    public static void CheckWriteAccess(string path)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete);
        var probe = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(path))!, ".quickexcel-write-" + Guid.NewGuid().ToString("N") + ".tmp");
        using var directoryAccess = new FileStream(probe, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1, FileOptions.DeleteOnClose);
    }

    private IDisposable AcquireFile(string path, bool allowMissing, bool readOnly)
    {
        var fullPath = Path.GetFullPath(path);
        lock (_exclusiveFiles)
        {
            if (_exclusiveFiles.ContainsKey(fullPath)) throw FileInUse();
            var lease = new ExclusiveFile(fullPath, allowMissing, readOnly, () =>
            {
                lock (_exclusiveFiles) _exclusiveFiles.Remove(fullPath);
            });
            _exclusiveFiles.Add(fullPath, lease);
            return lease;
        }
    }

    private ExclusiveFile? OwnedFile(string path)
    {
        lock (_exclusiveFiles) return _exclusiveFiles.GetValueOrDefault(Path.GetFullPath(path));
    }

    public FileVersion ReadVersion(string path) => OwnedFile(path)?.ReadVersion() ?? FileVersions.Read(path);

    private void RequireVersion(string path, FileVersion expected)
    {
        if (ReadVersion(path) != expected) throw new WorkbookFileConflictException();
    }

    private static IOException FileInUse(Exception? inner = null) => new(MessageText.Get(
        "此 XLSQ 文件正在被其他窗口或程序使用。请先关闭占用它的工作簿，再重试。"), inner);

    private sealed class ExclusiveFile : IDisposable
    {
        private readonly string _path;
        private readonly FileStream? _gate;
        private readonly Action _released;
        private FileStream? _content;
        private bool _disposed;

        public bool IsReadOnly { get; }
        public ExclusiveFile(string path, bool allowMissing, bool readOnly, Action released)
        {
            _path = path; _released = released; IsReadOnly = readOnly;
            // This OS-owned lease spans atomic replacement, when Windows requires the
            // data handle to close. DeleteOnClose also releases it after a process crash;
            // a leftover filename is never treated as evidence of a live owner.
            var identity = OperatingSystem.IsWindows() ? path.ToUpperInvariant() : path;
            var name = ".quickexcel-" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity))) + ".lock";
            try
            {
                if (!readOnly) _gate = new FileStream(Path.Combine(Path.GetDirectoryName(path)!, name), FileMode.OpenOrCreate,
                    FileAccess.ReadWrite, FileShare.None, 1, FileOptions.DeleteOnClose);
                try { _content = OpenContent(allowMissing); }
                catch { _gate?.Dispose(); throw; }
            }
            catch (IOException ex) when ((ex.HResult & 0xffff) is 32 or 33) { throw FileInUse(ex); }
        }

        private FileStream? OpenContent(bool allowMissing)
        {
            try { return new(_path, FileMode.Open, IsReadOnly ? FileAccess.Read : FileAccess.ReadWrite, IsReadOnly ? FileShare.Read : FileShare.None); }
            catch (FileNotFoundException) when (allowMissing) { return null; }
        }

        public byte[] Read(FileReadBudget budget)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _content ??= OpenContent(false);
            _content!.Position = 0;
            return PackageIO.ReadFile(_content, budget);
        }

        public FileVersion ReadVersion()
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            _content ??= OpenContent(true);
            if (_content is null) return new(null);
            _content.Position = 0;
            return FileVersions.Read(_content);
        }

        public void Commit(string temp, FileVersion? expected)
        {
            if (IsReadOnly) throw new InvalidOperationException(MessageText.Get(DocumentSession.ReadOnlyMessage));
            if (expected is not null && ReadVersion() != expected) throw new WorkbookFileConflictException();
            var written = FileVersions.Read(temp);
            _content?.Dispose(); _content = null;
            try
            {
                // Keep the companion lease while replacing, and retain version checks
                // for programs that do not participate in QuickExcel's lease protocol.
                if (expected is not null) FileVersions.Require(_path, expected);
                if (File.Exists(_path)) File.Replace(temp, _path, null);
                else File.Move(temp, _path);
            }
            finally
            {
                try { _content = OpenContent(true); }
                catch (UnauthorizedAccessException)
                {
                    // Permissions can change after opening. Keep the old file protected
                    // even when a failed replacement can only re-acquire read access.
                    _content = new(_path, FileMode.Open, FileAccess.Read, FileShare.None);
                }
            }
            // Do not mark the document saved if an external program raced the re-open.
            if (ReadVersion() != written) throw new WorkbookFileConflictException();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { _content?.Dispose(); }
            finally { try { _gate?.Dispose(); } finally { _released(); } }
        }
    }
}
