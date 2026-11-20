using QuickExcel.Serialization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

public sealed record RecoveryMetadata(int Version, Guid SessionId, Guid DocumentId, long Revision,
    string? SourcePath, string? SourceSha256, DateTimeOffset CapturedUtc, DateTimeOffset CompletedUtc, string PayloadSha256);
public sealed record RecoveryCandidate(RecoveryMetadata Metadata, string Slot, string BundleSha256, bool DamagedGeneration);
public sealed record RecoveryScan(IReadOnlyList<RecoveryCandidate> Items, IReadOnlyList<string> Notices);

/// <summary>Two independently verified generations per session. Leases exclude live instances; recovery never opens the original file.</summary>
public sealed class WorkbookRecoveryStore
{
    private const long MaxBytes = 256L * 1024 * 1024;
    private readonly string _root;
    private readonly IWorkbookStorage _storage;
    internal sealed record Owner(int Version, Guid SessionId, Guid DocumentId, bool Retired = false);
    private static readonly string[] Slots = ["draft0.qer", "draft1.qer"];

    public WorkbookRecoveryStore(string root, IWorkbookStorage? storage = null)
    { _root = Path.GetFullPath(root); _storage = storage ?? new TypedWorkbookStorage(); }

    private static void Regular(string path)
    {
        try
        {
            if (File.GetAttributes(path).HasFlag(FileAttributes.ReparsePoint)) throw new IOException("Recovery paths cannot be reparse points.");
        }
        catch (FileNotFoundException) { }
        catch (DirectoryNotFoundException) { }
    }
    private string Folder(Guid id)
    {
        if (id == Guid.Empty) throw new ArgumentException("A recovery session needs an identity.");
        Regular(_root); var folder = Path.Combine(_root, id.ToString("N")); Regular(folder); return folder;
    }
    private static string Child(string folder, string name) { var path = Path.Combine(folder, name); Regular(path); return path; }
    private static FileStream Lock(string folder, FileMode mode) => new(Child(folder, "lease.json"), mode, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.WriteThrough);
    private static void WriteOwner(FileStream lease, Owner owner)
    { lease.Position = 0; lease.SetLength(0); JsonSerializer.Serialize(lease, owner, InfrastructureJsonContext.Default.Owner); lease.Flush(true); }
    private static Owner ReadOwner(FileStream lease, Guid id)
    {
        if (lease.Length is 0 or > 4096) throw new InvalidDataException("Invalid recovery lease.");
        lease.Position = 0; var owner = JsonSerializer.Deserialize(lease, InfrastructureJsonContext.Default.Owner);
        if (owner is null || owner.Version != 1 || owner.SessionId != id || owner.DocumentId == Guid.Empty) throw new InvalidDataException("Invalid recovery owner.");
        return owner;
    }
    public RecoveryLease Start(Guid sessionId, Guid documentId)
    {
        if (documentId == Guid.Empty) throw new ArgumentException("A recovery document needs an identity.");
        var folder = Folder(sessionId); Directory.CreateDirectory(folder); Regular(folder);
        var exists = File.Exists(Child(folder, "lease.json")); var lease = Lock(folder, exists ? FileMode.Open : FileMode.CreateNew);
        try
        {
            if (exists && ReadOwner(lease, sessionId) is var previous && (!previous.Retired || previous.DocumentId != documentId))
                throw new IOException("An unfinished recovery session already exists.");
            var owner = new Owner(1, sessionId, documentId); WriteOwner(lease, owner); return new(this, folder, lease, sessionId, documentId);
        }
        catch { lease.Dispose(); throw; }
    }

    public sealed class RecoveryLease : IDisposable
    {
        private readonly WorkbookRecoveryStore _store;
        private readonly string _folder;
        private readonly FileStream _lease;
        private readonly Guid _session, _document;
        private readonly object _gate = new();
        private long _generation, _revision = -1;
        private bool _disposed, _retired;
        internal RecoveryLease(WorkbookRecoveryStore store, string folder, FileStream lease, Guid session, Guid document)
        { _store = store; _folder = folder; _lease = lease; _session = session; _document = document; }

        public RecoveryMetadata Capture(WorkbookSnapshot snapshot, string? sourcePath, FileVersion? version, DateTimeOffset capturedUtc)
        {
            lock (_gate)
            {
                ObjectDisposedException.ThrowIf(_disposed, this);
                if (_retired || snapshot.DocumentId != _document || snapshot.Revision < _revision) throw new InvalidOperationException("Stale recovery capture.");
                _store.Folder(_session);
                var payload = Child(_folder, "writing.xlsq"); var bundle = Child(_folder, "writing.tmp");
                try
                {
                    _store._storage.Save(payload, snapshot);
                    if (new FileInfo(payload).Length > MaxBytes) throw new IOException("Recovery copy exceeds 256 MiB.");
                    using var input = new FileStream(payload, FileMode.Open, FileAccess.Read, FileShare.Read);
                    var hash = Convert.ToHexString(SHA256.HashData(input)); input.Position = 0;
                    RecoveryMetadata metadata;
                    using (var output = new FileStream(bundle, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 65536, FileOptions.WriteThrough))
                    {
                        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
                        {
                            using (var target = zip.CreateEntry("workbook.xlsq", CompressionLevel.NoCompression).Open()) input.CopyTo(target);
                            metadata = new(1, _session, _document, snapshot.Revision, sourcePath, version?.Sha256, capturedUtc, DateTimeOffset.UtcNow, hash);
                            using var manifest = zip.CreateEntry("metadata.json").Open(); JsonSerializer.Serialize(manifest, metadata, InfrastructureJsonContext.Default.RecoveryMetadata);
                        }
                        output.Flush(true);
                    }
                    _ = _store.ReadBundle(_folder, "writing.tmp", new(1, _session, _document));
                    var targetPath = Child(_folder, Slots[_generation % 2]);
                    if (File.Exists(targetPath)) File.Replace(bundle, targetPath, null); else File.Move(bundle, targetPath);
                    _generation++; _revision = snapshot.Revision; return metadata;
                }
                finally { DeleteKnown(_folder, "writing.xlsq"); DeleteKnown(_folder, "writing.tmp"); }
            }
        }
        public void Discard()
        {
            lock (_gate)
            {
                ObjectDisposedException.ThrowIf(_disposed, this); _store.Folder(_session);
                foreach (var slot in Slots) DeleteKnown(_folder, slot);
                DeleteKnown(_folder, "writing.xlsq"); DeleteKnown(_folder, "writing.tmp");
                WriteOwner(_lease, new(1, _session, _document, true)); _retired = true;
            }
        }
        public void Dispose() { lock (_gate) { if (_disposed) return; _disposed = true; _lease.Dispose(); } }
    }

    private static void DeleteKnown(string folder, string name) { var path = Child(folder, name); if (File.Exists(path)) File.Delete(path); }
    private static string CopyPayload(Stream source, long expected, Stream? destination = null)
    {
        if (expected is < 0 or > MaxBytes) throw new InvalidDataException("Recovery payload exceeds its limit.");
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256); var buffer = new byte[65536]; long total = 0;
        int count;
        while ((count = source.Read(buffer)) != 0)
        {
            total += count; if (total > expected) throw new InvalidDataException("Recovery payload length mismatch.");
            hash.AppendData(buffer, 0, count); destination?.Write(buffer, 0, count);
        }
        if (total != expected) throw new InvalidDataException("Truncated recovery payload.");
        return Convert.ToHexString(hash.GetHashAndReset());
    }
    private RecoveryCandidate ReadBundle(string folder, string slot, Owner owner)
    {
        var path = Child(folder, slot);
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (file.Length > MaxBytes + 65536) throw new InvalidDataException("Recovery bundle exceeds its limit.");
        var digest = Convert.ToHexString(SHA256.HashData(file)); file.Position = 0;
        using var zip = new ZipArchive(file, ZipArchiveMode.Read);
        if (zip.Entries.Count != 2 || zip.GetEntry("metadata.json") is not { Length: > 0 and <= 32768 } manifest || zip.GetEntry("workbook.xlsq") is not { } payload || payload.Length > MaxBytes)
            throw new InvalidDataException("Invalid recovery bundle entries.");
        using var json = manifest.Open(); var jsonBytes = new byte[(int)manifest.Length]; json.ReadExactly(jsonBytes);
        if (json.ReadByte() != -1) throw new InvalidDataException("Recovery metadata exceeds its limit.");
        var metadata = JsonSerializer.Deserialize(jsonBytes, InfrastructureJsonContext.Default.RecoveryMetadata);
        if (metadata is null || metadata.Version != 1 || metadata.SessionId != owner.SessionId || metadata.DocumentId != owner.DocumentId || metadata.Revision < 0 || metadata.CompletedUtc < metadata.CapturedUtc)
            throw new InvalidDataException("Invalid recovery metadata.");
        using var bytes = payload.Open();
        if (CopyPayload(bytes, payload.Length) != metadata.PayloadSha256) throw new InvalidDataException("Recovery payload checksum mismatch.");
        return new(metadata, slot, digest, false);
    }
    private (RecoveryCandidate? Item, bool Damaged) Latest(string folder, Owner owner)
    {
        var valid = new List<RecoveryCandidate>(); var damaged = false;
        foreach (var slot in Slots)
        {
            if (!File.Exists(Child(folder, slot))) continue;
            try { valid.Add(ReadBundle(folder, slot, owner)); }
            catch (Exception ex) when (ex is IOException or JsonException or InvalidDataException or ArgumentException) { damaged = true; }
        }
        var latest = valid.OrderByDescending(c => c.Metadata.Revision).ThenByDescending(c => c.Metadata.CompletedUtc).FirstOrDefault();
        return (latest is null ? null : latest with { DamagedGeneration = damaged }, damaged);
    }
    public RecoveryScan Scan()
    {
        Regular(_root); if (!Directory.Exists(_root)) return new([], []);
        var items = new List<RecoveryCandidate>(); var notices = new List<string>(); var count = 0;
        foreach (var folder in Directory.EnumerateDirectories(_root))
        {
            if (++count > 10000) { notices.Add("Recovery directory scan limit reached."); break; }
            if (!Guid.TryParseExact(Path.GetFileName(folder), "N", out var id)) continue;
            try
            {
                Folder(id); if (!File.Exists(Child(folder, "lease.json"))) continue;
                using var lease = Lock(folder, FileMode.Open); var owner = ReadOwner(lease, id); if (owner.Retired) continue;
                var latest = Latest(folder, owner); if (latest.Item is { } item) items.Add(item);
                if (latest.Damaged) notices.Add(id.ToString("N") + ": damaged recovery generation; kept for inspection.");
            }
            catch (IOException ex) when ((ex.HResult & 0xffff) is 32 or 33 or 11) { /* A live lease is unavailable. */ }
            catch (IOException ex) { notices.Add(id + ": " + ex.Message); }
            catch (Exception ex) when (ex is UnauthorizedAccessException or JsonException or ArgumentException) { notices.Add(id + ": " + ex.Message); }
        }
        return new(items.OrderByDescending(c => c.Metadata.CompletedUtc).ToArray(), notices);
    }
    public WorkbookSnapshot Recover(RecoveryCandidate candidate)
    {
        var folder = Folder(candidate.Metadata.SessionId); using var lease = Lock(folder, FileMode.Open);
        var owner = ReadOwner(lease, candidate.Metadata.SessionId); if (owner.Retired) throw new InvalidOperationException("Recovery item was discarded.");
        var current = Latest(folder, owner).Item;
        if (current is null || current.BundleSha256 != candidate.BundleSha256) throw new InvalidOperationException("Recovery item changed. Refresh the list.");
        var extracted = Child(folder, "reading.xlsq");
        try
        {
            using (var zip = ZipFile.OpenRead(Child(folder, current.Slot)))
            using (var input = zip.GetEntry("workbook.xlsq")!.Open())
            using (var output = new FileStream(extracted, FileMode.Create, FileAccess.Write, FileShare.None))
                if (CopyPayload(input, zip.GetEntry("workbook.xlsq")!.Length, output) != current.Metadata.PayloadSha256) throw new InvalidDataException("Recovery payload changed.");
            var snapshot = _storage.Open(extracted);
            if (snapshot.DocumentId != current.Metadata.DocumentId || snapshot.Revision != current.Metadata.Revision) throw new InvalidDataException("Recovery workbook identity mismatch.");
            return snapshot;
        }
        finally { DeleteKnown(folder, "reading.xlsq"); }
    }
    public void Discard(RecoveryCandidate candidate)
    {
        var folder = Folder(candidate.Metadata.SessionId); using var lease = Lock(folder, FileMode.Open);
        var owner = ReadOwner(lease, candidate.Metadata.SessionId);
        if (owner.Retired) return;
        if (Latest(folder, owner).Item?.BundleSha256 != candidate.BundleSha256) throw new InvalidOperationException("Recovery item changed. Refresh the list.");
        foreach (var slot in Slots) DeleteKnown(folder, slot);
        WriteOwner(lease, owner with { Retired = true });
    }
}
