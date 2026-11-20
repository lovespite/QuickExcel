using QuickExcel.Serialization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace QuickExcel.Security;

public sealed record ScriptSigningPreferences(string? KeyFilePath = null, string? KeyFileHash = null, string? Password = null)
{
    [JsonIgnore] public bool RememberKeyFile => KeyFilePath is not null;
    [JsonIgnore] public bool RememberPassword => KeyFileHash is not null && Password is not null;
    public string PasswordForFile(string path) => RememberPassword && KeyFileHash == ScriptSigningPreferencesStore.FileHash(path) ? Password! : "";
}

/// <summary>Explicitly remembered signing inputs, protected for this OS user and local settings path; never workbook data.</summary>
public sealed class ScriptSigningPreferencesStore
{
    private readonly string _path;
    private readonly TrustDataProtection _protection;
    internal sealed record Envelope(int Version, string Protection, byte[] Data);
    private const int MaximumBytes = 128 * 1024;
    public ScriptSigningPreferencesStore(string path) : this(path, TrustDataProtection.Create()) { }
    internal ScriptSigningPreferencesStore(string path, TrustDataProtection protection)
    { _path = Path.GetFullPath(path); _protection = protection; }
    private byte[] Purpose => Encoding.UTF8.GetBytes("QuickExcel.ScriptSigningPreferences.v1\0" +
        (OperatingSystem.IsWindows() ? _path.ToUpperInvariant() : _path));

    public ScriptSigningPreferences Read()
    {
        if (!File.Exists(_path)) return new();
        using var file = File.OpenRead(_path);
        if (file.Length > MaximumBytes) throw new InvalidDataException("Signing preferences exceed the local storage limit.");
        using var json = JsonDocument.Parse(file, new JsonDocumentOptions { AllowDuplicateProperties = false });
        var envelope = json.RootElement.Deserialize(InfrastructureJsonContext.Strict.SigningEnvelope);
        if (envelope is not { Version: 1, Data.Length: > 0 } || envelope.Protection != _protection.Id)
            throw new InvalidDataException("Invalid local signing preferences protection.");
        var plaintext = _protection.Unprotect(envelope.Data, Purpose);
        try
        {
            using var payload = JsonDocument.Parse(plaintext, new JsonDocumentOptions { AllowDuplicateProperties = false });
            var value = payload.RootElement.Deserialize(InfrastructureJsonContext.Strict.ScriptSigningPreferences) ?? throw new InvalidDataException("Invalid signing preferences.");
            Validate(value); return value;
        }
        finally { CryptographicOperations.ZeroMemory(plaintext); }
    }

    public void Save(string keyFile, string password, bool rememberKeyFile, bool rememberPassword)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        using var lease = new FileStream(_path + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        if (!rememberKeyFile && !rememberPassword) { File.Delete(_path); return; }
        var value = new ScriptSigningPreferences(rememberKeyFile ? Path.GetFullPath(keyFile) : null,
            rememberPassword ? FileHash(keyFile) : null, rememberPassword ? password : null);
        Validate(value);
        var plaintext = JsonSerializer.SerializeToUtf8Bytes(value, InfrastructureJsonContext.Strict.ScriptSigningPreferences);
        byte[] data;
        try { data = _protection.Protect(plaintext, Purpose); }
        finally { CryptographicOperations.ZeroMemory(plaintext); }
        var encoded = JsonSerializer.SerializeToUtf8Bytes(new Envelope(1, _protection.Id, data), InfrastructureJsonContext.Strict.SigningEnvelope);
        if (encoded.Length > MaximumBytes) throw new InvalidDataException("Signing preferences exceed the local storage limit.");
        var temporary = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { file.Write(encoded); file.Flush(true); }
            File.Move(temporary, _path, true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    internal static string FileHash(string path)
    {
        using var file = File.OpenRead(path);
        if (file.Length > 16 * 1024 * 1024) throw new InvalidDataException("Signing certificate exceeds 16 MiB.");
        return Convert.ToHexString(SHA256.HashData(file));
    }
    private static void Validate(ScriptSigningPreferences value)
    {
        if (value.KeyFilePath is { } path && (path.Length > 32768 || !Path.IsPathFullyQualified(path)) ||
            value.Password is { Length: > 4096 } || (value.Password is null) != (value.KeyFileHash is null) ||
            value.KeyFileHash is { } hash && (hash.Length != 64 || !hash.All(char.IsAsciiHexDigit)))
            throw new InvalidDataException("Invalid local signing preferences.");
    }
}
