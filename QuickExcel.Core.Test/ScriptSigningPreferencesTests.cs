using System.Security.Cryptography;
using System.Text.Json;
using QuickExcel.Security;

namespace QuickExcel.Core.Test;

public sealed class ScriptSigningPreferencesTests : IDisposable
{
    private readonly string _directory = Path.Combine(AppContext.BaseDirectory, "test-data", "signing-preferences-" + Guid.NewGuid().ToString("N"));
    public ScriptSigningPreferencesTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);
    private string PathFor(string name) => Path.Combine(_directory, name);
    private string KeyFile() { var path = PathFor("key.pfx"); File.WriteAllBytes(path, [1, 2, 3]); return path; }

    [Fact]
    public void WindowsPreferencesAreEncryptedAndBoundToTheirLocalLocation()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = PathFor("prefs.json"); var key = KeyFile(); var store = new ScriptSigningPreferencesStore(path);
        store.Save(key, "test-secret-123", true, true);
        var stored = File.ReadAllText(path);
        Assert.DoesNotContain("test-secret-123", stored); Assert.DoesNotContain("key.pfx", stored);
        var reopened = new ScriptSigningPreferencesStore(path).Read();
        Assert.Equal(key, reopened.KeyFilePath); Assert.Equal("test-secret-123", reopened.PasswordForFile(key));
        var copy = PathFor("copied.json"); File.Copy(path, copy);
        Assert.Throws<CryptographicException>(() => new ScriptSigningPreferencesStore(copy).Read());
        using var json = JsonDocument.Parse(stored);
        var data = json.RootElement.GetProperty("Data").GetBytesFromBase64(); data[^1] ^= 1;
        File.WriteAllText(path, JsonSerializer.Serialize(new { Version = 1, Protection = "dpapi-current-user-v1", Data = data }));
        Assert.Throws<CryptographicException>(() => store.Read());
    }

    [Fact]
    public void ChoicesAreIndependentAndUncheckingRemovesStoredValues()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = PathFor("prefs.json"); var key = KeyFile(); var store = new ScriptSigningPreferencesStore(path);
        Assert.False(store.Read().RememberKeyFile); Assert.False(store.Read().RememberPassword);
        store.Save(key, "secret", true, false);
        Assert.True(store.Read().RememberKeyFile); Assert.Null(store.Read().Password);
        store.Save(key, "secret", false, true);
        var passwordOnly = store.Read(); Assert.Null(passwordOnly.KeyFilePath); Assert.Equal("secret", passwordOnly.PasswordForFile(key));
        File.WriteAllBytes(key, [4, 5, 6]); Assert.Equal("", passwordOnly.PasswordForFile(key));
        store.Save(key, "unused", false, false);
        Assert.False(File.Exists(path)); Assert.False(store.Read().RememberPassword);
    }

    [Fact]
    public void FailedProtectionPreservesPriorSettingsWithoutPlaintextFallback()
    {
        var path = PathFor("prefs.json"); var key = KeyFile(); File.WriteAllText(path, "previous encrypted state");
        var store = new ScriptSigningPreferencesStore(path, new UnavailableProtection());
        Assert.Throws<UserProtectionUnavailableException>(() => store.Save(key, "secret", true, true));
        Assert.Equal("previous encrypted state", File.ReadAllText(path));
        Assert.Empty(Directory.EnumerateFiles(_directory, "*.tmp"));
        store.Save(key, "unused", false, false); Assert.False(File.Exists(path));
    }

    private sealed class UnavailableProtection : TrustDataProtection
    {
        internal override string Id => "unavailable";
        internal override byte[] Protect(byte[] data, byte[] purpose) => throw new UserProtectionUnavailableException("Unavailable");
        internal override byte[] Unprotect(byte[] data, byte[] purpose) => throw new UserProtectionUnavailableException("Unavailable");
    }
}
