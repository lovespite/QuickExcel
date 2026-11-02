using QuickExcel.Presentation;

namespace QuickExcel.Core.Test;

public sealed class DesktopSettingsMigrationTests : IDisposable
{
    private readonly string _root = Path.Combine(AppContext.BaseDirectory, "test-data", "settings-migration-" + Guid.NewGuid().ToString("N"));
    private string Legacy => Path.Combine(_root, "old");
    private string Target => Path.Combine(_root, "new");
    public DesktopSettingsMigrationTests() { Directory.CreateDirectory(Legacy); Directory.CreateDirectory(Target); }
    [Fact]
    public void CopiesOnlyCompatiblePreferencesAndBacksUpBeforePublishingWithoutMovingTrust()
    {
        File.WriteAllText(Path.Combine(Legacy, "settings.json"), "{\"Sheets\":1,\"Zoom\":125}");
        File.WriteAllText(Path.Combine(Legacy, "recent-files.json"), "[]");
        File.WriteAllText(Path.Combine(Legacy, "trusted-script-certificates.json"), "private");
        Assert.Null(DesktopSettingsMigration.ImportOnce(Legacy, Target));
        var settings = new DesktopSettings(Path.Combine(Target, "settings.json")); settings.Load(); Assert.Equal(125, settings.Value.Zoom);
        Assert.Single(Directory.GetDirectories(Target, "migration-backup-*"));
        Assert.False(File.Exists(Path.Combine(Target, "trusted-script-certificates.json")));
        File.WriteAllText(Path.Combine(Legacy, "settings.json"), "{\"Sheets\":0,\"Zoom\":75}");
        Assert.Null(DesktopSettingsMigration.ImportOnce(Legacy, Target)); settings.Load(); Assert.Equal(125, settings.Value.Zoom);
        Assert.True(File.Exists(Path.Combine(Target, "migration-v1.json")));
    }
    [Fact]
    public void ExistingAvaloniaPreferencesAreNeverOverwrittenAndMalformedLegacyIsRetained()
    {
        File.WriteAllText(Path.Combine(Legacy, "settings.json"), "malformed");
        Assert.NotNull(DesktopSettingsMigration.ImportOnce(Legacy, Target));
        Assert.False(File.Exists(Path.Combine(Target, "settings.json")));
        Assert.Equal("malformed", File.ReadAllText(Path.Combine(Legacy, "settings.json")));
        File.WriteAllText(Path.Combine(Target, "settings.json"), "{\"Sheets\":2,\"Zoom\":150}");
        Assert.Null(DesktopSettingsMigration.ImportOnce(Legacy, Target));
        var settings = new DesktopSettings(Path.Combine(Target, "settings.json")); settings.Load(); Assert.Equal(150, settings.Value.Zoom);
    }
    public void Dispose() => Directory.Delete(_root, true);
}
