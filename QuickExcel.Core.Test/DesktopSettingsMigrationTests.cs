using QuickExcel.Presentation;
using QuickExcel.Core;

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
        Assert.Equal(DesktopCalculationMode.Automatic, settings.Value.Calculation);
        Assert.Equal(UiText.Chinese, settings.Value.Language);
        Assert.Equal(new ScriptNetworkOptions(), settings.Value.ScriptNetwork);
        settings.Save(settings.Value with { Calculation = DesktopCalculationMode.Manual, Language = UiText.English });
        var reloaded = new DesktopSettings(Path.Combine(Target, "settings.json")); reloaded.Load();
        Assert.Equal(DesktopCalculationMode.Manual, reloaded.Value.Calculation);
        Assert.Equal(UiText.English, reloaded.Value.Language);
        Assert.Single(Directory.GetDirectories(Target, "migration-backup-*"));
        Assert.False(File.Exists(Path.Combine(Target, "trusted-script-certificates.json")));
        File.WriteAllText(Path.Combine(Legacy, "settings.json"), "{\"Sheets\":0,\"Zoom\":75}");
        Assert.Null(DesktopSettingsMigration.ImportOnce(Legacy, Target)); settings.Load(); Assert.Equal(125, settings.Value.Zoom);
        Assert.True(File.Exists(Path.Combine(Target, "migration-v1.json")));
    }
    [Fact]
    public void NetworkLimitsRoundTripAndInvalidFilesFallBackToDefaults()
    {
        var path = Path.Combine(Target, "settings.json");
        var settings = new DesktopSettings(path); settings.Load();
        Assert.Equal(new ScriptNetworkOptions(8, 10, 50, 200), settings.Value.ScriptNetwork);
        var selected = new ScriptNetworkOptions(5, 25, 72, 350, AllowNetwork: false);
        settings.Save(settings.Value with { ScriptNetwork = selected });
        Assert.Null(settings.Error);
        var next = new DesktopSettings(path); next.Load();
        Assert.Equal(selected, next.Value.ScriptNetwork);
        Assert.False(next.Value.ScriptNetwork.AllowNetwork);
        Assert.Throws<ArgumentException>(() => next.Save(next.Value with { ScriptNetwork = selected with { BodyLimitMiB = 0 } }));
        Assert.Equal(selected, next.Value.ScriptNetwork);
        File.WriteAllText(path, "{\"ScriptNetwork\":{\"ActionTimeoutSeconds\":0}}"); next.Load();
        Assert.NotNull(next.Error); Assert.Equal(new ScriptNetworkOptions(), next.Value.ScriptNetwork);
        File.WriteAllText(path, "{\"ScriptNetwork\":null}"); next.Load();
        Assert.NotNull(next.Error); Assert.Equal(new ScriptNetworkOptions(), next.Value.ScriptNetwork);
    }

    [Fact]
    public void MigrationPreservesNetworkPreferencesAndRejectsInvalidLimits()
    {
        var legacyPath = Path.Combine(Legacy, "settings.json");
        File.WriteAllText(legacyPath, "{\"ScriptNetwork\":{\"PanelEntryLimit\":0}}");
        Assert.NotNull(DesktopSettingsMigration.ImportOnce(Legacy, Target));
        Assert.False(File.Exists(Path.Combine(Target, "settings.json")));
        File.WriteAllText(legacyPath, "{\"Zoom\":125,\"ScriptNetwork\":{\"BodyLimitMiB\":4,\"ActionTimeoutSeconds\":30,\"PanelMemoryLimitMiB\":32,\"PanelEntryLimit\":100}}");
        Assert.Null(DesktopSettingsMigration.ImportOnce(Legacy, Target));
        var migrated = new DesktopSettings(Path.Combine(Target, "settings.json")); migrated.Load();
        Assert.Equal(new ScriptNetworkOptions(4, 30, 32, 100), migrated.Value.ScriptNetwork); Assert.Equal(125, migrated.Value.Zoom);
    }
    [Fact]
    public void ExistingAvaloniaPreferencesAreNeverOverwrittenAndMalformedLegacyIsRetained()
    {
        File.WriteAllText(Path.Combine(Legacy, "settings.json"), "malformed");
        Assert.NotNull(DesktopSettingsMigration.ImportOnce(Legacy, Target));
        Assert.False(File.Exists(Path.Combine(Target, "settings.json")));
        Assert.Equal("malformed", File.ReadAllText(Path.Combine(Legacy, "settings.json")));
        File.WriteAllText(Path.Combine(Target, "settings.json"), "{\"Sheets\":2,\"Zoom\":150,\"Language\":\"unknown\"}");
        Assert.Null(DesktopSettingsMigration.ImportOnce(Legacy, Target));
        var settings = new DesktopSettings(Path.Combine(Target, "settings.json")); settings.Load(); Assert.Equal(150, settings.Value.Zoom);
        Assert.Equal(UiText.Chinese, settings.Value.Language);
    }
    public void Dispose() => Directory.Delete(_root, true);
}
