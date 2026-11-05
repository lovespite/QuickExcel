using System.Text.Json;
using QuickExcel.Presentation;

namespace QuickExcel.Core.Test;

public sealed class RecentFileStoreTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "QuickExcel-history-" + Guid.NewGuid().ToString("N"));
    private string Config => Path.Combine(_directory, "recent-files.json");
    private string Workbook(int index) => Path.Combine(_directory, index + ".xlsx");
    public RecentFileStoreTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);

    [Fact]
    public void LegacyPathsLoadWithoutRewritingAndUpgradeWhenPinned()
    {
        var legacy = JsonSerializer.Serialize(new[] { Workbook(1), Workbook(2), Workbook(1) });
        File.WriteAllText(Config, legacy);
        var store = new RecentFileStore(Config); store.Load();
        Assert.Null(store.Error); Assert.Equal(100, store.Limit);
        Assert.Equal(new[] { Workbook(1), Workbook(2) }, store.Paths);
        Assert.All(store.Entries, entry => Assert.False(entry.IsPinned));
        Assert.Equal(legacy, File.ReadAllText(Config));
        store.SetPinned(Workbook(2), true);
        var reopened = new RecentFileStore(Config); reopened.Load();
        Assert.Null(reopened.Error); Assert.True(reopened.Entries[0].IsPinned);
        Assert.Equal(new[] { Workbook(2), Workbook(1) }, reopened.Paths);
        using var json = JsonDocument.Parse(File.ReadAllText(Config));
        Assert.Equal(1, json.RootElement.GetProperty("Version").GetInt32());
    }

    [Fact]
    public void LimitAndPinsPersistAndStaleWritersRespectBoth()
    {
        var first = new RecentFileStore(Config); var second = new RecentFileStore(Config);
        first.Remember(Workbook(1)); first.Remember(Workbook(2)); second.Load();
        first.SetPinned(Workbook(1), true); first.ChangeLimit(2);
        second.Remember(Workbook(3)); second.Remember(Workbook(4));
        Assert.Equal(2, second.Limit);
        Assert.Equal(new[] { Workbook(1), Workbook(4), Workbook(3) }, second.Paths);
        first.ChangeLimit(1);
        Assert.Equal(new[] { Workbook(1), Workbook(4) }, first.Paths);
        second.Remember(Workbook(1));
        Assert.Equal(new[] { Workbook(1), Workbook(4) }, second.Paths);
        Assert.True(second.Entries[0].IsPinned);
        second.SetPinned(Workbook(1), false);
        Assert.Equal(new[] { Workbook(1) }, second.Paths); Assert.False(second.Entries[0].IsPinned);
        second.Clear(); first.SetPinned(Workbook(1), true);
        Assert.Empty(first.Paths); Assert.Equal(1, first.Limit);
        first.Remember(Workbook(5)); second.Load();
        Assert.Equal(new[] { Workbook(5) }, second.Paths); Assert.Equal(1, second.Limit);
    }

    [Fact]
    public void PinnedOrderSurvivesOpeningAndRemovalOnlyRemovesSelectedRecord()
    {
        var store = new RecentFileStore(Config);
        for (var i = 1; i <= 3; i++) store.Remember(Workbook(i));
        store.SetPinned(Workbook(1), true); store.SetPinned(Workbook(2), true);
        store.Remember(Workbook(1)); store.Remember(Workbook(4));
        Assert.Equal(new[] { Workbook(2), Workbook(1), Workbook(4), Workbook(3) }, store.Paths);
        store.Remove(Workbook(2)); store.Load();
        Assert.Equal(new[] { Workbook(1), Workbook(4), Workbook(3) }, store.Paths);
        Assert.True(store.Entries[0].IsPinned);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(1001)]
    public void InvalidLimitsAreRejectedWithoutChangingPersistedHistory(int limit)
    {
        var store = new RecentFileStore(Config); store.Remember(Workbook(1));
        var before = File.ReadAllText(Config);
        Assert.Throws<ArgumentOutOfRangeException>(() => store.ChangeLimit(limit));
        Assert.Equal(before, File.ReadAllText(Config));
        File.WriteAllText(Config, JsonSerializer.Serialize(new { Version = 1, Limit = limit, Entries = Array.Empty<RecentFileEntry>() }));
        var invalid = File.ReadAllText(Config); store.Load(); Assert.NotNull(store.Error);
        store.Remember(Workbook(2)); Assert.NotNull(store.Error); Assert.Equal(invalid, File.ReadAllText(Config));
    }

    [Fact]
    public void ConfiguredLongUnicodeHistoryCanExceedTheOld64KiBBudget()
    {
        var entries = Enumerable.Range(0, 1000).Select(i => new RecentFileEntry(
            Path.Combine(_directory, new string('文', 80) + i + ".xlsx"))).ToArray();
        File.WriteAllText(Config, JsonSerializer.Serialize(new { Version = 1, Limit = 1000, Entries = entries }));
        Assert.True(new FileInfo(Config).Length > 64 * 1024);
        var store = new RecentFileStore(Config); store.Load();
        Assert.Null(store.Error); Assert.Equal(1000, store.Paths.Count);
        store.SetPinned(entries[^1].Path, true); store.Load();
        Assert.Null(store.Error); Assert.Equal(1000, store.Paths.Count); Assert.Equal(entries[^1].Path, store.Paths[0]);
    }

    [Theory]
    [InlineData("{broken")]
    [InlineData("{\"Version\":2,\"Limit\":100,\"Entries\":[]}")]
    [InlineData("{\"Version\":1,\"Limit\":100,\"Entries\":null}")]
    public void InvalidOrFutureSettingsAreNeverOverwritten(string content)
    {
        File.WriteAllText(Config, content);
        var store = new RecentFileStore(Config); store.Remember(Workbook(1));
        Assert.NotNull(store.Error); Assert.Empty(store.Paths); Assert.Equal(content, File.ReadAllText(Config));
    }
}
