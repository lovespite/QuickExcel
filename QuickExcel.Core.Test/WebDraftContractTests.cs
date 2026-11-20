using System.Text.Json;
using QuickExcel.Domain;
using QuickExcel.Presentation.Web;

namespace QuickExcel.Core.Test;

public sealed class WebDraftContractTests
{
    private sealed class ManualClock : TimeProvider
    {
        public long Ticks;
        public override long TimestampFrequency => TimeSpan.TicksPerSecond;
        public override long GetTimestamp() => Ticks;
    }
    [Fact]
    public void ExpiredTransferPreservesConfirmedTextAndRequiresFreshTransfer()
    {
        var clock = new ManualClock(); var session = new ScriptDraftSession(ScriptModuleSet.Single("confirmed"), clock: clock);
        session.Observe("main", 2); var chunks = ScriptDraftSession.Split("main", new string('x', 60_000), 1, 2).ToArray();
        Assert.False(session.Receive(chunks[0])); clock.Ticks = TimeSpan.FromSeconds(11).Ticks;
        Assert.Throws<InvalidDataException>(() => session.Receive(chunks[1])); Assert.Equal("confirmed", session.Models["main"].Text);
        var retry = ScriptDraftSession.Split("main", "retry", 1, 2).Single(); Assert.True(session.Receive(retry));
        session.ConfirmFlush(new Dictionary<string, long> { ["main"] = 2 }); Assert.Equal("retry", session.BuildModules(ScriptModuleSet.Single("confirmed")).PrimarySource);
    }
    private static readonly HashSet<string> Allowed = ["textChanged", "draftChunk", "draftAck"];
    [Fact]
    public void ProtocolRejectsOtherViewsDocumentsGenerationsCapabilitiesAndDuplicateFields()
    {
        var body = WebMessage.Create("view", "doc", 3, "draftAck", "r1", new Dictionary<string, object?> { ["versions"] = new Dictionary<string, object?> { ["main"] = 1 } });
        Assert.Equal("draftAck", WebMessage.Parse(body, "view", "doc", 3, Allowed).Type);
        Assert.Throws<InvalidDataException>(() => WebMessage.Parse(body, "older", "doc", 3, Allowed));
        Assert.Throws<InvalidDataException>(() => WebMessage.Parse(body, "view", "other", 3, Allowed));
        Assert.Throws<InvalidDataException>(() => WebMessage.Parse(body, "view", "doc", 2, Allowed));
        Assert.Throws<InvalidDataException>(() => WebMessage.Parse(body, "view", "doc", 3, new HashSet<string> { "requestPrint" }));
        Assert.Throws<InvalidDataException>(() => WebMessage.Parse(body.Replace("\"main\":1", "\"main\":1,\"main\":2"), "view", "doc", 3, Allowed));
        Assert.Throws<InvalidDataException>(() => WebMessage.Parse(body.Replace("\"generation\":3,", ""), "view", "doc", 0, Allowed));
        Assert.Throws<InvalidDataException>(() => WebMessage.Create("view", "doc", 3, "draftAck", "r1", new Dictionary<string, object?> { ["text"] = new string('中', 100_000) }));
    }
    [Fact]
    public void LargeUnicodeDraftRequiresEveryVerifiedChunkAndAFlushBarrier()
    {
        var original = ScriptModuleSet.Single("old"); var session = new ScriptDraftSession(original); var text = string.Concat(Enumerable.Repeat("中文😀\n", 20_000));
        session.Observe("main", 2); var chunks = ScriptDraftSession.Split("main", text, 1, 2).ToArray(); Assert.True(chunks.Length > 1);
        foreach (var chunk in chunks[..^1]) { Assert.False(session.Receive(chunk)); Assert.Equal("old", session.Models["main"].Text); }
        Assert.Throws<InvalidOperationException>(() => session.ConfirmFlush(new Dictionary<string, long> { ["main"] = 2 }));
        Assert.True(session.Receive(chunks[^1])); Assert.Equal(text, session.Models["main"].Text);
        Assert.Throws<InvalidOperationException>(() => session.BuildModules(original));
        session.ConfirmFlush(new Dictionary<string, long> { ["main"] = 2 }); var updated = session.BuildModules(original);
        Assert.Equal(text, updated.PrimarySource); Assert.True(session.IsDirty); session.MarkCommitted(updated); Assert.False(session.IsDirty);
        Assert.Throws<InvalidOperationException>(() => session.BuildModules(updated));
    }
    [Fact]
    public void EditsDuringFlushNeverAuthorizeAnOlderSource()
    {
        var session = new ScriptDraftSession(ScriptModuleSet.Single("old")); session.Observe("main", 2);
        var chunks = ScriptDraftSession.Split("main", new string('x', 60_000), 1, 2).ToArray(); session.Receive(chunks[0]);
        session.Observe("main", 3); session.Receive(chunks[1]);
        Assert.True(session.HasPending); Assert.Throws<InvalidOperationException>(() => session.ConfirmFlush(new Dictionary<string, long> { ["main"] = 2 }));
        session.Receive(ScriptDraftSession.Split("main", "latest", 2, 3).Single());
        session.ConfirmFlush(new Dictionary<string, long> { ["main"] = 3 }); Assert.False(session.HasPending);
    }
    [Theory]
    [InlineData("reorder")]
    [InlineData("hash")]
    [InlineData("replay")]
    [InlineData("identity")]
    public void InvalidTransferNeverReplacesConfirmedText(string kind)
    {
        var session = new ScriptDraftSession(ScriptModuleSet.Single("confirmed")); session.Observe("main", 2);
        var chunks = ScriptDraftSession.Split("main", new string('x', 60_000), 1, 2).ToArray();
        if (kind == "reorder") Assert.Throws<InvalidDataException>(() => session.Receive(chunks[1]));
        else if (kind == "hash")
        {
            var invalid = chunks.Select(c => c with { Sha256 = new string('0', 64) }).ToArray(); session.Receive(invalid[0]);
            Assert.Throws<InvalidDataException>(() => session.Receive(invalid[1]));
        }
        else { session.Receive(chunks[0]); Assert.Throws<InvalidDataException>(() => session.Receive(kind == "replay" ? chunks[0] : chunks[1] with { Transfer = "other" })); }
        Assert.Equal("confirmed", session.Models["main"].Text); Assert.True(session.HasPending);
    }
    [Fact]
    public void CrashRecoveryKeepsConfirmedDraftAndRequiresFreshBarrier()
    {
        var modules = ScriptModuleSet.Single("committed"); var session = new ScriptDraftSession(modules);
        session.Observe("main", 2); session.Receive(ScriptDraftSession.Split("main", "confirmed draft", 1, 2).Single()); session.Observe("main", 3);
        session.MarkFailed(); Assert.Throws<InvalidOperationException>(() => session.BuildModules(modules));
        session.RecoverConfirmed(); Assert.Equal("confirmed draft", session.Models["main"].Text); Assert.True(session.IsDirty); Assert.False(session.HasPending);
        Assert.Throws<InvalidOperationException>(() => session.BuildModules(modules));
        session.ConfirmFlush(new Dictionary<string, long> { ["main"] = 1 }); Assert.Equal("confirmed draft", session.BuildModules(modules).PrimarySource);
    }
    [Fact]
    public void AllModulesShareExistingOneMiBSourceBudget()
    {
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule> { ["main"] = new("scripts/main.js", new string('x', 800_000)), ["other"] = new("scripts/other.js", "") }, ["main"], scriptApiVersion: 3);
        var session = new ScriptDraftSession(modules); session.Observe("other", 2);
        var chunks = ScriptDraftSession.Split("other", new string('y', 300_000), 1, 2).ToArray(); foreach (var chunk in chunks[..^1]) session.Receive(chunk);
        Assert.Throws<InvalidDataException>(() => session.Receive(chunks[^1])); Assert.Equal("", session.Models["other"].Text);
        Assert.Throws<InvalidDataException>(() => new ScriptDraftSession(modules, true).Observe("main", 2));
    }
    [Fact]
    public void ConcurrentWorkbookSourceChangeCannotBeOverwrittenByOldPage()
    {
        var modules = ScriptModuleSet.Single("old"); var session = new ScriptDraftSession(modules);
        session.ConfirmFlush(new Dictionary<string, long> { ["main"] = 1 });
        Assert.Throws<InvalidOperationException>(() => session.BuildModules(modules.WithPrimarySource("external")));
    }
}
