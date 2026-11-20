using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Core.Test;

public sealed class ScriptNetworkSessionTests
{
    [Fact]
    public void DefaultsAndLiveLimitChangesApplyWithoutLosingConsent()
    {
        using var network = new ScriptNetworkSession { RecordingEnabled = true };
        Assert.Equal(new ScriptNetworkOptions(8, 10, 50, 200, true), network.Options);
        var prompts = 0; network.Authorize = (_, _) => { prompts++; return true; };
        network.EnsureAuthorized(["example.com"], default);
        var first = new ScriptNetworkEntry { ResponseBody = new byte[700 * 1024] };
        var second = new ScriptNetworkEntry { ResponseBody = new byte[700 * 1024] };
        network.Record(first); network.Record(second);
        network.Configure(network.Options with { BodyLimitMiB = 3, ActionTimeoutSeconds = 20, PanelMemoryLimitMiB = 1, PanelEntryLimit = 1 });
        Assert.Single(network.Entries); Assert.Null(first.ResponseBody); Assert.NotNull(second.ResponseBody);
        Assert.InRange(network.RetainedMemoryBytes, 1, 1024 * 1024);
        network.EnsureAuthorized(["example.com"], default); Assert.Equal(1, prompts);
    }

    [Fact]
    public void DisabledSwitchRejectsGrantsCancelsLeasesAndRequiresFreshConsentWhenEnabled()
    {
        using var network = new ScriptNetworkSession(); var prompts = 0;
        network.Authorize = (_, _) => { prompts++; return true; };
        network.EnsureAuthorized(["example.com"], default);
        var active = network.CaptureNetworkAccess(["example.com"]);
        network.Configure(network.Options with { AllowNetwork = false });
        Assert.True(active.IsCancellationRequested);
        Assert.Throws<InvalidOperationException>(() => network.EnsureAuthorized(["example.com"], default));
        Assert.Throws<InvalidOperationException>(() => network.CaptureNetworkAccess()); Assert.Equal(1, prompts);
        network.Configure(network.Options with { AllowNetwork = true });
        Assert.Throws<InvalidOperationException>(() => network.CaptureNetworkAccess(["example.com"]));
        network.EnsureAuthorized(["example.com"], default); Assert.Equal(2, prompts);
        Assert.False(network.CaptureNetworkAccess(["example.com"]).IsCancellationRequested);
    }

    [Fact]
    public void MetadataCannotBypassThePanelMemoryLimit()
    {
        using var network = new ScriptNetworkSession { RecordingEnabled = true };
        network.Configure(network.Options with { PanelMemoryLimitMiB = 1, PanelEntryLimit = 5000 });
        for (var i = 0; i < 100; i++) network.Record(new()
        {
            RequestHeaders = [new("x-long-header", new string('a', 32000))]
        });
        Assert.InRange(network.RetainedMemoryBytes, 1, 1024 * 1024);
        Assert.InRange(network.Entries.Length, 1, 20);
        network.Clear(); Assert.Equal(0, network.RetainedMemoryBytes);
    }

    [Fact]
    public void DisablingWhileConsentIsPendingCannotAcceptItsResult()
    {
        using var network = new ScriptNetworkSession();
        network.Authorize = (_, _) => { network.Configure(network.Options with { AllowNetwork = false }); return true; };
        Assert.Throws<OperationCanceledException>(() => network.EnsureAuthorized(["example.com"], default));
        Assert.Throws<InvalidOperationException>(() => network.CaptureNetworkAccess());
    }

    [Fact]
    public void DocumentRevocationDropsConsentButPreservesDiagnosticsUntilClose()
    {
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new WorksheetSnapshot(SheetId.New(), "Sheet1")]);
        var session = new DocumentSession(snapshot);
        session.SetScriptsEnabled(true);
        var network = session.Network;
        network.Authorize = (_, _) => true; network.EnsureAuthorized(["example.com"], default);
        Assert.Same(network, session.BeginConsole("1").Network);
        network.RecordingEnabled = true;
        var entry = new ScriptNetworkEntry { ResponseBody = new byte[3] }; network.Record(entry);
        session.SetScriptsEnabled(false); network.Authorize = null;
        Assert.Throws<InvalidOperationException>(() => network.EnsureAuthorized(["example.com"], default));
        Assert.Single(network.Entries); Assert.NotNull(entry.ResponseBody);
        session.Dispose(); Assert.Empty(network.Entries); Assert.Null(entry.ResponseBody);
    }

    [Fact]
    public void ConsentDefaultsToDenyAndOnlyExplicitGrantsAreReused()
    {
        using var network = new ScriptNetworkSession();
        network.EnsureAuthorized([], default);
        Assert.Throws<InvalidOperationException>(() => network.EnsureAuthorized(["example.com"], default));
        var prompts = new List<string[]>();
        network.Authorize = (hosts, _) => { prompts.Add(hosts.ToArray()); return true; };
        network.EnsureAuthorized(["example.org", "example.com", "example.com"], default);
        network.EnsureAuthorized(["example.com", "example.org"], default);
        Assert.Single(prompts); Assert.Equal(["example.com", "example.org"], prompts[0]);
        network.EnsureAuthorized(["new.example.com"], default); Assert.Equal(2, prompts.Count);
        Assert.Throws<InvalidOperationException>(() => network.CaptureNetworkAccess(["example.com"]));
        network.Revoke(); network.EnsureAuthorized(["example.com"], default); Assert.Equal(3, prompts.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChangedListDropsPreviousConsentEvenIfReplacementIsRejected(bool cancelled)
    {
        using var network = new ScriptNetworkSession();
        network.Authorize = (_, _) => true;
        network.EnsureAuthorized(["example.com", "example.org"], default);
        using var cancellation = new CancellationTokenSource();
        network.Authorize = (hosts, _) =>
        {
            Assert.Equal(["example.com"], hosts);
            if (cancelled) cancellation.Cancel();
            return false;
        };
        if (cancelled)
            Assert.Throws<OperationCanceledException>(() => network.EnsureAuthorized(["example.com"], cancellation.Token));
        else
            Assert.Throws<InvalidOperationException>(() => network.EnsureAuthorized(["example.com"], default));
        Assert.Throws<InvalidOperationException>(() => network.CaptureNetworkAccess(["example.com"]));
        Assert.Throws<InvalidOperationException>(() => network.CaptureNetworkAccess(["example.org"]));
        network.Authorize = null;
        Assert.Throws<InvalidOperationException>(() => network.EnsureAuthorized(["example.com", "example.org"], default));
    }

    [Fact]
    public void RemovingAllDeclarationsDropsConsentWithoutPrompting()
    {
        using var network = new ScriptNetworkSession();
        var prompts = 0; network.Authorize = (_, _) => { prompts++; return true; };
        network.EnsureAuthorized(["example.com"], default);
        network.EnsureAuthorized([], default); Assert.Equal(1, prompts);
        Assert.Throws<InvalidOperationException>(() => network.CaptureNetworkAccess(["example.com"]));
        network.EnsureAuthorized(["example.com"], default); Assert.Equal(2, prompts);
    }

    [Fact]
    public void EmptyDeclarationInvalidatesAnEarlierPendingPrompt()
    {
        using var network = new ScriptNetworkSession();
        network.Authorize = (_, _) => { network.EnsureAuthorized([], default); return true; };
        Assert.Throws<OperationCanceledException>(() => network.EnsureAuthorized(["example.com"], default));
        Assert.Throws<InvalidOperationException>(() => network.CaptureNetworkAccess(["example.com"]));
    }

    [Fact]
    public void ReapplyAndDraftInvalidationKeepConsentButCancelOldRequestsAndPendingPrompts()
    {
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new WorksheetSnapshot(SheetId.New(), "Sheet1")]);
        using var session = new DocumentSession(snapshot);
        session.SetScriptsEnabled(true);
        var network = session.Network;
        var prompts = 0; network.Authorize = (_, _) => { prompts++; return true; };
        network.EnsureAuthorized(["example.com"], default);
        var first = session.BeginCalculation();
        session.SetScriptsEnabled(true);
        Assert.True(first.RuntimeLifetime.IsCancellationRequested);
        network.EnsureAuthorized(["example.com"], default); Assert.Equal(1, prompts);
        var second = session.BeginCalculation();
        session.InvalidateScriptExecution();
        Assert.False(session.ScriptsEnabled); Assert.True(second.RuntimeLifetime.IsCancellationRequested);
        session.SetScriptsEnabled(true);
        network.EnsureAuthorized(["example.com"], default); Assert.Equal(1, prompts);
        network.Authorize = (_, _) => { session.SetScriptsEnabled(true); return true; };
        Assert.Throws<OperationCanceledException>(() => network.EnsureAuthorized(["example.org"], default));
        Assert.Throws<InvalidOperationException>(() => network.CaptureNetworkAccess(["example.org"]));
    }

    [Fact]
    public void RevocationOrCancellationDuringPromptCannotRestoreConsent()
    {
        using var network = new ScriptNetworkSession();
        network.Authorize = (_, _) => { network.Revoke(); return true; };
        Assert.Throws<OperationCanceledException>(() => network.EnsureAuthorized(["example.com"], default));
        using var cancellation = new CancellationTokenSource();
        network.Authorize = (_, _) => { cancellation.Cancel(); return true; };
        Assert.Throws<OperationCanceledException>(() => network.EnsureAuthorized(["example.com"], cancellation.Token));
        network.Authorize = null;
        Assert.Throws<InvalidOperationException>(() => network.EnsureAuthorized(["example.com"], default));
    }

    [Fact]
    public void MemoryBudgetReleasesOldBodiesEvenWhenRowsRemainReferenced()
    {
        using var network = new ScriptNetworkSession(maxEntries: 3, maxBodyBytes: 10) { RecordingEnabled = true };
        var first = new ScriptNetworkEntry { ResponseBody = new byte[6] };
        var second = new ScriptNetworkEntry { RequestBody = new byte[3], ResponseBody = new byte[3] };
        network.Record(first); network.Record(second);
        Assert.True(first.BodiesEvicted); Assert.Null(first.ResponseBody); Assert.Equal(6, first.ResponseBytes);
        Assert.False(second.BodiesEvicted); Assert.Equal(6, network.RetainedBodyBytes); Assert.Equal(2, network.Entries.Length);
        var oversized = new ScriptNetworkEntry { ResponseBody = new byte[11] };
        network.Record(oversized);
        Assert.True(second.BodiesEvicted); Assert.True(oversized.BodiesEvicted); Assert.Equal(0, network.RetainedBodyBytes);
        network.Record(new() { Url = "https://example.com/latest" });
        Assert.Equal(3, network.Entries.Length); Assert.DoesNotContain(first, network.Entries);
    }

    [Fact]
    public void ClosedPanelClearAndDisposalReleasePayloadsAndObserversCannotBreakActions()
    {
        var network = new ScriptNetworkSession();
        var skipped = new ScriptNetworkEntry { ResponseBody = new byte[6] };
        network.Record(skipped); Assert.Empty(network.Entries); Assert.Null(skipped.ResponseBody);
        network.RecordingEnabled = true;
        network.Changed += () => throw new InvalidOperationException("diagnostic observer");
        var recorded = new ScriptNetworkEntry { ResponseBody = new byte[6] };
        network.Record(recorded); network.Clear();
        Assert.Null(recorded.ResponseBody); Assert.Empty(network.Entries); Assert.Equal(0, network.RetainedBodyBytes);
        var disposed = new ScriptNetworkEntry { ResponseBody = new byte[6] }; network.Record(disposed);
        network.Dispose(); Assert.False(network.RecordingEnabled); Assert.Null(disposed.ResponseBody); Assert.Empty(network.Entries);
        Assert.Throws<ObjectDisposedException>(() => network.EnsureAuthorized(["example.com"], default));
    }
}
