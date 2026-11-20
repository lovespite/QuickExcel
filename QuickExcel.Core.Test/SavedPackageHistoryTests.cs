using System.Collections.Immutable;
using System.Text;
using System.Text.Json.Nodes;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class SavedPackageHistoryTests : IDisposable
{
    private const string Marker = "savedPackageHistoryMarker";
    private readonly string _directory = Path.Combine(AppContext.BaseDirectory, "test-data", "saved-package-history-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();

    public SavedPackageHistoryTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);

    [Theory]
    [InlineData("//xlsq/manifest.json")]
    [InlineData("//xlsq/state/document.json")]
    public void SavingPreservesUndoableRawPackageMetadata(string part)
    {
        var path = Seed();
        using var document = new WorkbookDocument(_storage.Open(path), new(_storage, new UnusedCalculator()));
        document.Session.SetScriptsEnabled(true);
        var sessionId = document.Session.SessionId;
        WriteMarker(document, part, "first");
        document.Save(path);
        WriteMarker(document, part, "second");
        document.Save(path);
        Assert.Equal("second", ReadMarker(_storage.Open(path), part));

        Assert.True(document.Undo());
        Assert.Equal("first", ReadMarker(document.Snapshot, part));
        document.Save(path);
        Assert.Equal("first", ReadMarker(_storage.Open(path), part));

        Assert.True(document.Redo());
        Assert.Equal("second", ReadMarker(document.Snapshot, part));
        document.Save(path);
        Assert.Equal("second", ReadMarker(_storage.Open(path), part));
        Assert.Equal(sessionId, document.Session.SessionId);
        Assert.True(document.ScriptsEnabled);
        Assert.False(document.IsDirty);
    }

    [Theory]
    [InlineData("//xlsq/manifest.json")]
    [InlineData("//xlsq/state/document.json")]
    public async Task RawPackageEditsDuringSaveKeepTheirOwnStateAndRemainDirty(string part)
    {
        var path = Seed();
        using var storage = new GatedStorage(_storage);
        using var document = new WorkbookDocument(_storage.Open(path), new(storage, new UnusedCalculator()));
        document.Session.SetScriptsEnabled(true);
        WriteMarker(document, part, "captured");
        var capturedRevision = document.Snapshot.Revision;
        var save = Task.Run(() => document.Save(path));
        try
        {
            await storage.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            WriteMarker(document, part, "edited during save");
        }
        finally { storage.Release.Set(); }
        await save.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.Equal("captured", ReadMarker(_storage.Open(path), part));
        Assert.Equal("edited during save", ReadMarker(document.Snapshot, part));
        Assert.Equal(capturedRevision, document.Session.SavedRevision);
        Assert.True(document.IsDirty);
        Assert.True(document.Undo());
        Assert.Equal("captured", ReadMarker(document.Snapshot, part));
        Assert.True(document.Redo());
        Assert.Equal("edited during save", ReadMarker(document.Snapshot, part));
        document.Save(path);
        Assert.Equal("edited during save", ReadMarker(_storage.Open(path), part));
        Assert.False(document.IsDirty);
    }

    private string Seed()
    {
        var modules = new ScriptModuleSet([KeyValuePair.Create("main", new ScriptModule("scripts/main.js", "export function main() {}"))],
            ["main"], bindings: [new("main", "actions")]);
        var path = Path.Combine(_directory, "book.xlsq");
        _storage.Save(path, Book(Sheet("Data")).WithScriptModules(modules));
        return path;
    }

    private void WriteMarker(WorkbookDocument document, string part, string value)
    {
        var json = JsonNode.Parse(_storage.Capture(document.Snapshot).Entries[part].Data.AsSpan())!.AsObject();
        json[Marker] = value;
        var request = document.Session.BeginAction("editPackageMetadata");
        var result = document.Session.CommitAction(request, new(request.OperationId, request.Stamp.InputRevision, request.Name!, [], [])
        {
            FileEdits = [new("write", part, Encoding.UTF8.GetBytes(json.ToJsonString()).ToImmutableArray())]
        });
        Assert.True(result.Changed);
    }

    private string? ReadMarker(WorkbookSnapshot snapshot, string part) =>
        JsonNode.Parse(_storage.Capture(snapshot).Entries[part].Data.AsSpan())![Marker]?.GetValue<string>();

    private sealed class UnusedCalculator : IWorkbookCalculator
    {
        public CalculationResult Calculate(CalculationRequest request, CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("These save/history tests must not execute calculation or scripts.");
    }

    private sealed class GatedStorage(TypedWorkbookStorage inner) : IWorkbookSaveStateStorage, IWorkbookVirtualFiles, IDisposable
    {
        private int _saves;
        public TaskCompletionSource Entered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public ManualResetEventSlim Release { get; } = new();
        public WorkbookSnapshot Open(string path) => inner.Open(path);
        public void Save(string path, WorkbookSnapshot snapshot) => inner.Save(path, snapshot);
        public void Save(string path, WorkbookSnapshot snapshot, bool discardUnsupported) => inner.Save(path, snapshot, discardUnsupported);
        public FileVersion SaveChecked(string path, WorkbookSnapshot snapshot, bool discardUnsupported, FileVersion expected) =>
            inner.SaveChecked(path, snapshot, discardUnsupported, expected);
        public void Export(string path, WorkbookSnapshot snapshot) => inner.Export(path, snapshot);
        public void Export(string path, WorkbookSnapshot snapshot, bool allowStaleCaches, bool discardUnsupported) =>
            inner.Export(path, snapshot, allowStaleCaches, discardUnsupported);
        public VirtualFileSet Capture(WorkbookSnapshot snapshot) => inner.Capture(snapshot);
        public WorkbookSnapshot Apply(WorkbookSnapshot snapshot, IReadOnlyList<VirtualFileEdit> edits) => inner.Apply(snapshot, edits);
        public WorkbookSaveResult SaveWithState(string path, WorkbookSnapshot snapshot, bool discardUnsupported,
            FileVersion? expected = null, bool overwriteOriginal = false)
        {
            if (Interlocked.Increment(ref _saves) == 1)
            {
                Entered.SetResult();
                if (!Release.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("The test did not release the captured save.");
            }
            return inner.SaveWithState(path, snapshot, discardUnsupported, expected, overwriteOriginal);
        }
        public void Dispose() => Release.Dispose();
    }
}
