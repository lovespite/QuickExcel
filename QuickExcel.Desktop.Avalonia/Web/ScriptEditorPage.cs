using System.Text.Json;
using QuickExcel.Domain;
using QuickExcel.Infrastructure.Web;
using QuickExcel.Presentation.Web;

namespace QuickExcel.Desktop.Avalonia.Web;

internal sealed class ScriptEditorPage : IAsyncDisposable
{
    internal LocalWebView Browser { get; }
    internal ScriptDraftSession Drafts { get; }
    internal event Action<Exception>? Failed;
    internal event Action? SaveRequested, FocusRequested, DraftChanged;
    private Task? _flush;
    private bool _closed;
    private int _holds;
    private bool _locked;
    internal IAsyncDisposable HoldEdits() { _holds++; return new EditHold(this); }
    private sealed class EditHold(ScriptEditorPage page) : IAsyncDisposable
    {
        private ScriptEditorPage? _page = page;
        public async ValueTask DisposeAsync()
        { if (Interlocked.Exchange(ref _page, null) is { } owner && --owner._holds == 0) await owner.ReleaseAsync(); }
    }
    private async Task ReleaseAsync()
    { if (_locked && !_closed && Browser.IsReady) { await Browser.RequestAsync("releaseDraft", "released", new { }); _locked = false; } }
    private sealed record Changed(string Module, long Version);
    private sealed record Flushed(Dictionary<string, long>? Versions, string? Busy);
    internal ScriptEditorPage(WebAssetCatalog assets, ScriptDraftSession drafts, string document, long generation, string profileRoot)
    {
        Drafts = drafts;
        string[] capabilities = ["draftChanged", "draftChunk", "draftAck", "initialized", "loadAck", "loaded", "released", "selected", "located", "save", "focusHost"];
        Browser = new(assets, "script.html", document, generation, profileRoot, Program.WebTestMode ? capabilities.Append("testResult") : capabilities);
        Browser.Message += OnMessage;
        Browser.Failed += ex => { Drafts.MarkFailed(); Failed?.Invoke(ex); };
    }
    private async void OnMessage(WebMessage message)
    {
        try
        {
            switch (message.Type)
            {
                case "draftChanged":
                    var changed = message.Payload.Deserialize<Changed>(WebMessage.Json) ?? throw new InvalidDataException("缺少草稿版本。");
                    Drafts.Observe(changed.Module, changed.Version); DraftChanged?.Invoke(); break;
                case "draftChunk":
                    var chunk = message.Payload.Deserialize<DraftChunk>(WebMessage.Json) ?? throw new InvalidDataException("缺少草稿分块。");
                    Drafts.Receive(chunk);
                    await Browser.SendAsync("chunkAck", new { chunk.Transfer, chunk.Index, chunk.Version }, message.RequestId); break;
                case "save": SaveRequested?.Invoke(); break;
                case "focusHost": FocusRequested?.Invoke(); break;
                default: throw new InvalidDataException("脚本页发送了未请求的响应。");
            }
        }
        catch (Exception ex) { Drafts.MarkFailed(); Failed?.Invoke(ex); }
    }
    internal async Task LoadAsync(string module)
    {
        await Browser.Ready;
        await Browser.RequestAsync("initialize", "initialized", new { modules = Drafts.Models.Keys.ToArray(), readOnly = Drafts.IsReadOnly });
        foreach (var (id, draft) in Drafts.Models)
        foreach (var chunk in ScriptDraftSession.Split(id, draft.Text, 0, 1))
            await Browser.RequestAsync("loadChunk", "loadAck", chunk);
        await Browser.RequestAsync("activate", "loaded", new { module }, TimeSpan.FromSeconds(30));
    }
    internal Task FlushAsync(Func<ScriptModuleSet> current, Func<ScriptModuleSet, Task> commit)
    {
        if (_closed) throw new ObjectDisposedException(nameof(ScriptEditorPage));
        return _flush is { IsCompleted: false } ? _flush : _flush = FlushCoreAsync(current, commit);
    }
    private async Task FlushCoreAsync(Func<ScriptModuleSet> current, Func<ScriptModuleSet, Task> commit)
    {
        try
        {
            if (Drafts.Failed) throw new InvalidOperationException("脚本页面已失败；请恢复已确认草稿后再保存或运行。");
            var response = await Browser.RequestAsync("flushDraft", "draftAck", new { });
            var flushed = response.Payload.Deserialize<Flushed>(WebMessage.Json) ?? throw new InvalidDataException("缺少草稿确认。");
            if (flushed.Busy is { } busy) throw new InvalidOperationException(busy);
            _locked = true;
            Drafts.ConfirmFlush(flushed.Versions ?? throw new InvalidDataException("缺少模块版本。"));
            var next = Drafts.BuildModules(current());
            await commit(next); Drafts.MarkCommitted(current());
        }
        finally
        {
            try { if (_holds == 0) await ReleaseAsync(); }
            finally { _flush = null; }
        }
    }
    internal Task SelectAsync(string module) => Browser.RequestAsync("select", "selected", new { module });
    internal Task LocateAsync(string module, int line, int column) => Browser.RequestAsync("locate", "located", new { module, line, column, version = Drafts.Models[module].Version });
    public async ValueTask DisposeAsync() { _closed = true; await Browser.DisposeAsync(); }
}
