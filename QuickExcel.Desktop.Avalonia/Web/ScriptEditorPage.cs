using System.Text.Json;
using System.Text.RegularExpressions;
using QuickExcel.Domain;
using QuickExcel.Infrastructure.Web;
using QuickExcel.Presentation.Web;

namespace QuickExcel.Desktop.Avalonia.Web;

internal sealed class ScriptEditorPage : IAsyncDisposable
{
    internal LocalWebView Browser { get; }
    internal ScriptDraftSession Drafts { get; }
    internal event Action<Exception>? Failed;
    internal event Action? SaveRequested, FocusRequested, ConsoleRequested, ApplyRequested, DraftChanged;
    internal event Func<Navigation, Task<bool>>? NavigateRequested;
    internal sealed record Navigation(string? Module, string? Definition, int Line = 1, int Column = 1);
    internal static bool IsDefinition(string name) => name == "quickexcel.d.ts" || QuickExcel.ScriptContracts.BuiltInLibraryIdentity.IsDefinition(name) ||
        Regex.IsMatch(name, @"\Alib\.(?:es(?:5|6|20\d\d|next)(?:\.[a-z0-9]+)*|decorators(?:\.legacy)?)\.d\.ts\z", RegexOptions.CultureInvariant);
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
    { if (_locked && !_closed && Browser.IsReady) { await Browser.RequestAsync("releaseDraft", "released", new Dictionary<string, object?> {  }); _locked = false; } }
    internal sealed record Changed(string Module, long Version);
    internal sealed record Flushed(Dictionary<string, long>? Versions, string? Busy);
    internal ScriptEditorPage(WebAssetCatalog assets, ScriptDraftSession drafts, string document, long generation, string profileRoot)
    {
        Drafts = drafts;
        string[] capabilities = ["draftChanged", "draftChunk", "draftAck", "initialized", "loadAck", "loaded", "released", "selected", "located", "definitionSelected", "navigate", "save", "focusHost", "console", "insertionCaptured", "templateInserted", "commandExecuted"];
        capabilities = [.. capabilities, "apply"];
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
                    var changed = message.Payload.Deserialize(DesktopJsonContext.Web.Changed) ?? throw new InvalidDataException(UiText.Get("缺少草稿版本。"));
                    Drafts.Observe(changed.Module, changed.Version); DraftChanged?.Invoke(); break;
                case "draftChunk":
                    var chunk = message.Payload.Deserialize(WebJsonContext.Default.DraftChunk) ?? throw new InvalidDataException(UiText.Get("缺少草稿分块。"));
                    Drafts.Receive(chunk);
                    await Browser.SendAsync("chunkAck", new Dictionary<string, object?> { ["transfer"] = chunk.Transfer, ["index"] = chunk.Index, ["version"] = chunk.Version }, message.RequestId); break;
                case "save": SaveRequested?.Invoke(); break;
                case "focusHost": FocusRequested?.Invoke(); break;
                case "console": ConsoleRequested?.Invoke(); break;
                case "apply": ApplyRequested?.Invoke(); break;
                case "navigate":
                    var target = message.Payload.Deserialize(DesktopJsonContext.Web.Navigation) ?? throw new InvalidDataException(UiText.Get("缺少定义导航目标。"));
                    if ((target.Module is null) == (target.Definition is null) || target.Line < 1 || target.Column < 1 ||
                        target.Module is { } module && !Drafts.Models.ContainsKey(module) ||
                        target.Definition is { } definition && (definition.Length > 128 || !IsDefinition(definition)))
                        throw new InvalidDataException(UiText.Get("无效的定义导航目标。"));
                    var accepted = NavigateRequested is { } navigate && await navigate(target);
                    await Browser.SendAsync("navigated", new Dictionary<string, object?> { ["accepted"] = accepted }, message.RequestId); break;
                default: throw new InvalidDataException(UiText.Get("脚本页发送了未请求的响应。"));
            }
        }
        catch (Exception ex) { Drafts.MarkFailed(); Failed?.Invoke(ex); }
    }
    internal async Task LoadAsync(string module)
    {
        await Browser.Ready;
        await Browser.RequestAsync("initialize", "initialized", new Dictionary<string, object?> {
            ["modules"] = Drafts.Models.Keys.ToArray(),
            ["readOnly"] = Drafts.IsReadOnly,
            ["apiVersion"] = Drafts.ApiVersion,
            ["languages"] = Drafts.Models.ToDictionary(p => p.Key, p => p.Value.Language == ScriptLanguage.TypeScript ? "typescript" : "javascript")
        });
        foreach (var (id, draft) in Drafts.Models)
        foreach (var chunk in ScriptDraftSession.Split(id, draft.Text, 0, 1))
            await Browser.RequestAsync("loadChunk", "loadAck", chunk);
        await Browser.RequestAsync("activate", "loaded", new Dictionary<string, object?> { ["module"] = module }, TimeSpan.FromSeconds(30));
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
            if (Drafts.Failed) throw new InvalidOperationException(UiText.Get("脚本页面已失败；请恢复已确认草稿后再保存或运行。"));
            var response = await Browser.RequestAsync("flushDraft", "draftAck", new Dictionary<string, object?> {  });
            var flushed = response.Payload.Deserialize(DesktopJsonContext.Web.Flushed) ?? throw new InvalidDataException(UiText.Get("缺少草稿确认。"));
            if (flushed.Busy is { } busy) throw new InvalidOperationException(busy);
            _locked = true;
            Drafts.ConfirmFlush(flushed.Versions ?? throw new InvalidDataException(UiText.Get("缺少模块版本。")));
            var next = Drafts.BuildModules(current());
            await commit(next); Drafts.MarkCommitted(current());
        }
        finally
        {
            try { if (_holds == 0) await ReleaseAsync(); }
            finally { _flush = null; }
        }
    }
    internal Task SelectAsync(string module) => Browser.RequestAsync("select", "selected", new Dictionary<string, object?> { ["module"] = module });
    internal async Task<JsonElement> CaptureInsertionAsync(string module)
    {
        var reply = await Browser.RequestAsync("captureInsertion", "insertionCaptured", new Dictionary<string, object?> { ["module"] = module });
        if (!reply.Payload.GetProperty("accepted").GetBoolean()) throw new InvalidOperationException(UiText.Get("请先完成输入，并在可编辑的事件脚本中放置光标。"));
        return reply.Payload.GetProperty("anchor").Clone();
    }
    internal async Task InsertTemplateAsync(JsonElement anchor, string text)
    {
        var reply = await Browser.RequestAsync("insertTemplate", "templateInserted", new Dictionary<string, object?> { ["anchor"] = anchor, ["text"] = text });
        if (!reply.Payload.GetProperty("accepted").GetBoolean()) throw new InvalidOperationException(UiText.Get("脚本或光标所属模块已变化，请重新添加事件处理器。"));
    }
    internal Task ExecuteEditorCommandAsync(string command) => Browser.RequestAsync("editorCommand", "commandExecuted", new Dictionary<string, object?> { ["command"] = command });
    internal Task OpenDefinitionAsync(string name, int? line = null, int? column = null) => IsDefinition(name) ?
        Browser.RequestAsync("definition", "definitionSelected", new Dictionary<string, object?> { ["name"] = name, ["line"] = line, ["column"] = column }) : throw new InvalidDataException(UiText.Get("未知定义文件。"));
    internal Task LocateAsync(string module, int line, int column) => Browser.RequestAsync("locate", "located", new Dictionary<string, object?> { ["module"] = module, ["line"] = line, ["column"] = column, ["version"] = Drafts.Models[module].Version });
    public async ValueTask DisposeAsync() { _closed = true; await Browser.DisposeAsync(); }
}
