using System.Text;
using System.Text.Json;
using Avalonia.Controls;
using Avalonia.Platform;
using Avalonia.Threading;
using QuickExcel.Infrastructure.Web;
using QuickExcel.Presentation.Web;

namespace QuickExcel.Desktop.Avalonia.Web;

/// <summary>One local page, one capability set and a bounded asynchronous request/reply channel.</summary>
internal sealed class LocalWebView : IAsyncDisposable
{
    private readonly string _document;
    private readonly long _generation;
    private readonly HashSet<string> _capabilities;
    private readonly Dictionary<string, Pending> _pending = [];
    private sealed record Pending(string Expected, TaskCompletionSource<WebMessage> Completion);
    private readonly TaskCompletionSource _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly TaskCompletionSource _destroyed = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly SemaphoreSlim _send = new(1);
    private readonly DispatcherTimer _heartbeat = new() { Interval = TimeSpan.FromSeconds(2) };
    private bool _closing, _failed, _pinging, _hasLoaded;
    private Task? _closeTask;
    internal NativeWebView View { get; } = new();
    internal LocalWebOrigin Origin { get; }
    internal bool IsReady => _ready.Task.IsCompletedSuccessfully && !_failed && !_closing;
    internal event Action<WebMessage>? Message;
    internal event Action<Exception>? Failed;
    internal Task Ready => _ready.Task;

    internal LocalWebView(WebAssetCatalog assets, string page, string document, long generation, string profileRoot, IEnumerable<string> capabilities)
    {
        var browserDirectory = WebRuntimePrerequisite.Require();
        _document = document; _generation = generation;
        _capabilities = new(capabilities, StringComparer.Ordinal); _capabilities.UnionWith(["ready", "pong", "disposed", "error"]);
        // A stable application-owned cache with in-private browser sessions. No prefix sweep or recursive deletion.
        var profile = Path.GetFullPath(Path.Combine(profileRoot, "browser")); Directory.CreateDirectory(profile);
        if ((File.GetAttributes(profile) & FileAttributes.ReparsePoint) != 0) throw new IOException("浏览器配置目录不能是重解析点。");
        Origin = new(assets, page, document, generation);
        if (OperatingSystem.IsWindows()) View.AdapterPreference = [WebViewAdapterType.WebView2];
        View.EnvironmentRequested += (_, e) =>
        {
            e.EnableDevTools = false;
            if (e is WindowsWebView2EnvironmentRequestedEventArgs windows)
            { windows.BrowserExecutableFolder = browserDirectory; windows.UserDataFolder = profile; windows.IsInPrivateModeEnabled = true; windows.AllowSingleSignOnUsingOSPrimaryAccount = false; }
            if (e is AppleWKWebViewEnvironmentRequestedEventArgs apple) apple.NonPersistentDataStore = true;
            if (e is LinuxWpeWebViewEnvironmentRequestedEventArgs linux)
            { linux.DataDirectory = profile; linux.CacheDirectory = Path.Combine(profile, "cache"); }
        };
        View.NavigationStarted += (_, e) => { e.Cancel = e.Request != Origin.Page; if (!e.Cancel && _hasLoaded && !_closing) Fail(new IOException("页面已重新加载，需恢复已确认的编辑状态。")); };
        View.NewWindowRequested += (_, e) => e.Handled = true;
        View.NavigationCompleted += (_, e) => { if (!e.IsSuccess) Fail(new IOException("本地页面加载失败，请检查 WebView Runtime。")); else _hasLoaded = true; };
        View.AdapterDestroyed += (_, _) => { _destroyed.TrySetResult(); if (!_closing) Fail(new IOException("浏览器页面已退出。")); };
        View.WebMessageReceived += Receive;
        View.Source = Origin.Page;
        _heartbeat.Tick += async (_, _) =>
        {
            if (!IsReady || _pinging) return; _pinging = true;
            try { await RequestAsync("ping", "pong", new { }, TimeSpan.FromSeconds(5)); }
            catch (Exception ex) { if (!_closing) Fail(ex); }
            finally { _pinging = false; }
        };
        _heartbeat.Start(); _ = CheckStartupAsync();
    }
    private async Task CheckStartupAsync()
    { try { await _ready.Task.WaitAsync(TimeSpan.FromSeconds(30)); } catch (Exception ex) { if (!_closing) Fail(ex); } }
    private void Receive(object? sender, WebMessageReceivedEventArgs e)
    {
        if (_failed || _closing && _destroyed.Task.IsCompleted) return;
        try
        {
            if (View.Source != Origin.Page || e.Body is not { Length: > 0 } body) throw new InvalidDataException("非本地页面消息。");
            var message = WebMessage.Parse(body, Origin.Instance, _document, _generation, _capabilities);
            if (message.Type == "error") throw new InvalidDataException("页面失败：" + message.Payload.ToString());
            if (message.Type == "ready")
            { if (!_ready.TrySetResult()) throw new InvalidDataException("重复的页面 ready。"); return; }
            if (!_ready.Task.IsCompletedSuccessfully) throw new InvalidDataException("页面尚未 ready。");
            if (_pending.TryGetValue(message.RequestId, out var pending))
            {
                if (pending.Expected != message.Type) throw new InvalidDataException("响应类型与待处理请求不匹配。");
                _pending.Remove(message.RequestId); pending.Completion.TrySetResult(message); return;
            }
            if (message.Type is "pong" or "disposed" or "draftAck" or "loaded") throw new InvalidDataException("未知、过期或重放的响应。");
            Message?.Invoke(message);
        }
        catch (Exception ex) { Fail(ex); }
    }
    private void Fail(Exception error)
    {
        if (_failed) return; _failed = true; _heartbeat.Stop(); _ready.TrySetException(error);
        foreach (var pending in _pending.Values) pending.Completion.TrySetException(error); _pending.Clear(); Failed?.Invoke(error);
    }
    internal async Task SendAsync(string type, object payload, string? requestId = null)
    {
        if (_failed || _closing && type != "dispose") throw new InvalidOperationException("Web 页面不可用。");
        await _ready.Task;
        var body = WebMessage.Create(Origin.Instance, _document, _generation, type, requestId ?? Guid.NewGuid().ToString("N"), payload);
        // The script is constant except for Base64 ASCII data. User source is never interpolated as executable JavaScript.
        var data = Convert.ToBase64String(Encoding.UTF8.GetBytes(body));
        await _send.WaitAsync();
        try { await View.InvokeScript("window.quickExcelReceiveBase64('" + data + "'); true").WaitAsync(TimeSpan.FromSeconds(10)); }
        finally { _send.Release(); }
    }
    internal async Task<WebMessage> RequestAsync(string type, string expected, object payload, TimeSpan? timeout = null)
    {
        if (_pending.Count >= 32) throw new InvalidOperationException("Web 待确认请求过多。");
        var id = Guid.NewGuid().ToString("N"); var completion = new TaskCompletionSource<WebMessage>(TaskCreationOptions.RunContinuationsAsynchronously);
        _pending.Add(id, new(expected, completion));
        try { await SendAsync(type, payload, id); return await completion.Task.WaitAsync(timeout ?? TimeSpan.FromSeconds(10)); }
        finally { _pending.Remove(id); }
    }
    internal Task CloseAsync(Action detach) => _closeTask ??= CloseCoreAsync(detach);
    private async Task CloseCoreAsync(Action detach)
    {
        _heartbeat.Stop();
        try { if (IsReady) await RequestAsync("dispose", "disposed", new { }, TimeSpan.FromSeconds(5)); }
        catch (Exception ex) { Fail(ex); }
        finally
        {
            _closing = true; _ready.TrySetCanceled(); View.WebMessageReceived -= Receive;
            var attached = View.AdapterInfo is not null; detach();
            try
            {
                if (attached) await _destroyed.Task.WaitAsync(TimeSpan.FromSeconds(5));
                if (OperatingSystem.IsLinux()) await global::Avalonia.X11.Interop.GtkInteropHelper.RunOnGlibThread(() => true);
            }
            catch (TimeoutException) { /* Still revoke all local resources even if the adapter cannot acknowledge destruction. */ }
            await Origin.DisposeAsync();
            foreach (var pending in _pending.Values) pending.Completion.TrySetException(new ObjectDisposedException(nameof(LocalWebView))); _pending.Clear();
        }
    }
    public async ValueTask DisposeAsync()
    {
        if (_closeTask is not null) await _closeTask;
        else await CloseAsync(() => { if (View.Parent is Panel panel) panel.Children.Remove(View); else if (View.Parent is ContentControl content) content.Content = null; else if (View.Parent is Decorator decorator) decorator.Child = null; });
    }
}
