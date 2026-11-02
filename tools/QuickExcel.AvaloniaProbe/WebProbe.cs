using System.Text.Json;
using Avalonia.Controls;
using Avalonia.Platform;

namespace QuickExcel.AvaloniaProbe;

internal sealed class WebProbe : IAsyncDisposable
{
    private readonly ProbeResources _resources;
    private readonly TaskCompletionSource _completed = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly TaskCompletionSource _released = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly TaskCompletionSource _destroyed = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly string _requestId = Guid.NewGuid().ToString("N");
    private bool _ready;
    private bool _closing;
    private readonly Uri _page;
    private WindowsVirtualOrigin? _virtualOrigin;
    internal NativeWebView View { get; }

    internal WebProbe(string assets)
    {
        _resources = new ProbeResources(assets, Path.Combine(Program.Output, "shared.pdf"));
        if (Program.VirtualHttps && !OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("The virtual HTTPS experiment requires WebView2.");
        _page = Program.VirtualHttps ? new Uri($"https://a0-{_resources.Instance}.invalid/{_resources.Instance}/probe.html") : _resources.Page;
        View = new NativeWebView();
        if (!Program.VirtualHttps) View.Source = _page;
        else View.AdapterCreated += (_, e) =>
        {
            try
            {
                if (!OperatingSystem.IsWindows() || e.TryGetPlatformHandle() is not IWindowsWebView2PlatformHandle windows)
                    throw new PlatformNotSupportedException("WebView2 COM handle unavailable.");
                _virtualOrigin = new WindowsVirtualOrigin(windows.CoreWebView2, _page.Host, _resources.WriteVirtualDirectory());
                View.Source = _page;
            }
            catch (Exception ex) { _completed.TrySetException(ex); }
        };
        if (OperatingSystem.IsWindows()) View.AdapterPreference = [WebViewAdapterType.WebView2];
        View.EnvironmentRequested += (_, e) =>
        {
            e.EnableDevTools = false;
            if (e is WindowsWebView2EnvironmentRequestedEventArgs windows)
            {
                windows.UserDataFolder = Path.Combine(Program.Output, "webview-profile");
                windows.IsInPrivateModeEnabled = true;
                windows.AllowSingleSignOnUsingOSPrimaryAccount = false;
            }
            if (e is AppleWKWebViewEnvironmentRequestedEventArgs apple) apple.NonPersistentDataStore = true;
            if (e is LinuxWpeWebViewEnvironmentRequestedEventArgs linux)
            {
                linux.DataDirectory = Path.Combine(Program.Output, "webview-profile");
                linux.CacheDirectory = Path.Combine(Program.Output, "webview-cache");
            }
        };
        View.NavigationStarted += (_, e) => e.Cancel = e.Request != _page;
        View.NewWindowRequested += (_, e) => e.Handled = true;
        View.NavigationCompleted += (_, e) =>
        {
            if (!e.IsSuccess) _completed.TrySetException(new IOException("WebView page load failed."));
        };
        View.WebMessageReceived += MessageReceived;
        View.AdapterDestroyed += (_, _) => _destroyed.TrySetResult();
    }

    private async void MessageReceived(object? sender, WebMessageReceivedEventArgs e)
    {
        try
        {
            if (e.Body is not { Length: > 0 and <= 32768 } body || View.Source != _page)
                throw new InvalidDataException("Unexpected message source or length.");
            using var json = JsonDocument.Parse(body, new() { MaxDepth = 12 });
            var message = json.RootElement;
            if (message.GetProperty("protocol").GetInt32() != 1 || message.GetProperty("instance").GetString() != _resources.Instance)
                throw new InvalidDataException("Unexpected message identity.");
            var type = message.GetProperty("type").GetString();
            if (type == "error") throw new InvalidDataException(message.GetProperty("payload").ToString());
            if (type == "disposed" && _closing && message.GetProperty("payload").GetProperty("requestId").GetString() == _requestId)
            {
                File.WriteAllText(Path.Combine(Program.Output, "web-disposed.json"), body);
                _released.TrySetResult();
            }
            else if (type == "ready" && !_ready)
            {
                _ready = true;
                File.WriteAllText(Path.Combine(Program.Output, "web.json"), body);
                var request = JsonSerializer.Serialize(new { protocol = 1, instance = _resources.Instance, type = "ping", requestId = _requestId });
                await View.InvokeScript($"window.quickExcelProbe({request}); true");
            }
            else if (type == "pong" && _ready && message.GetProperty("payload").GetProperty("requestId").GetString() == _requestId)
                _completed.TrySetResult();
            else throw new InvalidDataException("Unexpected probe message or replay.");
        }
        catch (Exception ex) { _completed.TrySetException(ex); _released.TrySetException(ex); }
    }

    internal async Task CloseAsync(Action detach)
    {
        _closing = true;
        var request = JsonSerializer.Serialize(new { protocol = 1, instance = _resources.Instance, type = "dispose", requestId = _requestId });
        await View.InvokeScript($"window.quickExcelProbe({request}); true");
        await _released.Task.WaitAsync(TimeSpan.FromSeconds(5));
        if (OperatingSystem.IsWindows()) { var origin = _virtualOrigin; _virtualOrigin = null; origin?.Dispose(); }
        detach();
        await _destroyed.Task.WaitAsync(TimeSpan.FromSeconds(5));
        // AdapterDestroyed precedes the GTK adapter's queued native disposal.
        if (OperatingSystem.IsLinux())
            await Avalonia.X11.Interop.GtkInteropHelper.RunOnGlibThread(() => true);
    }

    internal async Task VerifyAsync()
    {
        await _resources.CheckBoundaryAsync();
        await _completed.Task.WaitAsync(TimeSpan.FromSeconds(45));
        File.WriteAllText(Path.Combine(Program.Output, "webview-adapter.json"), JsonSerializer.Serialize(View.AdapterInfo));
        Program.Checks.Add("Offline PDF.js canvas/text, real PDF/Monaco workers, bidirectional message and resource boundary");
    }

    public async ValueTask DisposeAsync()
    {
        if (OperatingSystem.IsWindows()) { var origin = _virtualOrigin; _virtualOrigin = null; origin?.Dispose(); }
        View.WebMessageReceived -= MessageReceived;
        _completed.TrySetCanceled();
        await _resources.DisposeAsync();
    }
}
