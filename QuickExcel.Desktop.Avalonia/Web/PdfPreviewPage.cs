using QuickExcel.Infrastructure.Web;
using QuickExcel.Rendering.Skia;

namespace QuickExcel.Desktop.Avalonia.Web;

internal sealed class PdfPreviewPage : IAsyncDisposable
{
    internal LocalWebView Browser { get; }
    private LocalWebOrigin.PdfLease? _resource;
    private IAsyncDisposable? _file;
    private Task? _dispose;
    internal event Action? ExportRequested, PrintRequested, FocusRequested;
    internal event Action<Exception>? Failed;
    internal PdfPreviewPage(WebAssetCatalog assets, string document, long generation, string profileRoot)
    {
        string[] capabilities = ["pdfLoaded", "exportPdf", "print", "focusHost"];
        Browser = new(assets, "preview.html", document, generation, profileRoot, Program.WebTestMode ? capabilities.Append("testResult") : capabilities);
        Browser.Failed += ex => Failed?.Invoke(ex);
        Browser.Message += message =>
        {
            switch (message.Type)
            {
                case "exportPdf": ExportRequested?.Invoke(); break;
                case "print": PrintRequested?.Invoke(); break;
                case "focusHost": FocusRequested?.Invoke(); break;
                default: Failed?.Invoke(new InvalidDataException(UiText.Get("预览页发送了未请求的响应。"))); break;
            }
        };
    }
    internal async Task LoadAsync(PrintSession session)
    {
        if (_file is not null || _dispose is not null) throw new InvalidOperationException(UiText.Get("预览页已绑定打印任务或已关闭。"));
        _file = session.RetainPdf();
        _resource = await Browser.Origin.AddPdfAsync(session.PdfPath);
        await Browser.Ready;
        var response = await Browser.RequestAsync("loadPdf", "pdfLoaded", new Dictionary<string, object?> { ["url"] = _resource.Uri.AbsoluteUri }, TimeSpan.FromSeconds(45));
        if (response.Payload.GetProperty("pages").GetInt32() != session.Job.Layout.Pages.Length) throw new InvalidDataException(UiText.Get("浏览器预览页数与冻结打印任务不一致。"));
    }
    public ValueTask DisposeAsync() => new(_dispose ??= DisposeCoreAsync());
    private async Task DisposeCoreAsync()
    {
        try { await Browser.DisposeAsync(); }
        finally
        {
            try { if (_resource is not null) await _resource.DisposeAsync(); }
            finally { if (_file is not null) await _file.DisposeAsync(); }
        }
    }
}
