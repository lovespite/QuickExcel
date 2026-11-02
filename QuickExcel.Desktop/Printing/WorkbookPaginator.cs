using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using QuickExcel.Core;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;

namespace QuickExcel.Desktop;

internal sealed class ReportDocument(DocumentPaginator paginator) : IDocumentPaginatorSource
{ public DocumentPaginator DocumentPaginator => paginator; }

/// <summary>WPF presentation/printing adapter for the immutable, platform-independent report.</summary>
internal sealed class WorkbookPaginator : DocumentPaginator, IDisposable, IAsyncDisposable
{
    internal SkiaReport Report { get; }
    internal IReadOnlyList<string> Warnings => Report.Job.Warnings;
    internal PrintLayout Layout => Report.Job.Layout;
    private readonly CancellationTokenSource _lifetime = new();
    internal CancellationToken Lifetime => _lifetime.Token;
    private readonly SemaphoreSlim _renderGate = new(1);
    private readonly object _cacheGate = new();
    private readonly Dictionary<(int Page, int Dpi), (BitmapSource Image, long Stamp)> _cache = [];
    private long _cacheBytes, _stamp;
    private bool _disposed;
    private Task? _disposeTask;
    internal WorkbookPaginator(WorkbookSnapshot book, WorksheetSnapshot sheet, string file) =>
        Report = new(book, sheet, file, ConditionalPresentation.GetStyle);
    public override bool IsPageCountValid => true;
    public override int PageCount => Layout.Pages.Length;
    public override Size PageSize { get => new(Layout.PaperWidth, Layout.PaperHeight); set { if (value != PageSize) throw new NotSupportedException("请使用页面设置更改纸张，以保持预览和打印分页一致。"); } }
    public override IDocumentPaginatorSource Source => new ReportDocument(this);
    internal int CachedPageCount { get { lock (_cacheGate) return _cache.Count; } }
    internal long CachedPageBytes { get { lock (_cacheGate) return _cacheBytes; } }
    internal FixedDocument CreatePreviewDocument()
    {
        var document = new FixedDocument(); document.DocumentPaginator.PageSize = PageSize;
        for (var index = 0; index < PageCount; index++)
        {
            var page = new FixedPage { Width = PageSize.Width, Height = PageSize.Height };
            page.Children.Add(new PreviewPage(this, index) { Width = PageSize.Width, Height = PageSize.Height });
            document.Pages.Add(new PageContent { Child = page });
        }
        return document;
    }
    internal FrameworkElement CreateThumbnails(DocumentViewer viewer)
    {
        var list = new ListBox { Width = 156, BorderThickness = new(0), Background = Brushes.WhiteSmoke };
        VirtualizingPanel.SetIsVirtualizing(list, true); VirtualizingPanel.SetVirtualizationMode(list, VirtualizationMode.Recycling);
        for (var i = 0; i < PageCount; i++)
            list.Items.Add(new PreviewPage(this, i) { Width = 128, Height = 128 * PageSize.Height / PageSize.Width, Margin = new(4), ToolTip = $"第 {i + 1} 页" });
        list.SelectionChanged += (_, _) => { if (list.SelectedIndex >= 0) viewer.GoToPage(list.SelectedIndex + 1); };
        return list;
    }
    private BitmapSource Render(int page, int dpi, CancellationToken token)
    {
        using var bitmap = Report.RenderPage(page, dpi, token); token.ThrowIfCancellationRequested();
        return SkiaWpfBridge.Copy(bitmap, dpi, dpi);
    }
    private int BoundedDpi(double requested)
    {
        var maximum = Math.Floor(96 * Math.Sqrt(Report.Renderer.Budget.MaxPreviewBytes / (4d * Math.Ceiling(PageSize.Width) * Math.Ceiling(PageSize.Height))));
        return (int)Math.Clamp(Math.Ceiling(requested / 12) * 12, 12, Math.Min(600, Math.Max(12, maximum - 1)));
    }
    internal async Task<BitmapSource> GetPreviewAsync(int page, double requestedDpi, CancellationToken token)
    {
        var dpi = BoundedDpi(requestedDpi);
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(token, Lifetime);
        await _renderGate.WaitAsync(linked.Token);
        try
        {
            linked.Token.ThrowIfCancellationRequested();
            lock (_cacheGate)
                if (_cache.TryGetValue((page, dpi), out var hit)) { _cache[(page, dpi)] = (hit.Image, ++_stamp); return hit.Image; }
            var image = await Task.Run(() => Render(page, dpi, linked.Token), linked.Token);
            linked.Token.ThrowIfCancellationRequested();
            lock (_cacheGate)
            {
                var bytes = (long)image.PixelWidth * image.PixelHeight * 4;
                while (_cache.Count > 0 && (_cache.Count >= 8 || _cacheBytes + bytes > Report.Renderer.Budget.MaxPreviewBytes))
                {
                    var oldest = _cache.MinBy(p => p.Value.Stamp);
                    _cacheBytes -= (long)oldest.Value.Image.PixelWidth * oldest.Value.Image.PixelHeight * 4; _cache.Remove(oldest.Key);
                }
                _cache[(page, dpi)] = (image, ++_stamp); _cacheBytes += bytes;
            }
            return image;
        }
        finally { _renderGate.Release(); }
    }
    internal async Task WritePdfAsync(string destination, CancellationToken token)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(token, Lifetime);
        await _renderGate.WaitAsync(linked.Token);
        try { await Task.Run(() => Report.WritePdf(destination, linked.Token), linked.Token); }
        finally { _renderGate.Release(); }
    }
    internal void Paint(DrawingContext drawing, int pageNumber)
    {
        var image = Render(pageNumber, 96, Lifetime);
        drawing.DrawImage(image, new Rect(0, 0, image.Width, image.Height));
    }
    public override DocumentPage GetPage(int pageNumber) => GetPageAtDpi(pageNumber, 96);
    private DocumentPage GetPageAtDpi(int pageNumber, int dpi)
    {
        if (pageNumber < 0 || pageNumber >= PageCount) return DocumentPage.Missing;
        var image = Render(pageNumber, dpi, Lifetime);
        // Pixel rounding adds at most one trailing pixel. Do not rescale every glyph
        // back onto the fractional paper edge (a second interpolation pass blurs small text).
        var visual = new DrawingVisual(); using (var dc = visual.RenderOpen()) dc.DrawImage(image, new Rect(0, 0, image.Width, image.Height));
        return new(visual, PageSize, new Rect(PageSize), new Rect(PageSize));
    }
    internal DocumentPaginator CreatePrinterPaginator() => new PrinterPaginator(this);
    private sealed class PrinterPaginator(WorkbookPaginator owner) : DocumentPaginator
    {
        public override bool IsPageCountValid => true;
        public override int PageCount => owner.PageCount;
        public override Size PageSize { get => owner.PageSize; set => owner.PageSize = value; }
        public override IDocumentPaginatorSource Source => new ReportDocument(this);
        public override DocumentPage GetPage(int pageNumber) => owner.GetPageAtDpi(pageNumber, owner.BoundedDpi(300));
    }
    private sealed class PreviewPage : FrameworkElement
    {
        private readonly WorkbookPaginator _owner;
        private readonly int _page;
        private BitmapSource? _image;
        private CancellationTokenSource? _request;
        private int _requestedDpi, _generation;
        internal PreviewPage(WorkbookPaginator owner, int page)
        {
            _owner = owner; _page = page;
            Loaded += (_, _) => { LayoutUpdated += OnLayoutUpdated; InvalidateVisual(); };
            Unloaded += (_, _) => { LayoutUpdated -= OnLayoutUpdated; _generation++; _request?.Cancel(); _request?.Dispose(); _request = null; _image = null; _requestedDpi = 0; };
            SizeChanged += (_, _) => InvalidateVisual();
        }
        private void OnLayoutUpdated(object? sender, EventArgs e) { if (!_owner._disposed && IsVisible && Dpi() != _requestedDpi) InvalidateVisual(); }
        private int Dpi()
        {
            var scale = ActualWidth / _owner.PageSize.Width;
            if (Window.GetWindow(this) is { } window)
            {
                var transform = TransformToAncestor(window); var origin = transform.Transform(new Point()); var unit = transform.Transform(new Point(1, 0));
                scale *= (unit - origin).Length * VisualTreeHelper.GetDpi(this).DpiScaleX;
            }
            return _owner.BoundedDpi(96 * Math.Max(.1, scale));
        }
        protected override void OnRender(DrawingContext drawing)
        {
            if (_owner._disposed) return;
            // Detached diagnostic/printer visuals are synchronous, with the same shared page renderer.
            if (PresentationSource.FromVisual(this) is null) { _owner.Paint(drawing, _page); return; }
            drawing.DrawRectangle(Brushes.White, null, new Rect(RenderSize));
            if (_image is not null) drawing.DrawImage(_image, new Rect(RenderSize));
            var dpi = Dpi(); if (dpi != _requestedDpi) Request(dpi);
        }
        private async void Request(int dpi)
        {
            _request?.Cancel(); _request?.Dispose(); _request = new();
            var generation = ++_generation; _requestedDpi = dpi; var token = _request.Token;
            try
            {
                await Task.Delay(50, token);
                var image = await _owner.GetPreviewAsync(_page, dpi, token);
                if (generation != _generation || !IsVisible || _owner._disposed) return;
                _image = image; InvalidateVisual();
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { if (generation == _generation) ToolTip = "页面预览未完成：" + ex.Message; }
        }
    }
    public ValueTask DisposeAsync() => new(_disposeTask ??= DisposeCoreAsync());
    private async Task DisposeCoreAsync()
    {
        _disposed = true; _lifetime.Cancel();
        await _renderGate.WaitAsync();
        try { lock (_cacheGate) { _cache.Clear(); _cacheBytes = 0; } Report.Dispose(); }
        finally { _renderGate.Release(); _lifetime.Dispose(); }
    }
    public void Dispose() => _ = DisposeAsync();
}
