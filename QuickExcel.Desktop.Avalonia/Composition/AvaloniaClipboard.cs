using Avalonia.Controls;
using Avalonia.Input;
using QuickExcel.Presentation;
using QuickExcel.Clipboard;
using Avalonia.Media.Imaging;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class AvaloniaClipboard(TopLevel owner) : IClipboardExchange
{
    private static readonly DataFormat<string> Token = DataFormat.CreateStringApplicationFormat("QuickExcel.WorkbookClipboard.Token.v1");
    private static readonly DataFormat<byte[]> Xml = DataFormat.CreateBytesPlatformFormat("XML Spreadsheet");
    private static readonly DataFormat<byte[]> PortableXml = DataFormat.CreateBytesApplicationFormat("QuickExcel.ExternalRegion.Xml.v1");
    private static readonly DataFormat<byte[]> Html = DataFormat.CreateBytesPlatformFormat("HTML Format");
    private static readonly DataFormat<byte[]> Biff8 = DataFormat.CreateBytesPlatformFormat("Biff8");
    private static readonly DataFormat<byte[]> Png = DataFormat.CreateBytesPlatformFormat("PNG");
    private static readonly DataFormat<byte[]> OfficeDrawing = DataFormat.CreateBytesPlatformFormat("Art::GVML ClipFormat");
    private static readonly DataFormat<byte[]> PortableDrawing = DataFormat.CreateBytesApplicationFormat("QuickExcel.ChartClipboard.Xlsx.v1");
    public async Task<ClipboardExchange> ReadAsync()
    {
        var clipboard = owner.Clipboard ?? throw new NotSupportedException(UiText.Get("此环境不提供系统剪贴板。"));
        using var data = await clipboard.TryGetDataAsync();
        if (data is null) return new(null);
        var notices = new List<string>();
        var text = await data.TryGetTextAsync(); var token = await data.TryGetValueAsync(Token);
        var xml = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Xml) ?? await data.TryGetValueAsync(PortableXml) : null;
        var html = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Html) : null;
        var biff8 = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Biff8) : null;
        var chartPackage = await data.TryGetValueAsync(PortableDrawing);
        var drawing = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(OfficeDrawing) : null;
        var png = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Png) : null;
        if (png is null)
        {
            try
            {
                using var bitmap = await data.TryGetBitmapAsync();
                if (bitmap is not null)
                {
                    if ((long)bitmap.PixelSize.Width * bitmap.PixelSize.Height > 16_777_216) notices.Add(UiText.Get("已跳过超过 1600 万像素的图片。"));
                    else { using var output = new MemoryStream(); bitmap.Save(output, PngBitmapEncoderOptions.Default); png = output.ToArray(); }
                }
            }
            catch (Exception ex) when (ex is NotSupportedException or InvalidDataException or ArgumentException) { notices.Add(UiText.Get("位图不可读取，已尝试其他内容。")); }
        }
        byte[]? Bounded(byte[]? bytes, int limit, string name)
        {
            if (bytes?.Length <= limit || bytes is null) return bytes;
            notices.Add(UiText.Get("已跳过超过大小限制的") + name + "。"); return null;
        }
        if (text?.Length > GridClipboard.MaxCharacters) { notices.Add(UiText.Get("已跳过超过大小限制的文本。")); text = null; }
        xml = Bounded(xml, SpreadsheetClipboardCodec.MaxBytes, UiText.Get("表格格式")); html = Bounded(html, SpreadsheetClipboardCodec.MaxBytes, "HTML");
        biff8 = Bounded(biff8, SpreadsheetClipboardCodec.MaxBytes, "Biff8"); png = Bounded(png, OfficeDrawingClipboardCodec.MaxBytes, UiText.Get("图片"));
        drawing = Bounded(drawing, OfficeDrawingClipboardCodec.MaxBytes, UiText.Get("对象")); chartPackage = Bounded(chartPackage, OfficeDrawingClipboardCodec.MaxBytes, UiText.Get("图表"));
        return new(text, token, xml, html, biff8, png, drawing, chartPackage, notices.Count == 0 ? null : string.Join(" ", notices));
    }
    public async Task WriteAsync(ClipboardExchange value)
    {
        var clipboard = owner.Clipboard ?? throw new NotSupportedException(UiText.Get("此环境不提供系统剪贴板。"));
        var item = new DataTransferItem(); item.SetText(value.Text); item.Set(Token, value.Token);
        if (value.ChartPackage is { } drawing) item.Set(PortableDrawing, drawing);
        if (OperatingSystem.IsWindows() && value.OfficeDrawing is { } office) item.Set(OfficeDrawing, office);
        if (value.Png is { } png) item.SetBitmap(new Bitmap(new MemoryStream(png, writable: false)));
        if (OperatingSystem.IsWindows())
        {
            // Excel's XML importer resets the destination workbook's date system. Publish portable XML
            // privately for other QuickExcel instances; Excel consumes the independently verified HTML.
            if (value.SpreadsheetXml is { } xml) item.Set(PortableXml, xml);
            if (value.Html is { } html) item.Set(Html, html);
            if (value.Png is { } bytes) item.Set(Png, bytes);
        }
        var data = new DataTransfer(); data.Add(item);
        // SetDataAsync takes ownership and disposes when the clipboard no longer needs it.
        await clipboard.SetDataAsync(data);
        await clipboard.FlushAsync();
    }
}
