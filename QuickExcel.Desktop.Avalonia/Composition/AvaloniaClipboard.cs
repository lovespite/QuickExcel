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
        var clipboard = owner.Clipboard ?? throw new NotSupportedException("此环境不提供系统剪贴板。");
        using var data = await clipboard.TryGetDataAsync();
        if (data is null) return new(null);
        var text = await data.TryGetTextAsync(); var token = await data.TryGetValueAsync(Token);
        var xml = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Xml) ?? await data.TryGetValueAsync(PortableXml) : null;
        var html = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Html) : null;
        var biff8 = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Biff8) : null;
        var chartPackage = await data.TryGetValueAsync(PortableDrawing);
        var drawing = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(OfficeDrawing) : null;
        var png = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Png) : null;
        if (png is null && xml is null && html is null)
        {
            using var bitmap = await data.TryGetBitmapAsync();
            if (bitmap is not null)
            {
                if ((long)bitmap.PixelSize.Width * bitmap.PixelSize.Height > 16_777_216) throw new InvalidDataException("剪贴板图片超过 1600 万像素。");
                using var output = new MemoryStream(); bitmap.Save(output, PngBitmapEncoderOptions.Default); png = output.ToArray();
            }
        }
        if (text?.Length > 4 * 1024 * 1024 || xml?.Length > SpreadsheetClipboardCodec.MaxBytes || html?.Length > SpreadsheetClipboardCodec.MaxBytes || biff8?.Length > SpreadsheetClipboardCodec.MaxBytes)
            throw new InvalidDataException("剪贴板载荷超过上限。");
        if (png?.Length > OfficeDrawingClipboardCodec.MaxBytes || drawing?.Length > OfficeDrawingClipboardCodec.MaxBytes || chartPackage?.Length > OfficeDrawingClipboardCodec.MaxBytes) throw new InvalidDataException("图片或图表剪贴板超过 16 MiB。");
        return new(text, token, xml, html, biff8, png, drawing, chartPackage);
    }
    public async Task WriteAsync(ClipboardExchange value)
    {
        var clipboard = owner.Clipboard ?? throw new NotSupportedException("此环境不提供系统剪贴板。");
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
