using Avalonia.Controls;
using Avalonia.Input;
using QuickExcel.Presentation;
using QuickExcel.Clipboard;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class AvaloniaClipboard(TopLevel owner) : IClipboardExchange
{
    private static readonly DataFormat<string> Token = DataFormat.CreateStringApplicationFormat("QuickExcel.WorkbookClipboard.Token.v1");
    private static readonly DataFormat<byte[]> Xml = DataFormat.CreateBytesPlatformFormat("XML Spreadsheet");
    private static readonly DataFormat<byte[]> PortableXml = DataFormat.CreateBytesApplicationFormat("QuickExcel.ExternalRegion.Xml.v1");
    private static readonly DataFormat<byte[]> Html = DataFormat.CreateBytesPlatformFormat("HTML Format");
    private static readonly DataFormat<byte[]> Biff8 = DataFormat.CreateBytesPlatformFormat("Biff8");
    public async Task<ClipboardExchange> ReadAsync()
    {
        var clipboard = owner.Clipboard ?? throw new NotSupportedException("此环境不提供系统剪贴板。");
        using var data = await clipboard.TryGetDataAsync();
        if (data is null) return new(null);
        var text = await data.TryGetTextAsync(); var token = await data.TryGetValueAsync(Token);
        var xml = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Xml) ?? await data.TryGetValueAsync(PortableXml) : null;
        var html = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Html) : null;
        var biff8 = OperatingSystem.IsWindows() ? await data.TryGetValueAsync(Biff8) : null;
        if (text?.Length > 4 * 1024 * 1024 || xml?.Length > SpreadsheetClipboardCodec.MaxBytes || html?.Length > SpreadsheetClipboardCodec.MaxBytes || biff8?.Length > SpreadsheetClipboardCodec.MaxBytes)
            throw new InvalidDataException("剪贴板载荷超过上限。");
        return new(text, token, xml, html, biff8);
    }
    public async Task WriteAsync(ClipboardExchange value)
    {
        var clipboard = owner.Clipboard ?? throw new NotSupportedException("此环境不提供系统剪贴板。");
        var item = new DataTransferItem(); item.SetText(value.Text); item.Set(Token, value.Token);
        if (OperatingSystem.IsWindows())
        {
            // Excel's XML importer resets the destination workbook's date system. Publish portable XML
            // privately for other QuickExcel instances; Excel consumes the independently verified HTML.
            if (value.SpreadsheetXml is { } xml) item.Set(PortableXml, xml);
            if (value.Html is { } html) item.Set(Html, html);
        }
        var data = new DataTransfer(); data.Add(item);
        // SetDataAsync takes ownership and disposes when the clipboard no longer needs it.
        await clipboard.SetDataAsync(data);
        await clipboard.FlushAsync();
    }
}
