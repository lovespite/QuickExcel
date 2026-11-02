using Avalonia.Controls;
using Avalonia.Input;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class AvaloniaClipboard(TopLevel owner) : IClipboardExchange
{
    private static readonly DataFormat<string> Token = DataFormat.CreateStringApplicationFormat("QuickExcel.WorkbookClipboard.Token.v1");
    public async Task<ClipboardExchange> ReadAsync()
    {
        var clipboard = owner.Clipboard ?? throw new NotSupportedException("此环境不提供系统剪贴板。");
        using var data = await clipboard.TryGetDataAsync();
        if (data is null) return new(null);
        return new(await data.TryGetTextAsync(), await data.TryGetValueAsync(Token));
    }
    public async Task WriteAsync(ClipboardExchange value)
    {
        var clipboard = owner.Clipboard ?? throw new NotSupportedException("此环境不提供系统剪贴板。");
        var item = new DataTransferItem(); item.SetText(value.Text); item.Set(Token, value.Token);
        var data = new DataTransfer(); data.Add(item);
        // SetDataAsync takes ownership and disposes when the clipboard no longer needs it.
        await clipboard.SetDataAsync(data);
        await clipboard.FlushAsync();
    }
}
