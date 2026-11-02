namespace QuickExcel.Presentation;

/// <summary>The OS carries bounded text and an opaque token; typed cells stay inside this process.</summary>
public sealed record ClipboardExchange(string? Text, string? Token = null);
public interface IClipboardExchange
{
    Task<ClipboardExchange> ReadAsync();
    Task WriteAsync(ClipboardExchange value);
}
