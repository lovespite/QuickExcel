namespace QuickExcel.Presentation;

/// <summary>One OS data object. The token only resolves a local capture; rich bytes are untrusted external data.</summary>
public sealed record ClipboardExchange(string? Text, string? Token = null, byte[]? SpreadsheetXml = null, byte[]? Html = null, byte[]? Biff8 = null);
public interface IClipboardExchange
{
    Task<ClipboardExchange> ReadAsync();
    Task WriteAsync(ClipboardExchange value);
}
