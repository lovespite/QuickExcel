using System.Text;
using QuickExcel.Core;

namespace QuickExcel.Storage;

public enum TextFileEncoding { Utf8, Utf16LittleEndian, Utf16BigEndian, Gb18030 }

public static class TextImportReader
{
    public static string[][] Read(byte[] bytes, char delimiter, TextFileEncoding encoding, CancellationToken cancellation = default)
    {
        if (bytes.Length > 64 * 1024 * 1024) throw new InvalidDataException(MessageText.Get("文本文件超过 64 MiB，请拆分后导入。"));
        if (delimiter is not (',' or '\t' or ';' or '|')) throw new ArgumentException(nameof(delimiter));
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding decoder = encoding switch
        {
            TextFileEncoding.Utf8 => new UTF8Encoding(false, true),
            TextFileEncoding.Utf16LittleEndian => new UnicodeEncoding(false, false, true),
            TextFileEncoding.Utf16BigEndian => new UnicodeEncoding(true, false, true),
            TextFileEncoding.Gb18030 => Encoding.GetEncoding(54936, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback),
            _ => throw new ArgumentException(nameof(encoding))
        };
        var text = decoder.GetString(bytes);
        if (text.StartsWith('\uFEFF')) text = text[1..];
        if (text.Length > 32 * 1024 * 1024) throw new InvalidDataException(MessageText.Get("文本交换超过 32 MiB 字符预算，请缩小范围。"));
        var rows = new List<string[]>();
        foreach (var row in DelimitedText.ReadRows(text, delimiter, WorkbookTextExchange.MaximumFields))
        { cancellation.ThrowIfCancellationRequested(); rows.Add(row); }
        return rows.ToArray();
    }

    public static byte[] ReadFile(string path, CancellationToken cancellation = default) => PackageIO.ReadFile(path, new(cancellation));
}
