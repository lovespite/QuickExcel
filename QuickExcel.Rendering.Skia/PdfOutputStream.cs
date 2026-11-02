using System.Runtime.ExceptionServices;
using SkiaSharp;

namespace QuickExcel.Rendering.Skia;

/// <summary>Keeps managed I/O failures inside the native callback boundary and bounds every write.</summary>
internal sealed class PdfOutputStream(Stream stream, long maximumBytes) : SKManagedWStream(stream)
{
    private ExceptionDispatchInfo? _failure;
    private long _written;

    protected override bool OnWrite(IntPtr buffer, IntPtr size)
    {
        if (_failure is not null) return false;
        try
        {
            var count = size.ToInt64();
            if (count < 0 || count > maximumBytes - _written) throw new IOException("PDF 超过输出字节预算。");
            if (!base.OnWrite(buffer, size)) throw new IOException("PDF 写入失败。");
            _written += count;
            return true;
        }
        catch (Exception ex)
        {
            // Exceptions must never unwind through Skia's unmanaged stack. The
            // export checks this captured failure after each native document call.
            _failure = ExceptionDispatchInfo.Capture(ex);
            return false;
        }
    }

    protected override void OnFlush()
    {
        if (_failure is not null) return;
        try { base.OnFlush(); }
        catch (Exception ex) { _failure = ExceptionDispatchInfo.Capture(ex); }
    }

    protected override IntPtr OnBytesWritten() => (IntPtr)_written;
    internal void ThrowIfFailed() => _failure?.Throw();
}
