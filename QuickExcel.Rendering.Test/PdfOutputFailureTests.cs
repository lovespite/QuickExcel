using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Rendering.Test;

public sealed class PdfOutputFailureTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativePdfWriterReturnsIoFailuresToManagedCaller(bool failOnFlush)
    {
        using var file = new FailingStream(failOnFlush);
        using var output = new PdfOutputStream(file, 1024 * 1024);
        using var pdf = SKDocument.CreatePdf(output);
        using var paint = new SKPaint { Color = SKColors.Blue };
        pdf.BeginPage(200, 200).DrawRect(10, 10, 100, 100, paint);
        pdf.EndPage(); pdf.Close(); output.Flush();
        var error = Assert.Throws<IOException>(output.ThrowIfFailed);
        Assert.Equal("Injected storage failure", error.Message);
        Assert.True(file.Failed);
    }

    [Fact]
    public void OutputLimitStopsNativeWritesBeforeExceedingBudget()
    {
        using var file = new MemoryStream(); using var output = new PdfOutputStream(file, 20);
        using var pdf = SKDocument.CreatePdf(output);
        pdf.BeginPage(200, 200).Clear(SKColors.White); pdf.EndPage(); pdf.Close();
        Assert.Throws<IOException>(output.ThrowIfFailed);
        Assert.InRange(file.Length, 0, 20);
    }

    private sealed class FailingStream(bool failOnFlush) : MemoryStream
    {
        internal bool Failed { get; private set; }
        public override void Write(byte[] buffer, int offset, int count)
        {
            if (!failOnFlush) { Failed = true; throw new IOException("Injected storage failure"); }
            base.Write(buffer, offset, count);
        }
        public override void Flush()
        {
            if (failOnFlush) { Failed = true; throw new IOException("Injected storage failure"); }
            base.Flush();
        }
    }
}
