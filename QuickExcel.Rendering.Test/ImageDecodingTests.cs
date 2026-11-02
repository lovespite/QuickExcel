using BitMiracle.LibTiff.Classic;
using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Rendering.Test;

public sealed class ImageDecodingTests
{
    [Theory]
    [InlineData("png")]
    [InlineData("jpg")]
    [InlineData("gif")]
    [InlineData("bmp")]
    [InlineData("tiff")]
    public void ExistingFormatsDecodeFirstFrameAndKeepOriginalBytes(string extension)
    {
        var bytes = Bytes(extension); var image = new SheetImage(bytes, extension, 0, 0);
        using var renderer = new SkiaRenderer(); using var decoded = renderer.DecodeImage(image);
        Assert.Equal(2, decoded.Width); Assert.Equal(1, decoded.Height);
        var pixel = decoded.GetPixel(0, 0); Assert.True(pixel.Red > 180 && pixel.Green < 80 && pixel.Blue < 80, pixel.ToString());
        Assert.Equal(bytes, image.Data.ToArray());
    }
    [Fact]
    public void DecodeRejectsExcessPixelsBeforeAllocatingPixelBuffer()
    {
        using var renderer = new SkiaRenderer(budget: new() { MaxImagePixels = 1 });
        foreach (var extension in new[] { "png", "jpg", "gif", "bmp", "tiff" })
            Assert.Throws<NotSupportedException>(() => renderer.DecodeImage(new(Bytes(extension), extension, 0, 0)));
    }
    [Fact]
    public void PngAndTiffRetainAlphaAndMonochromeOnlyChangesProjection()
    {
        using var renderer = new SkiaRenderer();
        foreach (var extension in new[] { "png", "tiff" })
        {
            using var decoded = renderer.DecodeImage(new(Bytes(extension, alpha: true), extension, 0, 0));
            Assert.InRange(decoded.GetPixel(0, 0).Alpha, (byte)127, (byte)128);
            using var output = new SKBitmap(2, 1); using var canvas = new SKCanvas(output); canvas.Clear(SKColors.White);
            SkiaRenderer.DrawImage(canvas, decoded, SKRect.Create(2, 1), true);
            var pixel = output.GetPixel(0, 0); Assert.Equal(pixel.Red, pixel.Green); Assert.Equal(pixel.Red, pixel.Blue);
        }
    }
    private static byte[] Bytes(string extension, bool alpha = false) => QuickExcel.Rendering.Smoke.ImageFixtures.Bytes(extension, alpha);
}
