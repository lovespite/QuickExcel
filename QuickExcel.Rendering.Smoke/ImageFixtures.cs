using BitMiracle.LibTiff.Classic;
using SkiaSharp;
namespace QuickExcel.Rendering.Smoke;
public static class ImageFixtures
{
    public static byte[] Bytes(string extension, bool alpha = false)
    {
        if (extension == "gif")
            return [.. "GIF89a"u8, 2, 0, 1, 0, 0x80, 0, 0, 255, 0, 0, 0, 0, 255, 0x2c, 0, 0, 0, 0, 2, 0, 1, 0, 0, 2, 2, 4, 10, 0, 0x3b];
        if (extension == "bmp")
        {
            using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
            writer.Write((ushort)0x4d42); writer.Write(62); writer.Write(0); writer.Write(54); writer.Write(40);
            writer.Write(2); writer.Write(1); writer.Write((ushort)1); writer.Write((ushort)24); writer.Write(0); writer.Write(8);
            writer.Write(3780); writer.Write(3780); writer.Write(0); writer.Write(0); writer.Write(new byte[] { 0, 0, 255, 0, 0, 255, 0, 0 });
            return stream.ToArray();
        }
        if (extension == "tiff")
        {
            using var stream = new MemoryStream();
            using (var tiff = Tiff.ClientOpen("fixture", "w", stream, new TiffStream()))
            {
                for (var page = 0; page < 2; page++)
                {
                    tiff.SetField(TiffTag.IMAGEWIDTH, 2); tiff.SetField(TiffTag.IMAGELENGTH, 1);
                    tiff.SetField(TiffTag.SAMPLESPERPIXEL, 4); tiff.SetField(TiffTag.BITSPERSAMPLE, 8);
                    tiff.SetField(TiffTag.ROWSPERSTRIP, 1); tiff.SetField(TiffTag.ORIENTATION, Orientation.TOPLEFT);
                    tiff.SetField(TiffTag.PLANARCONFIG, PlanarConfig.CONTIG); tiff.SetField(TiffTag.PHOTOMETRIC, Photometric.RGB);
                    tiff.SetField(TiffTag.COMPRESSION, Compression.LZW); tiff.SetField(TiffTag.EXTRASAMPLES, 1, new short[] { (short)ExtraSample.UNASSALPHA });
                    byte red = page == 0 ? (byte)255 : (byte)0; byte blue = page == 0 ? (byte)0 : (byte)255; byte a = alpha ? (byte)128 : (byte)255;
                    if (!tiff.WriteScanline([red, 0, blue, a, red, 0, blue, a], 0)) throw new InvalidDataException("TIFF fixture write failed."); tiff.WriteDirectory();
                }
            }
            return stream.ToArray();
        }
        using var bitmap = new SKBitmap(2, 1); bitmap.Erase(new SKColor(255, 0, 0, alpha ? (byte)128 : (byte)255));
        using var data = bitmap.Encode(extension == "jpg" ? SKEncodedImageFormat.Jpeg : SKEncodedImageFormat.Png, 100); return data.ToArray();
    }
}

