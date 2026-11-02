using BitMiracle.LibTiff.Classic;
using QuickExcel.Domain;
using SkiaSharp;

namespace QuickExcel.Rendering.Skia;

/// <summary>CPU/PDF content renderer. Each instance belongs to one render thread/task.</summary>
public sealed partial class SkiaRenderer : IDisposable
{
    public RenderBudget Budget { get; }
    public SkiaFonts Fonts { get; }
    private readonly List<RenderDiagnostic> _diagnostics = [];
    private readonly HashSet<RenderDiagnostic> _diagnosticSet = [];
    public IReadOnlyList<RenderDiagnostic> Diagnostics => _diagnostics;
    public SkiaRenderer(IEnumerable<string>? fontFiles = null, bool systemFonts = true, RenderBudget? budget = null)
    { Budget = budget ?? new(); Fonts = new(fontFiles, systemFonts, Budget.MaxFontFaces); }

    internal void Diagnostic(string code, string location, string message)
    {
        var value = new RenderDiagnostic(code, location, message);
        if (_diagnostics.Count < Budget.MaxDiagnostics && _diagnosticSet.Add(value)) _diagnostics.Add(value);
    }
    public SkiaTextLayout Layout(TextLayoutRequest request, string location = "") =>
        SkiaTextLayout.Create(request, Fonts, (code, message) => Diagnostic(code, location, message));
    public static SKColor Color(string? argb, SKColor fallback) => argb is { Length: 8 } && uint.TryParse(argb,
        System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out var color) ? new SKColor(color) : fallback;
    public static SKRect Rect(RenderRect r) => SKRect.Create((float)r.X, (float)r.Y, (float)r.Width, (float)r.Height);

    public void DrawCell(SKCanvas canvas, CellRenderInput input, RenderOutput output = RenderOutput.Print, bool monochrome = false, string location = "")
    {
        var rect = Rect(input.Bounds); var style = input.Text.Style;
        if (rect.Width <= 0 || rect.Height <= 0) return;
        using var paint = new SKPaint { Color = monochrome ? SKColors.White : Color(style.FillColor, SKColors.White), IsAntialias = true };
        canvas.DrawRect(rect, paint);
        if (input.GridLines)
        {
            paint.Color = SKColors.LightGray; paint.Style = SKPaintStyle.Stroke; paint.StrokeWidth = output == RenderOutput.Screen ? 1 : .5f;
            canvas.DrawRect(rect, paint); paint.Style = SKPaintStyle.Fill;
        }
        DrawText(canvas, input.Text, rect, monochrome, location);
        void Border(int kind, string? color, SKPoint from, SKPoint to)
        {
            if (kind == 0) return;
            paint.Color = monochrome ? SKColors.Black : Color(color, SKColors.Black);
            paint.StrokeWidth = kind is 2 or 5 ? 2 : 1;
            using var dash = kind is 3 or 4 ? SKPathEffect.CreateDash(kind == 3 ? [4, 3] : [1, 2], 0) : null;
            paint.PathEffect = dash; canvas.DrawLine(from, to, paint);
            if (kind == 6) { var delta = from.X == to.X ? new SKPoint(2, 0) : new SKPoint(0, 2); canvas.DrawLine(from + delta, to + delta, paint); }
            paint.PathEffect = null;
        }
        Border(style.LeftBorder, style.LeftBorderColor, new(rect.Left, rect.Top), new(rect.Left, rect.Bottom));
        Border(style.RightBorder, style.RightBorderColor, new(rect.Right, rect.Top), new(rect.Right, rect.Bottom));
        Border(style.TopBorder, style.TopBorderColor, new(rect.Left, rect.Top), new(rect.Right, rect.Top));
        Border(style.BottomBorder, style.BottomBorderColor, new(rect.Left, rect.Bottom), new(rect.Right, rect.Bottom));
        if (input.Note && output == RenderOutput.Screen)
        {
            paint.Color = SKColors.Red; using var marker = new SKPathBuilder();
            marker.MoveTo(rect.Right - 7, rect.Top + 1); marker.LineTo(rect.Right - 1, rect.Top + 1); marker.LineTo(rect.Right - 1, rect.Top + 7); marker.Close();
            using var path = marker.Detach(); canvas.DrawPath(path, paint);
        }
    }

    public void DrawText(SKCanvas canvas, TextLayoutRequest request, SKRect rect, bool monochrome = false, string location = "")
    {
        if (rect.Width <= 4 || rect.Height <= 0 || request.Text.Length == 0) return;
        using var layout = Layout(request with { Width = Math.Max(1, rect.Width - 4) }, location);
        var style = request.Style; var metrics = layout.Metrics;
        canvas.Save(); canvas.ClipRect(rect);
        var left = rect.Left + 2 + (style.Horizontal == 3 ? 0 : layout.Indent);
        if (style.TextRotation is > 0 and <= 180)
        {
            var angle = style.TextRotation <= 90 ? -style.TextRotation : style.TextRotation - 90;
            var matrix = SKMatrix.CreateRotationDegrees(angle);
            var bounds = matrix.MapRect(SKRect.Create((float)metrics.Width, (float)metrics.Height));
            var dx = style.Horizontal switch { 2 => (rect.Width - 4 - layout.Indent - bounds.Width) / 2, 3 => rect.Width - 4 - bounds.Width, _ => 0 };
            var dy = style.Vertical switch { 0 => 1, 1 => (rect.Height - bounds.Height) / 2, _ => rect.Height - bounds.Height - 1 };
            canvas.Translate(left + dx - bounds.Left, rect.Top + dy - bounds.Top); canvas.RotateDegrees(angle); layout.Draw(canvas, 0, 0, monochrome);
        }
        else
        {
            var y = style.Vertical switch { 0 => rect.Top + 1, 1 => rect.Top + (rect.Height - metrics.Height) / 2, _ => rect.Bottom - metrics.Height - 1 };
            layout.Draw(canvas, left, (float)y, monochrome);
        }
        canvas.Restore();
    }

    /// <summary>First frame only. TIFF uses an independent managed decoder; no WIC/GDI dependency.</summary>
    public SKBitmap DecodeImage(SheetImage image)
    {
        void Check(int width, int height)
        { if (width <= 0 || height <= 0 || (long)width * height > Budget.MaxImagePixels) throw new NotSupportedException("图片解码超过 1600 万像素或尺寸无效。"); }
        using var stream = new MemoryStream(image.Data.ToArray(), false);
        if (image.Extension is "tif" or "tiff")
        {
            using var tiff = Tiff.ClientOpen("image", "r", stream, new TiffStream()) ?? throw new InvalidDataException("TIFF 无法解码。");
            var width = tiff.GetField(TiffTag.IMAGEWIDTH)[0].ToInt(); var height = tiff.GetField(TiffTag.IMAGELENGTH)[0].ToInt(); Check(width, height);
            var pixels = new int[checked(width * height)];
            if (!tiff.ReadRGBAImageOriented(width, height, pixels, Orientation.TOPLEFT)) throw new InvalidDataException("TIFF 像素无法解码。");
            var bitmap = new SKBitmap(width, height, SKColorType.Bgra8888, SKAlphaType.Premul);
            try
            {
                // libtiff returns premultiplied ABGR. Reorder into Skia's native BGRA bytes.
                for (var i = 0; i < pixels.Length; i++)
                {
                    var p = (uint)pixels[i]; pixels[i] = (int)((p & 0xff00ff00) | (p & 0xff) << 16 | (p >> 16 & 0xff));
                }
                System.Runtime.InteropServices.Marshal.Copy(pixels, 0, bitmap.GetPixels(), pixels.Length); return bitmap;
            }
            catch { bitmap.Dispose(); throw; }
        }
        using var codec = SKCodec.Create(stream) ?? throw new InvalidDataException("图片无法解码。");
        Check(codec.Info.Width, codec.Info.Height);
        var result = new SKBitmap(new SKImageInfo(codec.Info.Width, codec.Info.Height, SKColorType.Bgra8888, SKAlphaType.Premul));
        var status = codec.GetPixels(result.Info, result.GetPixels());
        if (status != SKCodecResult.Success) { result.Dispose(); throw new InvalidDataException($"图片解码失败：{status}"); }
        return result;
    }
    public static void DrawImage(SKCanvas canvas, SKBitmap bitmap, SKRect bounds, bool monochrome)
    {
        using var paint = new SKPaint();
        using var filter = monochrome ? SKColorFilter.CreateColorMatrix([.299f, .587f, .114f, 0, 0, .299f, .587f, .114f, 0, 0, .299f, .587f, .114f, 0, 0, 0, 0, 0, 1, 0]) : null;
        paint.ColorFilter = filter;
        using var image = SKImage.FromBitmap(bitmap);
        canvas.DrawImage(image, bounds, new SKSamplingOptions(SKFilterMode.Linear), paint);
    }
    public void Dispose() => Fonts.Dispose();
}
