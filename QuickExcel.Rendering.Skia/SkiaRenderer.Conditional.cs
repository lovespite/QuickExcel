using QuickExcel.Domain;
using SkiaSharp;

namespace QuickExcel.Rendering.Skia;

public sealed partial class SkiaRenderer
{
    private static SKRect DrawConditional(SKCanvas canvas, ConditionalCellVisual? visual, SKRect rect, bool monochrome)
    {
        if (visual is null) return rect;
        canvas.Save(); canvas.ClipRect(rect);
        if (visual.Bar is { } bar && rect.Width > 4 && rect.Height > 4)
        {
            var width = rect.Width - 4; var start = rect.Left + 2 + (float)bar.Start * width; var end = rect.Left + 2 + (float)bar.End * width;
            var color = monochrome ? SKColors.Gray : Color(bar.Color, SKColors.CornflowerBlue);
            using var shader = bar.Gradient && start != end ? SKShader.CreateLinearGradient(new(start, 0), new(end, 0), [color, new SKColor(color.Red, color.Green, color.Blue, 30)], SKShaderTileMode.Clamp) : null;
            using var paint = new SKPaint { Color = color, Shader = shader, IsAntialias = true };
            var bounds = new SKRect(Math.Min(start, end), rect.Top + 3, Math.Max(start, end), rect.Bottom - 3);
            canvas.DrawRect(bounds, paint); paint.Shader = null;
            if (bar.BorderColor is { } border && bounds.Width > 0)
            { paint.Color = monochrome ? SKColors.Black : Color(border, color); paint.Style = SKPaintStyle.Stroke; paint.StrokeWidth = 1; canvas.DrawRect(bounds, paint); }
            if (bar.Axis is { } axis && axis > 0 && axis < 1)
            {
                paint.Color = monochrome ? SKColors.Black : Color(bar.AxisColor, SKColors.Black); paint.StrokeWidth = 1;
                var x = rect.Left + 2 + (float)axis * width; canvas.DrawLine(x, rect.Top + 2, x, rect.Bottom - 2, paint);
            }
        }
        if (visual.Icon is { } icon && rect.Width > 4 && rect.Height > 4)
        {
            var size = Math.Min(16, Math.Min(rect.Width - 4, rect.Height - 4));
            DrawConditionalIcon(canvas, icon, SKRect.Create(rect.Left + 3, rect.MidY - size / 2, size, size), monochrome);
            rect.Left += size + 5;
        }
        canvas.Restore(); return rect;
    }

    public static void DrawConditionalIcon(SKCanvas canvas, ConditionalIconVisual icon, SKRect bounds, bool monochrome = false)
    {
        var count = icon.Set[0] - '0'; var index = Math.Clamp(icon.Index, 0, count - 1);
        var gray = monochrome || icon.Set.Contains("Gray", StringComparison.Ordinal);
        var color = gray ? new SKColor((byte)(60 + index * 25), (byte)(60 + index * 25), (byte)(60 + index * 25)) :
            index == 0 ? new SKColor(220, 65, 57) : index == count - 1 ? new SKColor(61, 160, 83) : new SKColor(231, 177, 45);
        using var paint = new SKPaint { Color = color, IsAntialias = true, StrokeWidth = 1.2f, StrokeCap = SKStrokeCap.Round, StrokeJoin = SKStrokeJoin.Round };
        canvas.Save(); canvas.Translate(bounds.Left, bounds.Top); canvas.Scale(bounds.Width / 16, bounds.Height / 16);
        void Path(params SKPoint[] points)
        {
            using var builder = new SKPathBuilder(); builder.MoveTo(points[0]); foreach (var p in points.Skip(1)) builder.LineTo(p); builder.Close();
            using var path = builder.Detach(); canvas.DrawPath(path, paint);
        }
        if (icon.Set.Contains("Arrows", StringComparison.Ordinal))
        {
            var angle = count == 3 ? new[] { 180f, 90f, 0f }[index] : count == 4 ? new[] { 180f, 135f, 45f, 0f }[index] : new[] { 180f, 135f, 90f, 45f, 0f }[index];
            canvas.RotateDegrees(angle, 8, 8); Path(new(8, 1), new(15, 8), new(11, 8), new(11, 15), new(5, 15), new(5, 8), new(1, 8));
        }
        else if (icon.Set == "3Flags")
        { paint.Style = SKPaintStyle.Stroke; canvas.DrawLine(3, 2, 3, 15, paint); paint.Style = SKPaintStyle.Fill; Path(new(4, 2), new(14, 3), new(13, 10), new(4, 9)); }
        else if (icon.Set.Contains("Rating", StringComparison.Ordinal))
        {
            for (var i = 0; i < count; i++)
            { paint.Color = i <= index ? monochrome ? SKColors.Black : new SKColor(66, 136, 193) : new SKColor(218, 224, 229); canvas.DrawRect(i * 15f / count, 14 - (i + 1) * 12f / count, 12f / count, (i + 1) * 12f / count, paint); }
        }
        else if (icon.Set == "5Quarters")
        {
            paint.Color = SKColors.White; canvas.DrawCircle(8, 8, 6.5f, paint); paint.Color = monochrome ? SKColors.Black : new SKColor(74, 99, 126);
            if (index == 4) canvas.DrawCircle(8, 8, 6.5f, paint);
            else if (index > 0) canvas.DrawArc(new SKRect(1.5f, 1.5f, 14.5f, 14.5f), -90, index * 90, true, paint);
            paint.Style = SKPaintStyle.Stroke; canvas.DrawCircle(8, 8, 6.5f, paint);
        }
        else if (icon.Set is "3Symbols" or "3Symbols2")
        {
            if (icon.Set == "3Symbols") { canvas.DrawCircle(8, 8, 7, paint); paint.Color = SKColors.White; }
            paint.Style = SKPaintStyle.Stroke; paint.StrokeWidth = 2;
            if (index == 0) { canvas.DrawLine(5, 5, 11, 11, paint); canvas.DrawLine(11, 5, 5, 11, paint); }
            else if (index == 1) { canvas.DrawLine(8, 4, 8, 9, paint); canvas.DrawPoint(8, 12, paint); }
            else { canvas.DrawLine(4, 8, 7, 11, paint); canvas.DrawLine(7, 11, 12, 5, paint); }
        }
        else if (icon.Set == "3Signs")
        {
            if (index == 0) Path(new(8, 1), new(15, 8), new(8, 15), new(1, 8));
            else if (index == 1) Path(new(8, 1), new(15, 14), new(1, 14));
            else canvas.DrawCircle(8, 8, 7, paint);
        }
        else
        {
            if (icon.Set == "4RedToBlack") paint.Color = monochrome ? new SKColor((byte)(index * 65), (byte)(index * 65), (byte)(index * 65)) : new SKColor((byte)(index * 75), (byte)(index * 12), (byte)(index * 12));
            if (icon.Set == "4TrafficLights" && index == 0) paint.Color = SKColors.Black;
            else if (icon.Set == "4TrafficLights" && index == 1) paint.Color = monochrome ? SKColors.Gray : new SKColor(220, 65, 57);
            if (icon.Set == "3TrafficLights2") { var fill = paint.Color; paint.Color = new SKColor(55, 55, 55); canvas.DrawRoundRect(new SKRect(2, 0, 14, 16), 3, 3, paint); paint.Color = fill; canvas.DrawCircle(8, 8, 4.8f, paint); }
            else canvas.DrawCircle(8, 8, 6.5f, paint);
        }
        canvas.Restore();
    }
}
