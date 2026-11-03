using QuickExcel.Domain;

namespace QuickExcel.Presentation;

public enum DrawingHandle { Move, TopLeft, Top, TopRight, Right, BottomRight, Bottom, BottomLeft, Left, Rotate }
public readonly record struct DrawingPoint(double X, double Y);
public readonly record struct DrawingPlacement(DrawingBounds Bounds, double Rotation);

/// <summary>Screen-independent geometry for mouse/touch transforms; coordinates are 96-DPI pixels.</summary>
public static class DrawingObjectInteraction
{
    public static DrawingPoint Rotate(DrawingPoint point, DrawingPoint center, double degrees)
    {
        var a = degrees * Math.PI / 180; var x = point.X - center.X; var y = point.Y - center.Y;
        return new(center.X + x * Math.Cos(a) - y * Math.Sin(a), center.Y + x * Math.Sin(a) + y * Math.Cos(a));
    }
    public static DrawingPoint HandlePoint(DrawingPlacement placement, DrawingHandle handle, double rotationDistance = 28)
    {
        var b = placement.Bounds; var (x, y) = Factors(handle); var center = new DrawingPoint(b.X + b.Width / 2, b.Y + b.Height / 2);
        var point = handle == DrawingHandle.Rotate ? new(center.X, b.Y - rotationDistance) : new DrawingPoint(center.X + x * b.Width / 2, center.Y + y * b.Height / 2);
        return Rotate(point, center, placement.Rotation);
    }
    public static bool Contains(DrawingPlacement placement, DrawingPoint point)
    {
        var b = placement.Bounds; var p = Rotate(point, new(b.X + b.Width / 2, b.Y + b.Height / 2), -placement.Rotation);
        return p.X >= b.X && p.X <= b.X + b.Width && p.Y >= b.Y && p.Y <= b.Y + b.Height;
    }
    private static (int X, int Y) Factors(DrawingHandle handle) => handle switch
    {
        DrawingHandle.TopLeft => (-1, -1), DrawingHandle.Top => (0, -1), DrawingHandle.TopRight => (1, -1), DrawingHandle.Right => (1, 0),
        DrawingHandle.BottomRight => (1, 1), DrawingHandle.Bottom => (0, 1), DrawingHandle.BottomLeft => (-1, 1), DrawingHandle.Left => (-1, 0), _ => (0, 0)
    };
    public static DrawingPlacement Update(DrawingPlacement original, DrawingHandle handle, DrawingPoint start, DrawingPoint current, bool constrain = false)
    {
        var b = original.Bounds; var center = new DrawingPoint(b.X + b.Width / 2, b.Y + b.Height / 2);
        if (handle == DrawingHandle.Rotate)
        {
            var delta = (Math.Atan2(current.Y - center.Y, current.X - center.X) - Math.Atan2(start.Y - center.Y, start.X - center.X)) * 180 / Math.PI;
            var angle = original.Rotation + delta; if (constrain) angle = Math.Round(angle / 15) * 15;
            // DrawingML stores 1/60000 degrees. Quantize here to keep save/reopen equality exact.
            return original with { Rotation = DrawingGeometry.NormalizeRotation(Math.Round(angle * 60000) / 60000) };
        }
        if (handle == DrawingHandle.Move)
        {
            var dx = current.X - start.X; var dy = current.Y - start.Y;
            if (constrain) { if (Math.Abs(dx) > Math.Abs(dy)) dy = 0; else dx = 0; }
            return original with { Bounds = b with { X = Math.Max(0, b.X + dx), Y = Math.Max(0, b.Y + dy) } };
        }
        var from = Rotate(start, center, -original.Rotation); var to = Rotate(current, center, -original.Rotation); var (hx, hy) = Factors(handle);
        var width = hx == 0 ? b.Width : Math.Clamp(b.Width + (to.X - from.X) * hx, 8, 100000);
        var height = hy == 0 ? b.Height : Math.Clamp(b.Height + (to.Y - from.Y) * hy, 8, 100000);
        if (constrain && hx != 0 && hy != 0)
        {
            var scale = Math.Max(width / b.Width, height / b.Height);
            scale = Math.Clamp(scale, Math.Max(8 / b.Width, 8 / b.Height), Math.Min(100000 / b.Width, 100000 / b.Height)); width = b.Width * scale; height = b.Height * scale;
        }
        var shifted = Rotate(new(center.X + hx * (width - b.Width) / 2, center.Y + hy * (height - b.Height) / 2), center, original.Rotation);
        return original with { Bounds = new(Math.Max(0, shifted.X - width / 2), Math.Max(0, shifted.Y - height / 2), width, height) };
    }
}
