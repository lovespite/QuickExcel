using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

/// <summary>HSV mouse picker. External RGB edits move its markers without raising ColorChanged.</summary>
internal sealed class ColorSpectrum : Control
{
    private static readonly LinearGradientBrush Shade = new()
    {
        StartPoint = new(0, 0, RelativeUnit.Relative), EndPoint = new(0, 1, RelativeUnit.Relative),
        GradientStops = [new(Color.FromArgb(0, 0, 0, 0), 0), new(Colors.Black, 1)]
    };
    private static readonly LinearGradientBrush HueStrip = new()
    {
        StartPoint = new(0, 0, RelativeUnit.Relative), EndPoint = new(1, 0, RelativeUnit.Relative),
        GradientStops = [new(Colors.Red, 0), new(Colors.Yellow, 1d / 6), new(Colors.Lime, 2d / 6),
            new(Colors.Cyan, .5), new(Colors.Blue, 4d / 6), new(Colors.Magenta, 5d / 6), new(Colors.Red, 1)]
    };
    private Color _color = Colors.Black;
    private double _hue, _saturation, _value;
    private IPointer? _pointer;
    private bool _dragHue;

    public ColorSpectrum()
    {
        Focusable = true; Cursor = new Cursor(StandardCursorType.Cross);
        AutomationProperties.SetName(this, "色盘，左右调整饱和度，上下调整明度，Ctrl 加左右调整色相");
    }

    public event Action<Color>? ColorChanged;
    public Color SelectedColor
    {
        get => _color;
        set
        {
            if (_color == value) return;
            _color = value;
            var hsv = value.ToHsv();
            // Grey and black have no unique hue. Retain the chosen hue (and black's
            // saturation) so dragging out of either endpoint is continuous.
            if (hsv.S > 0) _hue = hsv.H;
            if (hsv.V > 0) _saturation = hsv.S;
            _value = hsv.V; InvalidateVisual();
        }
    }
    internal Rect SaturationValueBounds => new(6, 6, Math.Max(1, Bounds.Width - 12), Math.Max(1, Bounds.Height - 44));
    internal Rect HueBounds => new(6, Math.Max(6, Bounds.Height - 26), Math.Max(1, Bounds.Width - 12), 18);

    public override void Render(DrawingContext context)
    {
        var square = SaturationValueBounds; var hue = HueBounds;
        var saturation = new LinearGradientBrush
        {
            StartPoint = new(0, 0, RelativeUnit.Relative), EndPoint = new(1, 0, RelativeUnit.Relative),
            GradientStops = [new(Colors.White, 0), new(new HsvColor(1, _hue, 1, 1).ToRgb(), 1)]
        };
        context.DrawRectangle(saturation, null, square); context.DrawRectangle(Shade, new Pen(Brushes.Gray, 1), square);
        context.DrawRectangle(HueStrip, new Pen(Brushes.Gray, 1), hue);
        var marker = new Point(square.X + square.Width * _saturation, square.Y + square.Height * (1 - _value));
        context.DrawEllipse(null, new Pen(Brushes.Black, 3), marker, 5, 5);
        context.DrawEllipse(null, new Pen(Brushes.White, 1.5), marker, 5, 5);
        var hueMarker = new Rect(hue.X + hue.Width * _hue / 360 - 3, hue.Y - 2, 6, hue.Height + 4);
        context.DrawRectangle(null, new Pen(Brushes.Black, 3), hueMarker);
        context.DrawRectangle(null, new Pen(Brushes.White, 1.5), hueMarker);
    }

    private void Pick(Point point)
    {
        var bounds = _dragHue ? HueBounds : SaturationValueBounds;
        var x = Math.Clamp((point.X - bounds.X) / bounds.Width, 0, 1);
        if (_dragHue) _hue = x * 360;
        else { _saturation = x; _value = 1 - Math.Clamp((point.Y - bounds.Y) / bounds.Height, 0, 1); }
        PublishColor();
    }
    private void PublishColor()
    {
        _color = new HsvColor(_color.A / 255d, _hue, _saturation, _value).ToRgb();
        InvalidateVisual(); ColorChanged?.Invoke(_color);
    }
    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        var point = e.GetPosition(this);
        if (_pointer is not null || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed ||
            !SaturationValueBounds.Contains(point) && !HueBounds.Contains(point)) return;
        Focus(); _pointer = e.Pointer; _dragHue = HueBounds.Contains(point);
        e.Pointer.Capture(this); Pick(point); e.Handled = true;
    }
    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (e.Pointer != _pointer) return;
        Pick(e.GetPosition(this)); e.Handled = true;
    }
    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (e.Pointer != _pointer || e.InitialPressMouseButton != MouseButton.Left) return;
        Pick(e.GetPosition(this)); _pointer = null; e.Pointer.Capture(null); e.Handled = true;
    }
    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
    {
        base.OnPointerCaptureLost(e); _pointer = null;
    }
    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        var pointer = _pointer; _pointer = null; pointer?.Capture(null); base.OnDetachedFromVisualTree(e);
    }
    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (e.Key is not (Key.Left or Key.Right or Key.Up or Key.Down)) return;
        var step = e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? .1 : .01;
        if (e.KeyModifiers.HasFlag(KeyModifiers.Control)) _hue = (_hue + (e.Key is Key.Right or Key.Up ? 1 : -1) * step * 100 + 360) % 360;
        else if (e.Key is Key.Left or Key.Right) _saturation = Math.Clamp(_saturation + (e.Key == Key.Right ? step : -step), 0, 1);
        else _value = Math.Clamp(_value + (e.Key == Key.Up ? step : -step), 0, 1);
        PublishColor(); e.Handled = true;
    }
}
