using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

/// <summary>Small vector samples of conditional formatting for the ribbon and its menu.</summary>
internal sealed class ConditionalFormatIcon(int kind) : Control
{
    public ConditionalFormatIcon() : this(0) { }
    protected override Size MeasureOverride(Size availableSize) => new(20, 20);
    public override void Render(DrawingContext context)
    {
        var line = new Pen(Brushes.DimGray, .8);
        IBrush red = SolidColorBrush.Parse("#F4878C"), blue = SolidColorBrush.Parse("#6BB5E7"), green = SolidColorBrush.Parse("#79BC84"), yellow = SolidColorBrush.Parse("#F4D777");
        context.DrawRectangle(Brushes.White, line, new Rect(1, 1, 18, 18));
        for (var row = 0; row < 3; row++)
        {
            if (kind == 2) context.DrawRectangle(blue, null, new Rect(2, 2 + row * 6, 15 - row * 4, 4));
            else if (kind == 3) context.DrawRectangle(row == 0 ? green : row == 1 ? yellow : red, null, new Rect(2, 2 + row * 6, 16, 4));
            else if (kind == 4)
            {
                context.DrawEllipse(row == 0 ? green : row == 1 ? yellow : red, null, new Point(5, 4 + row * 6), 2, 2);
                context.DrawLine(line, new Point(10, 4 + row * 6), new Point(16, 4 + row * 6));
            }
            else context.DrawRectangle(kind == 1 ? yellow : red, null, new Rect(7, 2 + row * 6, 5, 4));
        }
        if (kind != 4)
        {
            for (var n = 1; n < 3; n++)
            {
                context.DrawLine(line, new Point(1, 1 + n * 6), new Point(19, 1 + n * 6));
                if (kind is 0 or 1 or 5) context.DrawLine(line, new Point(1 + n * 6, 1), new Point(1 + n * 6, 19));
            }
        }
        if (kind == 5) context.DrawLine(new Pen(Brushes.MediumVioletRed, 3), new Point(9, 18), new Point(18, 9));
    }
}
