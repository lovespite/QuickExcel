using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Media;
using FastWpfGrid;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using SkiaSharp;

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    private readonly Button _chartFilterButton = new() { Width = 30, Height = 30, Padding = new(5), IsVisible = false,
        Background = Brushes.White, BorderBrush = Brushes.Gray, BorderThickness = new(1) };
    public event Action? ChartFilterRequested;
    public Rect? ChartFilterButtonBounds => _chartFilterButton.IsVisible ? _chartFilterButton.Bounds : null;
    public IReadOnlyList<ChartSourceRange> SelectedChartSources => _selectedObject is { Kind: DrawingObjectKind.Chart } id && ObjectAnchor(id) is not null
        ? ChartSelection.Sources(Sheet.Data.Charts[id.Index]).Where(s => s.Sheet.Equals(Sheet.Name, StringComparison.OrdinalIgnoreCase)).ToArray() : [];

    private void InstallChartControls()
    {
        _chartFilterButton.Content = new global::Avalonia.Controls.Shapes.Path { Data = Geometry.Parse("M 1,1 L 17,1 L 11,8 L 11,16 L 7,13 L 7,8 Z"),
            Stroke = Brushes.SeaGreen, StrokeThickness = 1.5, Fill = Brushes.Honeydew, Width = 18, Height = 18 };
        AutomationProperties.SetName(_chartFilterButton, "图表筛选器"); ToolTip.SetTip(_chartFilterButton, "图表筛选器：选择显示的系列和类别");
        _chartFilterButton.Click += (_, e) => { e.Handled = true; ChartFilterRequested?.Invoke(); };
        _canvas.Children.Add(_chartFilterButton);
    }
    private void PositionChartControls()
    {
        _chartFilterButton.IsVisible = false;
        if (_snapshot is null || _selectedObject is not { Kind: DrawingObjectKind.Chart } id || ObjectAnchor(id) is not { } anchor || IsDrawingObjectDragging) return;
        var visible = ObjectFragments(anchor).Select(f => f.Bounds.Intersect(f.Clip)).Where(r => r.Width > 0 && r.Height > 0).OrderByDescending(r => r.Width * r.Height).FirstOrDefault();
        if (visible.Width <= 0 || _canvas.Bounds.Width < 32 || _canvas.Bounds.Height < 32) return;
        var x = visible.Right * _zoom + 8; if (x + 30 > _canvas.Bounds.Width) x = visible.Right * _zoom - 38;
        Canvas.SetLeft(_chartFilterButton, Math.Clamp(x, 0, _canvas.Bounds.Width - 30));
        Canvas.SetTop(_chartFilterButton, Math.Clamp(visible.Top * _zoom + 6, 0, _canvas.Bounds.Height - 30));
        _chartFilterButton.IsVisible = true;
        _chartFilterButton.IsEnabled = !IsComposing && !IsCommitPending;
        _chartFilterButton.Background = Sheet.Data.Charts[id.Index].Series.Any(s => s.IsFiltered || !s.HiddenPoints.IsEmpty) ? Brushes.Honeydew : Brushes.White;
    }
    private void PaintChartSources(SKCanvas canvas)
    {
        using var paint = new SKPaint { IsAntialias = true };
        foreach (var source in SelectedChartSources)
        foreach (var fragment in Fragments(new GridCellRange(source.Range.Top, source.Range.Left, source.Range.Bottom, source.Range.Right)))
        {
            var color = SKColor.Parse("#" + source.Color); var bounds = SkiaRect(fragment.Bounds);
            canvas.Save(); canvas.ClipRect(SkiaRect(fragment.Clip));
            paint.Style = SKPaintStyle.Fill; paint.Color = color.WithAlpha(20); canvas.DrawRect(bounds, paint);
            paint.Style = SKPaintStyle.Stroke; paint.Color = color; paint.StrokeWidth = (float)(1.5 / _zoom); canvas.DrawRect(bounds, paint);
            paint.Style = SKPaintStyle.Fill; var size = (float)(5 / _zoom);
            foreach (var p in new[] { new SKPoint(bounds.Left, bounds.Top), new SKPoint(bounds.Right, bounds.Top), new SKPoint(bounds.Left, bounds.Bottom), new SKPoint(bounds.Right, bounds.Bottom) })
                canvas.DrawRect(SKRect.Create(p.X - size / 2, p.Y - size / 2, size, size), paint);
            canvas.Restore();
        }
    }
}
