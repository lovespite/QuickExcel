using System.Windows;
using System.Windows.Media;

namespace FastWpfGrid;

/// <summary>A view-only retained layer. Selection updates never allocate a viewport bitmap.</summary>
public sealed class GridSelectionOverlay : FrameworkElement
{
    internal FastGridControl Grid { get; set; }
    private bool _pending;
    internal bool IsPending => _pending && IsVisible;

    internal void Refresh()
    {
        _pending = true;
        InvalidateVisual();
    }

    protected override void OnRender(DrawingContext drawingContext)
    {
        base.OnRender(drawingContext);
        Grid?.RenderSelectionOverlay(drawingContext);
        _pending = false;
    }
}
