using System.Diagnostics;
using System.Runtime.CompilerServices;
using Avalonia;

[assembly: InternalsVisibleTo("QuickExcel.Avalonia")]

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    // Diagnostic entry uses the production animation scheduler; it does not claim OS pointer input.
    internal event Action<long>? DragFrameApplied;
    internal void SetBenchmarkHorizontalOffset(double value) => _horizontal.Value = value;
    internal void BeginBenchmarkDrag(Point outsideViewport)
    {
        CancelGesture(); _anchor = _current; _dragAxis = null; _dragPoint = outsideViewport;
        _dragging = true; _lastDragFrame = null; RequestDragFrame();
    }
    public void CancelInputGesture() => CancelGesture();
}
