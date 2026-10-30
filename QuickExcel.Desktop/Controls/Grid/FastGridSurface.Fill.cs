using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

internal sealed partial class FastGridSurface
{
    internal Thumb FillHandle { get; } = new() { Width = 8, Height = 8, Background = Brushes.SeaGreen, BorderBrush = Brushes.White,
        BorderThickness = new(1), Cursor = Cursors.Cross, ToolTip = "沿行或列拖动，复制单元格和公式；数字/日期序列请使用填充菜单", Visibility = Visibility.Hidden };
    internal Func<bool>? BeforeFill { get; set; }
    internal event Action<GridRange, GridRange>? FillRequested;
    private GridRange? _fillSource, _fillExtent;
    private GridRange? _fillHandleSelection;
    private bool _fillSelectionDirty = true;
    private FillPreview? _fillPreview;

    private void InstallFillHandle()
    {
        _fillPreview = new(this) { IsHitTestVisible = false }; _overlay.Children.Insert(0, _fillPreview); _overlay.Children.Add(FillHandle);
        System.Windows.Automation.AutomationProperties.SetName(FillHandle, "复制填充柄");
        FillHandle.DragStarted += (_, _) => { if (!BeginFillGesture()) FillHandle.CancelDrag(); };
        FillHandle.DragDelta += (_, _) =>
        {
            var hit = Control.HitTestModel(Mouse.GetPosition(Control));
            PreviewFillAt(hit.IsCell ? new CellPosition(hit.Row!.Value, hit.Column!.Value) : null);
        };
        FillHandle.DragCompleted += (_, e) => CompleteFillGesture(e.Canceled);
        Element.Unloaded += (_, _) => CancelFillGesture();
    }

    internal bool BeginFillGesture()
    {
        if (_snapshot is null || _editing || IsComposing || _snapshot.FileState?.FormulaReadOnly == true || BeforeFill?.Invoke() == false) return false;
        if (Control.SelectedCellCount > 20_000) return false;
        if (GetFillSelection() is not { } selection) return false;
        _fillSource = selection; _fillExtent = null; Editor.Visibility = Visibility.Hidden;
        return true;
    }
    internal void PreviewFillAt(CellPosition? target)
    {
        _fillExtent = null;
        if (_fillSource is { } source && target is { } point && point.Row >= 0 && point.Row < RowCount && point.Column >= 0 && point.Column < ColumnCount)
        {
            var rows = Math.Max(source.Top - point.Row, point.Row - source.Bottom);
            var columns = Math.Max(source.Left - point.Column, point.Column - source.Right);
            if (rows > 0 || columns > 0)
                _fillExtent = rows > 0 && (columns <= 0 || (long)rows * source.Columns >= (long)columns * source.Rows)
                    ? new(Math.Min(source.Top, point.Row), source.Left, Math.Max(source.Bottom, point.Row), source.Right)
                    : new(source.Top, Math.Min(source.Left, point.Column), source.Bottom, Math.Max(source.Right, point.Column));
            if (_fillExtent?.Count > 20_000) _fillExtent = null;
        }
        _fillPreview?.InvalidateVisual();
    }
    internal void CompleteFillGesture(bool cancel)
    {
        var source = _fillSource; var extent = _fillExtent; _fillSource = _fillExtent = null;
        _fillPreview?.InvalidateVisual();
        if (!cancel && source.HasValue && extent.HasValue) FillRequested?.Invoke(source.Value, extent.Value);
        PrepareEditor(false); PositionFillHandle();
    }
    internal void CancelFillGesture()
    {
        CompleteFillGesture(true);
        if (FillHandle.IsDragging) FillHandle.CancelDrag();
    }
    private void PositionFillHandle()
    {
        FillHandle.Visibility = Visibility.Hidden;
        _fillPreview?.InvalidateVisual();
        if (_snapshot is null || _editing || IsComposing || Control.IsSelectionDragging || _snapshot.FileState?.FormulaReadOnly == true) return;
        if (Control.SelectedCellCount > 20_000) return;
        if (_fillSelectionDirty)
        {
            _fillHandleSelection = GetFillSelection();
            _fillSelectionDirty = false;
        }
        if (_fillHandleSelection is not { } range) return;
        var fragments = Control.GetModelRangeFragments(new(range.Bottom, range.Right, range.Bottom, range.Right));
        if (fragments.Count == 0) return;
        foreach (var fragment in fragments)
        {
            var point = new Point(fragment.Bounds.Right - 1, fragment.Bounds.Bottom - 1);
            if (!fragment.Clip.Contains(point)) continue;
            Canvas.SetLeft(FillHandle, point.X - 4); Canvas.SetTop(FillHandle, point.Y - 4); FillHandle.Visibility = Visibility.Visible; break;
        }
    }
    private GridRange? GetFillSelection()
    {
        var ranges = SelectionRanges;
        if (ranges.Count == 0) return null;
        // Accept contiguous selection fragments only when they cover the entire rectangle.
        var rectangle = new GridRange(ranges.Min(r => r.Top), ranges.Min(r => r.Left), ranges.Max(r => r.Bottom), ranges.Max(r => r.Right));
        return rectangle.Count <= 20_000 && ranges.Sum(r => r.Count) == rectangle.Count ? rectangle : null;
    }
    private sealed class FillPreview(FastGridSurface surface) : FrameworkElement
    {
        protected override void OnRender(DrawingContext context)
        {
            if (surface._fillExtent is not { } range) return;
            var pen = new Pen(Brushes.SeaGreen, 2) { DashStyle = DashStyles.Dash };
            foreach (var fragment in surface.Control.GetModelRangeFragments(new(range.Top, range.Left, range.Bottom, range.Right)))
            {
                context.PushClip(new RectangleGeometry(fragment.Clip)); context.DrawRectangle(null, pen, fragment.Bounds); context.Pop();
            }
        }
    }
}
