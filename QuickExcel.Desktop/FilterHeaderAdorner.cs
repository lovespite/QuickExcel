using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace QuickExcel.Desktop;

/// <summary>Buttons belong to data-header cells, not the A/B/C column labels. Only realized cells have buttons.</summary>
internal sealed class FilterHeaderAdorner : Adorner
{
    private readonly DataGrid _grid;
    private readonly WpfGridSurface _surface;
    private readonly Canvas _canvas = new();
    private readonly Dictionary<int, Button> _buttons = [];
    private readonly Action<int, FrameworkElement> _open;
    public FilterHeaderAdorner(DataGrid grid, WpfGridSurface surface, Action<int, FrameworkElement> open) : base(grid)
    {
        _grid = grid; _surface = surface; _open = open; AddVisualChild(_canvas); AddLogicalChild(_canvas);
        ClipToBounds = true;
    }
    public int ButtonCount => _buttons.Count;
    public void Refresh()
    {
        var region = _surface.View.Region;
        var visible = new HashSet<int>();
        if (region is { HasHeader: true } && _grid.IsVisible)
        {
            var item = _surface.Rows[region.Range.Top];
            for (var c = region.Range.Left; c <= Math.Min(region.Range.Right, _grid.Columns.Count - 1); c++)
            {
                if (_grid.Columns[c].Visibility != Visibility.Visible || _grid.Columns[c].GetCellContent(item) is not FrameworkElement content ||
                    content.Parent is not DataGridCell cell || !cell.IsVisible) continue;
                var bounds = cell.TransformToAncestor(_grid).TransformBounds(new Rect(cell.RenderSize));
                if (bounds.Bottom <= _grid.ColumnHeaderHeight || bounds.Top >= _grid.ActualHeight || bounds.Right <= 0 || bounds.Right > _grid.ActualWidth) continue;
                // A scrolling column underneath frozen columns must not leave an active button over the frozen pane.
                var frozenWidth = _grid.RowHeaderActualWidth + _grid.Columns.Take(_grid.FrozenColumnCount).Sum(column => column.ActualWidth);
                var marker = _surface.View.HeaderMarker(c);
                var width = marker.Length > 2 ? 34d : 23d;
                if (bounds.Width < width || !cell.Column.IsFrozen && bounds.Right - width < frozenWidth) continue;
                visible.Add(c);
                if (!_buttons.TryGetValue(c, out var button))
                {
                    var column = c;
                    button = new Button { Width = 19, Height = 20, Padding = new(0), FontSize = 10, Focusable = true };
                    button.Click += (_, _) => _open(column, button);
                    _buttons.Add(c, button); _canvas.Children.Add(button);
                }
                button.Width = width;
                if (!Equals(button.Content, marker)) button.Content = marker;
                button.ToolTip = $"{MainWindow.ColumnName(c)} 列筛选和排序" + (_surface.View.Filters.ContainsKey(c) ? "（已筛选）" : "");
                System.Windows.Automation.AutomationProperties.SetName(button, button.ToolTip.ToString());
                Canvas.SetLeft(button, bounds.Right - width); Canvas.SetTop(button, bounds.Top + Math.Max(0, (bounds.Height - 20) / 2));
            }
        }
        foreach (var column in _buttons.Keys.Where(column => !visible.Contains(column)).ToArray())
        { _canvas.Children.Remove(_buttons[column]); _buttons.Remove(column); }
    }
    protected override int VisualChildrenCount => 1;
    protected override Visual GetVisualChild(int index) => _canvas;
    protected override Size MeasureOverride(Size constraint) { _canvas.Measure(AdornedElement.RenderSize); return AdornedElement.RenderSize; }
    protected override Size ArrangeOverride(Size finalSize) { _canvas.Arrange(new Rect(finalSize)); return finalSize; }
    protected override HitTestResult? HitTestCore(PointHitTestParameters hitTestParameters) => null;
}
