// QuickExcel host boundary. Only view geometry is exposed; this assembly has no document dependency.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;

namespace FastWpfGrid
{
    partial class FastGridControl
    {
        private bool _hasModelDimensions;
        public event EventHandler ViewportChanged;
        public void SetContentDecoration(FrameworkElement decoration)
        {
            decoration.IsHitTestVisible = false;
            imageGrid.Children.Insert(imageGrid.Children.IndexOf(selectionOverlay), decoration);
        }
        public Rect? GetModelCellBounds(int row, int column)
        {
            var real = ModelToReal(new FastGridCellAddress(row, column));
            if (!real.IsCell || real.Row < 0 || real.Column < 0 ||
                !_rowSizes.IsVisible(real.Row.Value, FirstVisibleRowScrollIndex, GridScrollAreaHeight + ScrollOffsetY) ||
                !_columnSizes.IsVisible(real.Column.Value, FirstVisibleColumnScrollIndex, GridScrollAreaWidth + ScrollOffsetX)) return null;
            var rect = GetCellRect(real.Row.Value, real.Column.Value);
            return new Rect(rect.Left / DpiDetector.DpiXKoef, rect.Top / DpiDetector.DpiYKoef,
                Math.Max(0, (rect.Right - rect.Left) / DpiDetector.DpiXKoef), Math.Max(0, (rect.Bottom - rect.Top) / DpiDetector.DpiYKoef));
        }
        public IReadOnlyList<GridCellRange> GetSelectedModelRanges()
        {
            static List<(int First, int Last)> Runs(IEnumerable<int> values)
            {
                var result = new List<(int First, int Last)>();
                foreach (var value in values.Distinct().OrderBy(v => v))
                {
                    if (result.Count > 0 && result[^1].Last + 1 == value) result[^1] = (result[^1].First, value);
                    else result.Add((value, value));
                }
                return result;
            }
            var result = new List<GridCellRange>();
            foreach (var range in SelectedRanges)
            {
                if (_rowSizes.FrozenCount == 0 && _columnSizes.FrozenCount == 0 && _rowSizes.RealCount == _rowSizes.Count && _columnSizes.RealCount == _columnSizes.Count)
                {
                    result.Add(IsTransposed ? new(range.Left, range.Top, range.Right, range.Bottom) : range);
                    continue;
                }
                var rows = Runs(Enumerable.Range(range.Top, range.Bottom - range.Top + 1).Select(_rowSizes.RealToModel));
                var columns = Runs(Enumerable.Range(range.Left, range.Right - range.Left + 1).Select(_columnSizes.RealToModel));
                foreach (var row in rows) foreach (var column in columns)
                    result.Add(IsTransposed ? new(column.First, row.First, column.Last, row.Last) : new(row.First, column.First, row.Last, column.Last));
            }
            return result;
        }
        public void SetModelDimensions(IReadOnlyDictionary<int, double> rows, IReadOnlyDictionary<int, double> columns, double? defaultRowHeight = null)
        {
            _hasModelDimensions = true;
            if (defaultRowHeight.HasValue) _rowSizes.DefaultSize = Math.Max(1, (int)Math.Round(defaultRowHeight.Value * DpiDetector.DpiYKoef));
            _rowSizes.ReplaceSizeOverrides(rows.ToDictionary(p => p.Key, p => (int)Math.Round(p.Value * DpiDetector.DpiYKoef)));
            _columnSizes.ReplaceSizeOverrides(columns.ToDictionary(p => p.Key, p => (int)Math.Round(p.Value * DpiDetector.DpiXKoef)));
            // NotifyRefresh runs before the host installs the new sheet's dimensions.
            // Recompute header geometry too, instead of retaining the previous sheet's height.
            RecalculateHeaderSize();
            AdjustScrollbars(); InvalidateAll(); NotifyViewportChanged();
        }
        public FastGridCellAddress HitTestModel(Point point) => RealToModel(GetCellAddress(new Point(point.X * DpiDetector.DpiXKoef, point.Y * DpiDetector.DpiYKoef)));
        public IReadOnlyList<int> GetVisibleModelRows() => VisibleAxis(_rowSizes, FirstVisibleRowScrollIndex, VisibleRowCount, HeaderHeight, _drawBuffer?.PixelHeight ?? 0)
            .SelectMany(a => Enumerable.Range(a.First, Math.Max(0, Math.Min(a.Last + 1, _realRowCount) - a.First))).Select(_rowSizes.RealToModel).ToArray();
        public IReadOnlyList<int> GetVisibleModelColumns() => VisibleAxis(_columnSizes, FirstVisibleColumnScrollIndex, VisibleColumnCount, HeaderWidth, _drawBuffer?.PixelWidth ?? 0)
            .SelectMany(a => Enumerable.Range(a.First, Math.Max(0, Math.Min(a.Last + 1, _realColumnCount) - a.First))).Select(_columnSizes.RealToModel).ToArray();
        public double RowHeaderWidth => HeaderWidth / DpiDetector.DpiXKoef;
        public double ColumnHeaderHeight => HeaderHeight / DpiDetector.DpiYKoef;
        public int FrozenRowCount => _rowSizes.FrozenCount;
        public int FrozenColumnCount => _columnSizes.FrozenCount;

        /// <summary>Full object geometry and visible clip for each frozen/scrolling pane, in DIP.</summary>
        public IReadOnlyList<(Rect Bounds, Rect Clip)> GetModelRangeFragments(GridCellRange range)
        {
            var result = new List<(Rect, Rect)>();
            if (_drawBuffer == null) return result;
            var rr = _rowSizes.GetRealRanges(range.Top, range.Bottom);
            var cc = _columnSizes.GetRealRanges(range.Left, range.Right);
            if (rr.Count == 0 || cc.Count == 0) return result;
            double Position(SeriesSizes sizes, int index) => index < sizes.FrozenCount ? sizes.GetFrozenPosition(index) :
                sizes.FrozenSize + sizes.GetPositionByScrollIndex(index - sizes.FrozenCount);
            var top = rr.Min(r => r.First); var bottom = rr.Max(r => r.Last) + 1;
            var left = cc.Min(c => c.First); var right = cc.Max(c => c.Last) + 1;
            var dx = DpiDetector.DpiXKoef; var dy = DpiDetector.DpiYKoef;
            foreach (var rows in VisibleAxis(_rowSizes, FirstVisibleRowScrollIndex, VisibleRowCount, HeaderHeight, _drawBuffer.PixelHeight))
            foreach (var columns in VisibleAxis(_columnSizes, FirstVisibleColumnScrollIndex, VisibleColumnCount, HeaderWidth, _drawBuffer.PixelWidth))
            {
                var x = HeaderWidth + Position(_columnSizes, left) - (columns.First < _columnSizes.FrozenCount ? 0 : _columnSizes.GetPositionByScrollIndex(FirstVisibleColumnScrollIndex) + ScrollOffsetX);
                var y = HeaderHeight + Position(_rowSizes, top) - (rows.First < _rowSizes.FrozenCount ? 0 : _rowSizes.GetPositionByScrollIndex(FirstVisibleRowScrollIndex) + ScrollOffsetY);
                var rect = new Rect(x / dx, y / dy, (Position(_columnSizes, right) - Position(_columnSizes, left)) / dx,
                    (Position(_rowSizes, bottom) - Position(_rowSizes, top)) / dy);
                var clip = new Rect(columns.Min / dx, rows.Min / dy, (columns.Max - columns.Min) / dx, (rows.Max - rows.Min) / dy);
                if (rect.IntersectsWith(clip)) result.Add((rect, clip));
            }
            return result;
        }
        private void NotifyViewportChanged() => ViewportChanged?.Invoke(this, EventArgs.Empty);
    }
}
