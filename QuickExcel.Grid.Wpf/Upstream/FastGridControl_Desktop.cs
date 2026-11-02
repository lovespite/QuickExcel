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
        private IReadOnlyDictionary<int, double> _modelRowDips;
        private IReadOnlyDictionary<int, double> _modelColumnDips;
        private double? _modelDefaultRowDip;
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
            return new Rect(rect.Left / _renderScaleX, rect.Top / _renderScaleY,
                Math.Max(0, (rect.Right - rect.Left) / _renderScaleX), Math.Max(0, (rect.Bottom - rect.Top) / _renderScaleY));
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
            // Workbook dimensions are DIP; the viewport works in final device
            // pixels, so the current render scale (monitor DPI x layout zoom) is
            // applied here and re-applied if that scale later changes.
            _modelRowDips = rows;
            _modelColumnDips = columns;
            _modelDefaultRowDip = defaultRowHeight;
            ApplyStoredDimensions();
            // NotifyRefresh runs before the host installs the new sheet's dimensions.
            // Recompute header geometry too, instead of retaining the previous sheet's height.
            RecalculateHeaderSize();
            AdjustScrollbars(); InvalidateAll(); NotifyViewportChanged();
        }
        private void ApplyStoredDimensions()
        {
            if (_modelRowDips is null || _modelColumnDips is null) return;
            if (_modelDefaultRowDip is { } defaultRowHeight)
                _rowSizes.DefaultSize = Math.Max(1, (int)Math.Round(defaultRowHeight * _renderScaleY));
            _rowSizes.ReplaceSizeOverrides(_modelRowDips.ToDictionary(p => p.Key, p => (int)Math.Round(p.Value * _renderScaleY)));
            _columnSizes.ReplaceSizeOverrides(_modelColumnDips.ToDictionary(p => p.Key, p => (int)Math.Round(p.Value * _renderScaleX)));
        }
        public FastGridCellAddress HitTestModel(Point point) => RealToModel(GetCellAddress(new Point(point.X * _renderScaleX, point.Y * _renderScaleY)));
        /// <summary>Visible row/column trailing header edge, in control DIP. No document writes.</summary>
        public FastGridCellAddress HitTestResizeHeader(Point point, double tolerance = 4)
        {
            if (_drawBuffer == null || !double.IsFinite(point.X) || !double.IsFinite(point.Y) ||
                !double.IsFinite(tolerance) || tolerance < 0) return FastGridCellAddress.Empty;
            point = TranslatePoint(point, image);
            var x = point.X * _renderScaleX;
            var y = point.Y * _renderScaleY;
            if (x < 0 || y < 0 || x >= _drawBuffer.PixelWidth || y >= _drawBuffer.PixelHeight)
                return FastGridCellAddress.Empty;
            var columnHeader = y < HeaderHeight && x > HeaderWidth;
            var rowHeader = x < HeaderWidth && y > HeaderHeight;
            if (!columnHeader && !rowHeader) return FastGridCellAddress.Empty;
            var sizes = rowHeader ? _rowSizes : _columnSizes;
            var coordinate = rowHeader ? y : x;
            var nearest = tolerance * (rowHeader ? _renderScaleY : _renderScaleX);
            var index = -1;
            foreach (var pane in VisibleAxis(sizes,
                rowHeader ? FirstVisibleRowScrollIndex : FirstVisibleColumnScrollIndex,
                rowHeader ? VisibleRowCount : VisibleColumnCount,
                rowHeader ? HeaderHeight : HeaderWidth,
                rowHeader ? _drawBuffer.PixelHeight : _drawBuffer.PixelWidth))
            {
                for (var real = pane.First; real <= pane.Last && real < sizes.RealCount; real++)
                {
                    var edge = (rowHeader ? GetRowTop(real) : GetColumnLeft(real)) + sizes.GetSizeByRealIndex(real);
                    // Clipped/offscreen edges must not resize an invisible item or a frozen neighbor.
                    if (edge <= pane.Min || edge > pane.Max) continue;
                    var distance = Math.Abs(coordinate - edge);
                    if (distance > nearest || index >= 0 && distance == nearest) continue;
                    nearest = distance;
                    index = real;
                }
            }
            return index < 0 ? FastGridCellAddress.Empty :
                RealToModel(rowHeader ? new FastGridCellAddress(index, null) : new FastGridCellAddress(null, index));
        }
        public IReadOnlyList<int> GetVisibleModelRows() => VisibleAxis(_rowSizes, FirstVisibleRowScrollIndex, VisibleRowCount, HeaderHeight, _drawBuffer?.PixelHeight ?? 0)
            .SelectMany(a => Enumerable.Range(a.First, Math.Max(0, Math.Min(a.Last + 1, _realRowCount) - a.First))).Select(_rowSizes.RealToModel).ToArray();
        public IReadOnlyList<int> GetVisibleModelColumns() => VisibleAxis(_columnSizes, FirstVisibleColumnScrollIndex, VisibleColumnCount, HeaderWidth, _drawBuffer?.PixelWidth ?? 0)
            .SelectMany(a => Enumerable.Range(a.First, Math.Max(0, Math.Min(a.Last + 1, _realColumnCount) - a.First))).Select(_columnSizes.RealToModel).ToArray();
        public double RowHeaderWidth => HeaderWidth / _renderScaleX;
        public double ColumnHeaderHeight => HeaderHeight / _renderScaleY;
        /// <summary>Offset of the raster viewport inside the control, in DIP.</summary>
        /// <remarks>
        /// The vector layers (artwork, selection, editor) draw from the control
        /// origin, so the bitmap must not be re-centred by layout: an explicit
        /// <c>Width</c>/<c>Height</c> with Stretch alignment shifts it by ~1 DIP.
        /// </remarks>
        public Point RasterOrigin => image.TranslatePoint(new Point(0, 0), this);
        public int FrozenRowCount => _rowSizes.FrozenCount;
        public int FrozenColumnCount => _columnSizes.FrozenCount;

        public void RestoreModelScroll(int row, int column)
        {
            var real = ModelToReal(new FastGridCellAddress(row, column));
            if (!real.IsCell) return;
            ScrollContent(Math.Max(0, real.Row.Value - _rowSizes.FrozenCount), Math.Max(0, real.Column.Value - _columnSizes.FrozenCount));
        }
        public FastGridCellAddress? GetScrollingModelCell()
        {
            if (_rowSizes.ScrollCount == 0 || _columnSizes.ScrollCount == 0) return null;
            return new FastGridCellAddress(
                _rowSizes.RealToModel(Math.Min(_rowSizes.RealCount - 1, FirstVisibleRowScrollIndex + _rowSizes.FrozenCount)),
                _columnSizes.RealToModel(Math.Min(_columnSizes.RealCount - 1, FirstVisibleColumnScrollIndex + _columnSizes.FrozenCount)));
        }

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
            var dx = _renderScaleX; var dy = _renderScaleY;
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
