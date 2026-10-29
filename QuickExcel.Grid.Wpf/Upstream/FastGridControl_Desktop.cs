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
        public event EventHandler ViewportChanged;
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
        public void SetModelDimensions(IReadOnlyDictionary<int, double> rows, IReadOnlyDictionary<int, double> columns)
        {
            _rowSizes.ReplaceSizeOverrides(rows.ToDictionary(p => p.Key, p => (int)Math.Round(p.Value * DpiDetector.DpiYKoef)));
            _columnSizes.ReplaceSizeOverrides(columns.ToDictionary(p => p.Key, p => (int)Math.Round(p.Value * DpiDetector.DpiXKoef)));
            AdjustScrollbars(); InvalidateAll();
        }
        private void NotifyViewportChanged() => ViewportChanged?.Invoke(this, EventArgs.Empty);
    }
}
