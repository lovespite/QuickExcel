// Derived from FastWpfGrid (MIT), Jan Prochazka. QuickExcel range selection.
using System;
using System.Collections.Generic;
using System.Linq;
namespace FastWpfGrid
{
    partial class FastGridControl
    {
        public event EventHandler<SelectionChangedEventArgs> SelectedCellsChanged;
        private readonly GridRangeSelection _selectedCells = new();
        public IReadOnlyList<GridCellRange> SelectedRanges => _selectedCells.Ranges;
        public long SelectedCellCount => _selectedCells.Count;
        int? _selectedRealRowCountLimit;
        bool _selectedRealRowCountLimitLoaded;
        public int? SelectedRealRowCountLimit
        {
            get
            {
                if (!_selectedRealRowCountLimitLoaded)
                { _selectedRealRowCountLimitLoaded = true; _selectedRealRowCountLimit = Model == null ? null : IsTransposed ? Model.SelectedColumnCountLimit : Model.SelectedRowCountLimit; }
                return _selectedRealRowCountLimit;
            }
        }
        int? _selectedRealColumnCountLimit;
        bool _selectedRealColumnCountLimitLoaded;
        public int? SelectedRealColumnCountLimit
        {
            get
            {
                if (!_selectedRealColumnCountLimitLoaded)
                { _selectedRealColumnCountLimitLoaded = true; _selectedRealColumnCountLimit = Model == null ? null : IsTransposed ? Model.SelectedRowCountLimit : Model.SelectedColumnCountLimit; }
                return _selectedRealColumnCountLimit;
            }
        }
        private void ClearSelectedCells() => _selectedCells.Clear();
        private void AddSelectedCell(FastGridCellAddress cell) => _selectedCells.Add(cell);
        private void RemoveSelectedCell(FastGridCellAddress cell) => _selectedCells.Remove(cell);
        public bool IsLimitedSelection => SelectedRealRowCountLimit.HasValue && _selectedCells.RowCount >= SelectedRealRowCountLimit.Value ||
            SelectedRealColumnCountLimit.HasValue && _selectedCells.ColumnCount >= SelectedRealColumnCountLimit.Value;
        private void SetSelectedRectangle(FastGridCellAddress origin, FastGridCellAddress cell)
        {
            var next = GetCellRange(origin, cell);
            if (_selectedCells.Ranges.SequenceEqual(next.Ranges))
            {
                var previous = _currentCell;
                SetCurrentCell(cell);
                if (previous != _currentCell) OnChangeSelectedCells(true);
                return;
            }
            if (UseSelectionOverlay)
            {
                _selectedCells.Set(next.Ranges);
                SetCurrentCell(cell);
                OnChangeSelectedCells(true);
                return;
            }
            // Keep the selection actually painted at the start of this dispatcher frame.
            // Repaint only membership changes; old/new current cells and headers are invalidated separately.
            var wasLimited = IsLimitedSelection;
            if (_selectionBeforeRender == null)
            {
                _selectionBeforeRender = new GridRangeSelection();
                _selectionBeforeRender.Set(_selectedCells.Ranges);
            }
            _selectedCells.Set(next.Ranges);
            SetCurrentCell(cell);
            _isInvalidated = true; CheckInvalidation();
            if (wasLimited != IsLimitedSelection) InvalidateAll();
            OnChangeSelectedCells(true);
        }
        public bool SelectModelRange(GridCellRange range, FastGridCellAddress current)
        {
            var rows = _rowSizes.GetRealRanges(IsTransposed ? range.Left : range.Top, IsTransposed ? range.Right : range.Bottom);
            var columns = _columnSizes.GetRealRanges(IsTransposed ? range.Top : range.Left, IsTransposed ? range.Bottom : range.Right);
            if (rows.Count == 0 || columns.Count == 0) return false;
            _selectedCells.Clear();
            foreach (var row in rows) foreach (var column in columns)
                _selectedCells.Add(new GridCellRange(row.First, column.First, row.Last, column.Last));
            var mapped = ModelToReal(current);
            if (!mapped.IsCell || !_selectedCells.Contains(mapped)) mapped = new(rows[0].First, columns[0].First);
            SetCurrentCell(mapped); InvalidateSelection(); OnChangeSelectedCells(true); ScrollCurrentCellIntoView();
            return true;
        }
        private void OnChangeSelectedCells(bool isInvokedByUser)
        {
            if (UseSelectionOverlay) selectionOverlay.Refresh();
            SelectedCellsChanged?.Invoke(this, new SelectionChangedEventArgs { IsInvokedByUser = isInvokedByUser });
        }
    }
}
