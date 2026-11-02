using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace FastWpfGrid
{
    /// <summary>
    /// Interaction logic for FastGridControl.xaml
    /// </summary>
    public partial class FastGridControl : UserControl, IFastGridView
    {
        //private double _lastvscroll;
        private IFastGridModel _model;

        private FastGridCellAddress _currentCell;

        private int _headerHeight;
        private int _headerWidth;
        private Dictionary<Tuple<bool, bool>, GridFont> _glyphFonts = new Dictionary<Tuple<bool, bool>, GridFont>();
        private QuickExcel.Rendering.Skia.SkiaRenderer _textRenderer;
        private Dictionary<Color, Brush> _solidBrushes = new Dictionary<Color, Brush>();
        private int _rowHeightReserve = 5;
        //private Color _headerBackground = Color.FromRgb(0xDD, 0xDD, 0xDD);
        // Private software canvas; the compositor only receives immutable frame snapshots.
        private WriteableBitmap _drawBuffer;

        private bool _isTransposed;

        private bool _isReadOnly;

        private static Dictionary<string, ImageHolder> _imageCache = new Dictionary<string, ImageHolder>();

        public FastGridControl()
        {
            InitializeComponent();
            selectionOverlay.Grid = this;
            //gridCore.Grid = this;
            CellFontSize = 11;
            _dragTimer = new DispatcherTimer();
            _dragTimer.IsEnabled = false;
            _dragTimer.Interval = TimeSpan.FromSeconds(0.05);
            _dragTimer.Tick += _dragTimer_Tick;
            Loaded += (_, _) => { if (UpdateRenderScale()) RebuildViewport(); InvalidateAll(); };
            // A layout zoom can also reach us through an ancestor scale change
            // that resizes this control without a separate viewport event.
            SizeChanged += (_, _) => { if (UpdateRenderScale()) RebuildViewport(); };
            Unloaded += (_, _) => { StopDragScrolling(); CancelPendingHorizontalScroll(); _contentTiles.Clear(); _tileCanvas?.Freeze(); _tileCanvas = null; ClearCaches(); _textRenderer?.Dispose(); _textRenderer = null; };
            AllowSelectAll = true;
        }

        public bool AllowSelectAll { get; set; }

        /// <summary>
        /// Prevents the inline editor from being used if the control is in a read-only state.
        /// </summary>
        public bool IsReadOnly
        {
            get { return _isReadOnly; }
            set { _isReadOnly = value; }
        }

        public GridFont GetFont(bool isBold, bool isItalic)
        {
            var key = Tuple.Create(isBold, isItalic);
            if (!_glyphFonts.ContainsKey(key))
            {
                var font = new GridFont(_textRenderer ??= new(), new PortableFontDesc(CellFontName, CellFontSize, isBold, isItalic, false,
                    RenderScaleX, RenderScaleY));
                _glyphFonts[key] = font;
            }
            return _glyphFonts[key];
        }

        public void ClearCaches()
        {
            _glyphFonts.Clear();
        }

        public int GetTextWidth(string text, bool isBold, bool isItalic)
        {
            return GetFont(isBold, isItalic).GetTextWidth(text);
            //double size = CellFontSize;
            //int totalWidth = 0;
            //var glyphTypeface = GetFont(isBold, isItalic);

            //for (int n = 0; n < text.Length; n++)
            //{
            //    ushort glyphIndex = glyphTypeface.CharacterToGlyphMap[text[n]];
            //    double width = Math.Round(glyphTypeface.AdvanceWidths[glyphIndex] * size);
            //    totalWidth += (int)width;
            //}
            //return totalWidth;
        }

        private void RecalculateDefaultCellSize()
        {
            ClearCaches();
            int rowHeight = GetFont(false, false).TextHeight + CellPaddingVertical*2 + 2 + RowHeightReserve;
            int columnWidth = MinColumnWidthOverride ?? rowHeight*4;

            _rowSizes.DefaultSize = rowHeight;
            _columnSizes.DefaultSize = columnWidth;

            RecalculateHeaderSize();

            InvalidateAll();
        }

        private void RecalculateHeaderSize()
        {
            HeaderWidth = GetTextWidth("0000000", false, false);
            HeaderHeight = _rowSizes.DefaultSize;

            if (IsTransposed) CountTransposedHeaderWidth();
            if (Model != null)
            {
                int width = GetCellContentWidth(Model.GetGridHeader(this));
                if (width + 2 * CellPaddingHorizontal > HeaderWidth) HeaderWidth = width + 2 * CellPaddingHorizontal;
            }
        }

        private void CountTransposedHeaderWidth()
        {
            int maxw = 0;
            for (int col = 0; col < _modelColumnCount; col++)
            {
                var cell = Model.GetColumnHeader(this, col);
                int width = GetCellContentWidth(cell) + 2*CellPaddingHorizontal;
                if (width > maxw) maxw = width;
            }
            HeaderWidth = maxw;
        }

        //public int RowHeight
        //{
        //    get { return _rowHeight; }
        //}

        //public int ColumnWidth
        //{
        //    get { return _columnWidth; }
        //}

        private void ScrollChanged()
        {
            int rowIndex = (int) Math.Round(vscroll.Value);
            if (RetainedContentEnabled)
            {
                // The horizontal thumb measures device pixels, including the
                // partial leading column. Vertical input preserves that offset.
                var horizontalPosition = hscroll.Value;
                ScrollContent(rowIndex, FirstVisibleColumnScrollIndex);
                ScrollHorizontalPixels(horizontalPosition);
            }
            else ScrollContent(rowIndex, (int)Math.Round(hscroll.Value));
            AdjustInlineEditorPosition();
            AdjustSelectionMenuPosition();
        }


        public Color GetAlternateBackground(int row)
        {
            return _alternatingColors[row%_alternatingColors.Length];
        }

        private void hscroll_Scroll(object sender, ScrollEventArgs e)
        {
            if (RetainedContentEnabled && e.ScrollEventType == ScrollEventType.ThumbTrack)
            {
                QueueHorizontalScroll(e.NewValue);
                return;
            }
            // Release/end/line/page events commit the latest thumb value now.
            CancelPendingHorizontalScroll();
            ScrollChanged();
        }

        private void vscroll_Scroll(object sender, ScrollEventArgs e)
        {
            ScrollChanged();
        }

        private void OnModelPropertyChanged()
        {
            _hasModelDimensions = false;
            if (_model != null) _model.DetachView(this);
            _model = Model;
            if (_model != null) _model.AttachView(this);
            NotifyRefresh();
        }


        public int GridScrollAreaWidth
        {
            get
            {
                if (_drawBuffer == null) return 1;
                return _drawBuffer.PixelWidth - HeaderWidth - FrozenWidth;
            }
        }

        public int GridScrollAreaHeight
        {
            get
            {
                if (_drawBuffer == null) return 1;
                return _drawBuffer.PixelHeight - HeaderHeight - FrozenHeight;
            }
        }

        private void AdjustScrollbars()
        {
            CancelPendingHorizontalScroll();
            hscroll.Minimum = 0;
            hscroll.Maximum = RetainedContentEnabled ? MaximumHorizontalPixelScroll : _columnSizes.GetMaximumScrollIndex(GridScrollAreaWidth);
            hscroll.SmallChange = RetainedContentEnabled ? Math.Max(1, _columnSizes.DefaultSize) : 1;
            hscroll.LargeChange = RetainedContentEnabled ? Math.Max(1, GridScrollAreaWidth) : 3;

            vscroll.Minimum = 0;
            // Imported/custom row heights must participate after filtering, just as column widths do.
            vscroll.Maximum = FlexibleRows ? Math.Max(0, _rowSizes.ScrollCount - 1) : _rowSizes.GetMaximumScrollIndex(GridScrollAreaHeight);
            vscroll.SmallChange = 1; // _rowSizes.DefaultSize;
            vscroll.LargeChange = 10; // GridScrollAreaHeight / 2.0;

            // WPF coerces the thumb value when Maximum shrinks, but not our viewport indices.
            // Clamp before measuring the viewport, otherwise filtering near the bottom reports zero visible rows.
            ScrollContent(Math.Clamp(FirstVisibleRowScrollIndex, 0, (int)vscroll.Maximum),
                Math.Clamp(FirstVisibleColumnScrollIndex, 0, _columnSizes.GetMaximumScrollIndex(GridScrollAreaWidth)));
            if (RetainedContentEnabled) ScrollHorizontalPixels(_columnSizes.GetPositionByScrollIndex(FirstVisibleColumnScrollIndex) + _scrollOffsetX);
            hscroll.ViewportSize = RetainedContentEnabled ? Math.Max(0, GridScrollAreaWidth) : VisibleColumnCount;
            vscroll.ViewportSize = VisibleRowCount;
            AdjustScrollBarPositions();
        }

        private void AdjustScrollBarPositions()
        {
            hscroll.Value = RetainedContentEnabled ? _columnSizes.GetPositionByScrollIndex(FirstVisibleColumnScrollIndex) + _scrollOffsetX : FirstVisibleColumnScrollIndex;
            vscroll.Value = FirstVisibleRowScrollIndex;
        }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);
            AdjustScrollbars();
        }

        public void NotifyRefresh()
        {
            _modelRowCount = 0;
            _modelColumnCount = 0;
            if (_model != null)
            {
                _modelRowCount = _model.RowCount;
                _modelColumnCount = _model.ColumnCount;
            }

            UpdateSeriesCounts();
            RecalculateHeaderSize();
            FixCurrentCellAndSetSelectionToCurrentCell();

            RecountColumnWidths();
            RecountRowHeights();
            AdjustScrollbars();
            SetScrollbarMargin();
            FixScrollPosition();
            InvalidateAll();
        }

        private void FixCurrentCellAndSetSelectionToCurrentCell()
        {
            int? col = _currentCell.Column;
            int? row = _currentCell.Row;

            if (col.HasValue)
            {
                if (col >= _realColumnCount) col = _realColumnCount - 1;
                if (col < 0) col = null;
            }

            if (row.HasValue)
            {
                if (row >= _realRowCount) row = _realRowCount - 1;
                if (row < 0) row = null;
            }

            ClearSelectedCells();
            SetCurrentCell(new FastGridCellAddress(row, col));
            if (_currentCell.IsCell) AddSelectedCell(_currentCell);
            OnChangeSelectedCells(false);
        }

        public void NotifyAddedRows()
        {
            NotifyRefresh();
        }

        public Brush GetSolidBrush(Color color)
        {
            if (!_solidBrushes.ContainsKey(color))
            {
                _solidBrushes[color] = new SolidColorBrush(color);
            }
            return _solidBrushes[color];
        }


        private IFastGridCell GetModelRowHeader(int row)
        {
            if (_model == null) return null;
            if (row < 0 || row >= _modelRowCount) return null;
            return _model.GetRowHeader(this, row);
        }

        private IFastGridCell GetModelColumnHeader(int col)
        {
            if (_model == null) return null;
            if (col < 0 || col >= _modelColumnCount) return null;
            return _model.GetColumnHeader(this, col);
        }

        private IFastGridCell GetModelCell(int row, int col)
        {
            if (_model == null) return null;
            if (row < 0 || row >= _modelRowCount) return null;
            if (col < 0 || col >= _modelColumnCount) return null;
            return _model.GetCell(this, row, col);
        }

        private IFastGridCell GetColumnHeader(int col)
        {
            if (IsTransposed) return GetModelRowHeader(_columnSizes.RealToModel(col));
            return GetModelColumnHeader(_columnSizes.RealToModel(col));
        }

        private IFastGridCell GetRowHeader(int row)
        {
            if (IsTransposed) return GetModelColumnHeader(_rowSizes.RealToModel(row));
            return GetModelRowHeader(_rowSizes.RealToModel(row));
        }

        private IFastGridCell GetCell(int row, int col)
        {
            if (IsTransposed) return GetModelCell(_columnSizes.RealToModel(col), _rowSizes.RealToModel(row));
            return GetModelCell(_rowSizes.RealToModel(row), _columnSizes.RealToModel(col));
        }

        private IFastGridCell GetCell(FastGridCellAddress addr)
        {
            if (addr.IsCell) return GetCell(addr.Row.Value, addr.Column.Value);
            if (addr.IsRowHeader) return GetRowHeader(addr.Row.Value);
            if (addr.IsColumnHeader) return GetColumnHeader(addr.Column.Value);
            if (addr.IsGridHeader && _model != null) return _model.GetGridHeader(this);
            return null;
        }

        public void HideInlineEditor(bool saveCellValue = true)
        {
            using (var ctx = CreateInvalidationContext())
            {
                if (saveCellValue && _inplaceEditorCell.IsCell && _inlineTextChanged)
                {
                    var cell = GetCell(_inplaceEditorCell.Row.Value, _inplaceEditorCell.Column.Value);
                    cell.SetEditText(edText.Text);
                    InvalidateCell(_inplaceEditorCell);
                }
                _inplaceEditorCell = new FastGridCellAddress();
                edText.Text = "";
                edText.Visibility = Visibility.Hidden;
            }
            Keyboard.Focus(image);
        }

        private void ShowInlineEditor(FastGridCellAddress cell, string textValueOverride = null)
        {
            if (_isReadOnly) return;
            if (!cell.IsCell) return;
            var cellObj = GetCell(cell.Row.Value, cell.Column.Value);
            if (cellObj == null) return;
            string text = cellObj.GetEditText();
            if (text == null) return;

            _inplaceEditorCell = cell;

            edText.Text = textValueOverride ?? text;
            edText.Visibility = Visibility.Visible;
            AdjustInlineEditorPosition();

            if (edText.IsFocused)
            {
                if (textValueOverride == null)
                {
                    edText.SelectAll();
                }
            }
            else
            {
                edText.Focus();
                if (textValueOverride == null)
                {
                    Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.Input, (Action) edText.SelectAll);
                }
            }

            if (textValueOverride != null)
            {
                edText.SelectionStart = textValueOverride.Length;
            }

            _inlineTextChanged = !String.IsNullOrEmpty(textValueOverride);
        }

        private void AdjustInlineEditorPosition()
        {
            if (_inplaceEditorCell.IsCell)
            {
                bool visible = _rowSizes.IsVisible(_inplaceEditorCell.Row.Value, FirstVisibleRowScrollIndex, GridScrollAreaHeight)
                               && _columnSizes.IsVisible(_inplaceEditorCell.Column.Value, FirstVisibleColumnScrollIndex, GridScrollAreaWidth);
                edText.Visibility = visible ? Visibility.Visible : Visibility.Hidden;
                var rect = GetCellRect(_inplaceEditorCell.Row.Value, _inplaceEditorCell.Column.Value);

                edText.Margin = new Thickness
                    {
                        Left = rect.Left / _renderScaleX,
                        Top = rect.Top / _renderScaleY,
                        Right = imageGrid.ActualWidth - rect.Right / _renderScaleX,
                        Bottom = imageGrid.ActualHeight - rect.Bottom / _renderScaleY,
                    };
            }
        }

        private void AdjustSelectionMenuPosition()
        {
            FastGridCellAddress maxaddr = FastGridCellAddress.Empty;

            if (mnuSelection.Visibility != Visibility.Visible) return;
            foreach (var range in _selectedCells.Ranges)
            {
                var addr = new FastGridCellAddress(range.Bottom, range.Right);
                if (!maxaddr.IsCell || addr.Row.Value + addr.Column.Value > maxaddr.Row.Value + maxaddr.Column.Value) maxaddr = addr;
            }

            if (!maxaddr.IsCell) return;

            int left = GetColumnLeft(maxaddr.Column.Value);
            int top = GetRowTop(maxaddr.Row.Value + 1);

            mnuSelection.Margin = new Thickness
            {
                Left = left / _renderScaleX,
                Top = top / _renderScaleY,
            };
        }

        private void InvalidateCurrentCell()
        {
            if (UseSelectionOverlay) { selectionOverlay.Refresh(); return; }
            if (_currentCell.IsCell) InvalidateCell(_currentCell);
            if (_currentCell.Column.HasValue) InvalidateColumnHeader(_currentCell.Column.Value);
            if (_currentCell.Row.HasValue) InvalidateRowHeader(_currentCell.Row.Value);
        }

        private void SetCurrentCell(FastGridCellAddress cell)
        {
            if (cell.IsRowHeader && _currentCell.IsCell) cell = new FastGridCellAddress(cell.Row, _currentCell.Column);
            if (cell.IsColumnHeader && _currentCell.IsCell) cell = new FastGridCellAddress(_currentCell.Row, cell.Column);
            if (cell.IsCell)
            {
                var merge = GetMergedRealCellRange(cell);
                cell = new FastGridCellAddress(merge.Top, merge.Left);
            }
            if (cell == _currentCell) return;

            using (var ctx = CreateInvalidationContext())
            {
                InvalidateCurrentCell();
                _currentCell = cell;
                InvalidateCurrentCell();
            }
        }

        /// returns cell range. 
        /// if a is row header, returns full rows.
        /// if a is column header, returns full columns
        private GridRangeSelection GetCellRange(FastGridCellAddress a, FastGridCellAddress b)
        {
            var res = new GridRangeSelection();

            int minrow;
            int maxrow;
            int mincol;
            int maxcol;

            if (a.IsRowHeader)
            {
                mincol = 0;
                maxcol = _columnSizes.RealCount - 1;
            }
            else
            {
                if (a.Column == null || b.Column == null) return res;
                mincol = Math.Min(a.Column.Value, b.Column.Value);
                maxcol = Math.Max(a.Column.Value, b.Column.Value);
            }

            if (a.IsColumnHeader)
            {
                minrow = 0;
                maxrow = _rowSizes.RealCount - 1;
            }
            else
            {
                if (a.Row == null || b.Row == null) return res;
                minrow = Math.Min(a.Row.Value, b.Row.Value);
                maxrow = Math.Max(a.Row.Value, b.Row.Value);
            }

            const int LIMIT_RESERVE = 3;
            if (SelectedRealRowCountLimit.HasValue && maxrow - minrow > SelectedRealRowCountLimit.Value + LIMIT_RESERVE)
            {
                maxrow = minrow + SelectedRealRowCountLimit.Value + LIMIT_RESERVE;
            }
            if (SelectedRealColumnCountLimit.HasValue && maxcol - mincol > SelectedRealColumnCountLimit.Value + LIMIT_RESERVE)
            {
                maxcol = mincol + SelectedRealColumnCountLimit.Value + LIMIT_RESERVE;
            }

            if (minrow <= maxrow && mincol <= maxcol) res.Set(ExpandMergedRange(new GridCellRange(minrow, mincol, maxrow, maxcol)));
            return res;
        }


        private void SetHoverRow(int? row)
        {
            if (row == _mouseOverRow) return;
            if (UseSelectionOverlay) { _mouseOverRow = row; selectionOverlay.Refresh(); return; }
            using (var ctx = CreateInvalidationContext())
            {
                if (_mouseOverRow.HasValue) InvalidateRow(_mouseOverRow.Value);
                _mouseOverRow = row;
                if (_mouseOverRow.HasValue) InvalidateRow(_mouseOverRow.Value);
            }
        }

        private void SetHoverRowHeader(int? row)
        {
            if (row == _mouseOverRowHeader) return;
            if (UseSelectionOverlay) { _mouseOverRowHeader = row; selectionOverlay.Refresh(); return; }
            using (var ctx = CreateInvalidationContext())
            {
                if (_mouseOverRowHeader.HasValue) InvalidateRowHeader(_mouseOverRowHeader.Value);
                _mouseOverRowHeader = row;
                if (_mouseOverRowHeader.HasValue) InvalidateRow(_mouseOverRowHeader.Value);
            }
        }

        private void SetHoverColumnHeader(int? column)
        {
            if (column == _mouseOverColumnHeader) return;
            if (UseSelectionOverlay) { _mouseOverColumnHeader = column; selectionOverlay.Refresh(); return; }
            using (var ctx = CreateInvalidationContext())
            {
                if (_mouseOverColumnHeader.HasValue) InvalidateColumnHeader(_mouseOverColumnHeader.Value);
                _mouseOverColumnHeader = column;
                if (_mouseOverColumnHeader.HasValue) InvalidateColumn(_mouseOverColumnHeader.Value);
            }
        }

        private void SetHoverCell(FastGridCellAddress cell)
        {
            if (cell == _mouseOverCell) return;
            if (UseSelectionOverlay && IsSelectionDragging)
            {
                _mouseOverCell = cell;
                CurrentCellActiveRegions.Clear(); CurrentHoverRegion = null;
                return;
            }
            using (var ctx = CreateInvalidationContext())
            {
                if (!_mouseOverCell.IsEmpty) InvalidateCell(_mouseOverCell);
                _mouseOverCell = cell.IsEmpty ? FastGridCellAddress.Empty : cell;
                if (!_mouseOverCell.IsEmpty) InvalidateCell(_mouseOverCell);
            }
        }

        private void imageGridResized(object sender, SizeChangedEventArgs e)
        {
            // Layout zoom and monitor DPI change the raster density before any
            // geometry is measured, so the viewport is rebuilt at the new scale.
            UpdateRenderScale();
            RebuildViewport();
        }

        private void RebuildViewport()
        {
            bool wasEmpty = _drawBuffer == null;
            int width = (int) imageGrid.ActualWidth - 2;
            int height = (int) imageGrid.ActualHeight - 2;
            if (width > 0 && height > 0)
            {
                //To avoid flicker (blank image) while resizing, crop the current buffer and set it as the image source instead of using a new one.
                //This will be shown during the refresh.
                int pixelWidth = (int) Math.Ceiling(width * _renderScaleX);
                int pixelHeight = (int) Math.Ceiling(height * _renderScaleY);
                if (_drawBuffer == null || Math.Abs(_drawBuffer.DpiX - _renderScaleX * 96) > 0.01 || Math.Abs(_drawBuffer.DpiY - _renderScaleY * 96) > 0.01)
                {
                    var previous = _drawBuffer;
                    _drawBuffer = BitmapFactory.New(pixelWidth, pixelHeight, _renderScaleX, _renderScaleY);
                    //The unmanaged memory when crating new WritableBitmaps doesn't reliably garbage collect and can still cause out of memory exceptions
                    //Profiling revealed handles on the object that aren't able to be collected.
                    //Freezing the object removes all handles and should help in garbage collection.
                    previous?.Freeze();
                }
                else
                {
                    var oldBuffer = _drawBuffer;
                    _drawBuffer = CropBuffer(oldBuffer, pixelWidth, pixelHeight);

                    //The unmanaged memory when crating new WritableBitmaps doesn't reliably garbage collect and can still cause out of memory exceptions
                    //Profiling revealed handles on the object that aren't able to be collected.
                    //Freezing the object removes all handles and should help in garbage collection.
                    oldBuffer.Freeze();
                }
            }
            else
            {
                _drawBuffer = null;
            }
            if (_drawBuffer == null) image.Source = null;
            image.Margin = new Thickness(0);
            //Present the bitmap at its natural size. Stretching the viewport to a
            //rounded DIP width resamples every rasterized glyph.
            if (_drawBuffer != null)
            {
                image.Width = _drawBuffer.PixelWidth / _renderScaleX;
                image.Height = _drawBuffer.PixelHeight / _renderScaleY;
            }
            else
            {
                image.Width = Math.Max(0, width);
                image.Height = Math.Max(0, height);
            }
            selectionOverlay.Width = imageGrid.ActualWidth; selectionOverlay.Height = imageGrid.ActualHeight;

            //var screenPos = imageGrid.PointToScreen(new Point(0, 0));
            //double fracX = screenPos.X - Math.Truncate(screenPos.X);
            //double fracY = screenPos.Y - Math.Truncate(screenPos.Y);
            //double dleft = 1 - fracX;
            //double dtop = 1 - fracY;
            //if (fracX == 0) dleft = 0;
            //if (fracY == 0) dtop = 0;
            //image.Margin = new Thickness(dleft, dtop, imageGrid.ActualWidth - width - dleft - 1, imageGrid.ActualHeight - height - dtop - 25);

            // A host can bind workbook dimensions before the first layout (or while hidden).
            // Creating the raster viewport must not replace those dimensions with text auto-fit.
            if (wasEmpty && _drawBuffer != null && !_hasModelDimensions)
            {
                RecountColumnWidths();
                RecountRowHeights();
            }
            AdjustScrollbars();
            InvalidateAll();
            NotifyViewportChanged();
        }

        private bool MoveCurrentCell(int? row, int? col, KeyEventArgs e = null)
        {
            if (e != null) e.Handled = true;
            if (!ShiftPressed)
            {
                InvalidateSelection();
                ClearSelectedCells();
            }

            InvalidateCurrentCell();

            if (row < 0) row = 0;
            if (row >= _realRowCount) row = _realRowCount - 1;
            if (col < 0) col = 0;
            if (col >= _realColumnCount) col = _realColumnCount - 1;

            SetCurrentCell(new FastGridCellAddress(row, col));
            if (_currentCell.IsCell) AddSelectedCell(_currentCell);
            InvalidateCurrentCell();
            ScrollCurrentCellIntoView();
            OnChangeSelectedCells(true);
            return true;
        }

        private void RenderChanged()
        {
            InvalidateAll();
        }

        private void OnUseClearTypePropertyChanged()
        {
            ClearCaches();
            RecalculateDefaultCellSize();
            RenderChanged();
        }

        private bool IsModelCellInValidRange(FastGridCellAddress cell)
        {
            if (cell.Row.HasValue && (cell.Row.Value < 0 || cell.Row.Value >= _modelRowCount)) return false;
            if (cell.Column.HasValue && (cell.Column.Value < 0 || cell.Column.Value >= _modelColumnCount)) return false;
            return true;
        }

        public HashSet<FastGridCellAddress> GetSelectedModelCells()
        {
            var res = new HashSet<FastGridCellAddress>();
            foreach (var cell in _selectedCells)
            {
                var cellModel = RealToModel(cell);
                if (cellModel.IsCell && IsModelCellInValidRange(cellModel)) res.Add(cellModel);
            }
            return res;
        }

        public FastGridCellAddress CurrentModelCell
        {
            get
            {
                var cell = RealToModel(CurrentCell);
                return IsModelCellInValidRange(cell) ? cell : FastGridCellAddress.Empty;
            }
            set { CurrentCell = ModelToReal(value); }
        }

        public void ShowSelectionMenu(IEnumerable<string> commands)
        {
            if (commands == null)
            {
                mnuSelection.ItemsSource = null;
                mnuSelection.Visibility = Visibility.Hidden;
            }
            else
            {
                mnuSelection.ItemsSource = commands.Select(x => new SelectionQuickCommand(Model, x)).ToList();
                mnuSelection.Visibility = Visibility.Visible;
                AdjustSelectionMenuPosition();
            }
        }
    }
}
