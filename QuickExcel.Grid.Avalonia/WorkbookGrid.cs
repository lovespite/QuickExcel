using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using FastWpfGrid;
using QuickExcel.Core;
using QuickExcel.GridGeometry;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Rendering;
using QuickExcel.Rendering.Skia;
using SkiaSharp;

namespace QuickExcel.Grid.Avalonia;

/// <summary>A bounded visible-cell canvas, one native editor, and independent selection overlay.</summary>
public sealed class WorkbookGrid : UserControl, IGridSurface, IDisposable
{
    private const double HeaderWidth = 52, HeaderHeight = 28;
    private readonly ViewportAxis _rows = new(), _columns = new();
    private readonly Canvas _canvas = new() { ClipToBounds = true };
    private readonly GridArtwork _artwork;
    private readonly SelectionOverlay _overlay;
    private readonly ScrollBar _vertical = new() { Orientation = global::Avalonia.Layout.Orientation.Vertical, Width = 16 };
    private readonly ScrollBar _horizontal = new() { Orientation = global::Avalonia.Layout.Orientation.Horizontal, Height = 16 };
    private readonly SkiaRenderer _renderer = new();
    private WorkbookSnapshot? _snapshot;
    private int _sheetIndex;
    private CellPosition _current, _anchor, _draftOwner;
    private GridCellRange[] _merges = [];
    private readonly GridRangeSelection _selection = new();
    private bool _settingText, _editing, _scrolling, _dragging, _disposed;
    private double _zoom = 1;
    public CompositionTextBox CellEditor { get; } = new() { Padding = new(2, 0), MinHeight = 0, MinWidth = 0,
        BorderThickness = new(1), BorderBrush = Brushes.Teal, Background = Brushes.White, Opacity = 0 };
    public Func<int, CellPosition, string, bool>? CommitCell { get; set; }
    public Func<bool>? BeforeSelection { get; set; }
    public Func<WorkbookSnapshot, WorksheetSnapshot, int, int, StyleDefinition>? ResolveStyle { get; set; }
    public event Action? SelectionChanged;
    public event Action? ViewportChanged;
    public CellPosition Current => _current;
    public bool IsEditing => _editing;
    public bool IsComposing => CellEditor.IsComposing;
    public double Zoom => _zoom;
    public WorksheetSnapshot Sheet => _snapshot!.Sheets[_sheetIndex];
    public IReadOnlyList<CellPosition> Selection => _selection.Select(c => new CellPosition(c.Row!.Value, c.Column!.Value)).ToArray();
    public IReadOnlyList<GridRange> SelectionRanges => _selection.Ranges.Select(r => new GridRange(r.Top, r.Left, r.Bottom, r.Right)).ToArray();
    private double ViewWidth => Math.Max(0, _canvas.Bounds.Width / _zoom - HeaderWidth);
    private double ViewHeight => Math.Max(0, _canvas.Bounds.Height / _zoom - HeaderHeight);

    public WorkbookGrid()
    {
        Focusable = true;
        _artwork = new(this); _overlay = new(this) { IsHitTestVisible = false };
        _canvas.Children.Add(_artwork); _canvas.Children.Add(_overlay); _canvas.Children.Add(CellEditor);
        var layout = new global::Avalonia.Controls.Grid { RowDefinitions = new("*,16"), ColumnDefinitions = new("*,16") };
        layout.Children.Add(_canvas); layout.Children.Add(_vertical); layout.Children.Add(_horizontal);
        global::Avalonia.Controls.Grid.SetColumn(_vertical, 1); global::Avalonia.Controls.Grid.SetRow(_horizontal, 1);
        Content = layout;
        _canvas.SizeChanged += (_, _) => { _artwork.Width = _overlay.Width = _canvas.Bounds.Width; _artwork.Height = _overlay.Height = _canvas.Bounds.Height; UpdateScrollbars(); InvalidateContent(); PositionEditor(); };
        _vertical.ValueChanged += (_, _) => { if (!_scrolling) Scroll(_columns.Offset, _vertical.Value); };
        _horizontal.ValueChanged += (_, _) => { if (!_scrolling) Scroll(_horizontal.Value, _rows.Offset); };
        _canvas.PointerPressed += PointerDown; _canvas.PointerMoved += PointerMove;
        _canvas.PointerReleased += (_, e) => { _dragging = false; e.Pointer.Capture(null); PrepareEditor(true); };
        _canvas.PointerCaptureLost += (_, _) => _dragging = false;
        _canvas.PointerWheelChanged += (_, e) =>
        {
            if (!Commit()) return;
            Scroll(_columns.Offset - (e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? e.Delta.Y : e.Delta.X) * 60, _rows.Offset - (e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? 0 : e.Delta.Y) * 60);
            e.Handled = true;
        };
        CellEditor.PropertyChanged += (_, e) => { if (e.Property == TextBox.TextProperty && !_settingText) { _editing = true; CellEditor.Opacity = 1; } };
        CellEditor.CompositionChanged += () => { if (IsComposing) { _editing = true; CellEditor.Opacity = 1; } };
        AddHandler(KeyDownEvent, HandleKey, RoutingStrategies.Tunnel);
        DetachedFromVisualTree += (_, _) => _artwork.Release();
    }
    public void Bind(WorkbookSnapshot snapshot, int sheetIndex, int rows, int columns)
    {
        Cancel(); _snapshot = snapshot; _sheetIndex = sheetIndex; _zoom = Sheet.View.Zoom / 100d;
        Configure(rows, columns); _rows.ScrollTo(0, ViewHeight); _columns.ScrollTo(0, ViewWidth);
        var stored = Sheet.View.HasFreeze ? Sheet.View.PaneTopLeftCell : Sheet.View.TopLeftCell;
        if (_rows.Cell(stored.Row) is { Frozen: false } row) _rows.ScrollTo(row.Start - _rows.FrozenSize, ViewHeight);
        if (_columns.Cell(stored.Column) is { Frozen: false } col) _columns.ScrollTo(col.Start - _columns.FrozenSize, ViewWidth);
        SelectCore(Sheet.View.ActiveCell, false, false); InvalidateContent(); UpdateScrollbars();
    }
    public void Refresh(WorkbookSnapshot snapshot)
    {
        _snapshot = snapshot; _zoom = Sheet.View.Zoom / 100d; Configure(_rows.Count, _columns.Count);
        UpdateScrollbars(); InvalidateContent(); if (!_editing) PrepareEditor(false);
    }
    private void Configure(int rows, int columns)
    {
        var view = new SheetViewState(); view.Synchronize(_snapshot!, Sheet.Id, _current); view.RefreshFilter(_snapshot!);
        var hiddenRows = Sheet.Rows.Where(p => p.Value.Hidden).Select(p => p.Key).ToHashSet();
        if (view.FilterRange is { } filter)
            for (var r = filter.Top; r <= filter.Bottom; r++) if (view.IsFilteredOut(_snapshot!, _sheetIndex, r)) hiddenRows.Add(r);
        _rows.Configure(rows, Math.Clamp((int)(Sheet.DefaultRowHeight ?? 31), 1, 10000),
            Sheet.Rows.Where(p => p.Value.Height.HasValue).Select(p => new KeyValuePair<int, int>(p.Key, (int)Math.Clamp(p.Value.Height!.Value, 1, 10000))),
            hiddenRows, Enumerable.Range(0, Math.Min(rows, Sheet.View.FrozenRows)).ToHashSet());
        _columns.Configure(columns, (int)Math.Clamp(Sheet.Data.DefaultColumnWidth is { } width ? Math.Floor(width * 7 + 5) : 112, 1, 10000),
            Sheet.Columns.Keys.Append(0).Distinct().Select(c => new KeyValuePair<int, int>(c, (int)Math.Clamp(DrawingGeometry.ColumnSize(Sheet, c), 1, 10000))),
            Sheet.Columns.Where(p => p.Value.Hidden).Select(p => p.Key).ToHashSet(), Enumerable.Range(0, Math.Min(columns, Sheet.View.FrozenColumns)).ToHashSet());
        _merges = Sheet.MergedCells.Select(CellRectangle.Parse).Select(m => new GridCellRange(m.Top, m.Left, m.Bottom, m.Right)).ToArray();
    }
    public void Select(CellPosition position, bool extend = false)
    { if (Commit() && (BeforeSelection?.Invoke() ?? true)) SelectCore(position, extend, true); }
    private void SelectCore(CellPosition position, bool extend, bool focus)
    {
        var row = Math.Clamp(position.Row, 0, _rows.Count - 1); var col = Math.Clamp(position.Column, 0, _columns.Count - 1);
        if (_rows.Cell(row) is null) row = _rows.Visible(ViewHeight).FirstOrDefault().ModelIndex;
        if (_columns.Cell(col) is null) col = _columns.Visible(ViewWidth).FirstOrDefault().ModelIndex;
        var merge = _merges.FirstOrDefault(m => m.Contains(row, col));
        _current = merge == default ? new(row, col) : new(merge.Top, merge.Left);
        if (!extend) _anchor = _current;
        var range = GridRange.Between(_anchor, _current);
        _selection.Set(GridMergedSelection.Expand(new(range.Top, range.Left, range.Bottom, range.Right), _merges));
        var offsets = (_rows.Offset, _columns.Offset);
        _rows.EnsureVisible(_current.Row, ViewHeight); _columns.EnsureVisible(_current.Column, ViewWidth);
        UpdateScrollbars();
        if (offsets != (_rows.Offset, _columns.Offset)) InvalidateContent(); else _overlay.InvalidateVisual();
        PrepareEditor(focus); SelectionChanged?.Invoke();
    }
    public void SelectRange(GridRange range, CellPosition current)
    {
        if (!Commit() || !(BeforeSelection?.Invoke() ?? true)) return;
        _selection.Set(GridMergedSelection.Expand(new(range.Top, range.Left, range.Bottom, range.Right), _merges));
        _current = current; _anchor = new(range.Top, range.Left); _overlay.InvalidateVisual(); PrepareEditor(false); SelectionChanged?.Invoke();
    }
    public bool Commit()
    {
        if (IsComposing) return false;
        if (!_editing) return true;
        if (CommitCell?.Invoke(_sheetIndex, _draftOwner, CellEditor.Text ?? "") != true) return false;
        _editing = false; CellEditor.Opacity = 0; return true;
    }
    public void Cancel() { _editing = false; if (_snapshot is not null) PrepareEditor(false); }
    public bool BeginEdit()
    {
        if (_snapshot is null || IsComposing || _snapshot.FileState?.FormulaReadOnly == true || !ProtectionSafety.CanEdit(_snapshot, Sheet, _current.Row, _current.Column)) return false;
        _editing = true; _draftOwner = _current; CellEditor.Opacity = 1; CellEditor.Focus(); CellEditor.CaretIndex = CellEditor.Text?.Length ?? 0; return true;
    }
    private void PrepareEditor(bool focus)
    {
        if (_snapshot is null || _editing || IsComposing) return;
        _draftOwner = _current; _settingText = true;
        CellEditor.Text = ProtectionSafety.HidesFormula(_snapshot, Sheet, _current.Row, _current.Column) ? "" :
            Sheet.GetCell(_current.Row, _current.Column) is { } cell ? CellInput.EditText(cell.Content) : "";
        CellEditor.IsReadOnly = _snapshot.FileState?.FormulaReadOnly == true || !ProtectionSafety.CanEdit(_snapshot, Sheet, _current.Row, _current.Column);
        _settingText = false; CellEditor.Opacity = 0;
        PositionEditor(); if (focus) { CellEditor.Focus(); CellEditor.SelectAll(); }
    }
    private Rect? CellBounds(CellPosition position)
    {
        if (_rows.Cell(position.Row) is not { } row || _columns.Cell(position.Column) is not { } col) return null;
        return new((col.Start + HeaderWidth) * _zoom, (row.Start + HeaderHeight) * _zoom, col.Size * _zoom, row.Size * _zoom);
    }
    private void PositionEditor()
    {
        if (CellBounds(_draftOwner) is not { } rect || !rect.Intersects(new Rect(_canvas.Bounds.Size))) { CellEditor.IsVisible = false; return; }
        CellEditor.IsVisible = !_dragging;
        Canvas.SetLeft(CellEditor, rect.X); Canvas.SetTop(CellEditor, rect.Y);
        CellEditor.Width = Math.Max(1, Math.Min(rect.Width, _canvas.Bounds.Width - rect.X)); CellEditor.Height = Math.Max(1, Math.Min(rect.Height, _canvas.Bounds.Height - rect.Y));
        CellEditor.FontSize = 14 * _zoom;
    }
    private void Scroll(double x, double y)
    {
        _rows.ScrollTo(y, ViewHeight); _columns.ScrollTo(x, ViewWidth); UpdateScrollbars(); InvalidateContent(); PositionEditor(); ViewportChanged?.Invoke();
    }
    private void UpdateScrollbars()
    {
        _scrolling = true;
        _rows.ScrollTo(_rows.Offset, ViewHeight); _columns.ScrollTo(_columns.Offset, ViewWidth);
        _vertical.Maximum = Math.Max(0, _rows.ScrollSize - Math.Max(0, ViewHeight - _rows.FrozenSize)); _vertical.ViewportSize = ViewHeight; _vertical.Value = _rows.Offset;
        _horizontal.Maximum = Math.Max(0, _columns.ScrollSize - Math.Max(0, ViewWidth - _columns.FrozenSize)); _horizontal.ViewportSize = ViewWidth; _horizontal.Value = _columns.Offset;
        _scrolling = false;
    }
    private CellPosition? Hit(Point point)
    {
        var row = _rows.HitTest(point.Y / _zoom - HeaderHeight, ViewHeight); var col = _columns.HitTest(point.X / _zoom - HeaderWidth, ViewWidth);
        return row >= 0 && col >= 0 ? new(row, col) : null;
    }
    private void PointerDown(object? sender, PointerPressedEventArgs e)
    {
        if (e.Source == CellEditor || !e.GetCurrentPoint(_canvas).Properties.IsLeftButtonPressed || Hit(e.GetPosition(_canvas)) is not { } cell) return;
        if (!Commit() || !(BeforeSelection?.Invoke() ?? true)) return;
        SelectCore(cell, e.KeyModifiers.HasFlag(KeyModifiers.Shift), false);
        if (e.ClickCount == 2) { BeginEdit(); e.Handled = true; return; }
        _dragging = true; CellEditor.IsVisible = false; e.Pointer.Capture(_canvas); e.Handled = true;
    }
    private void PointerMove(object? sender, PointerEventArgs e)
    {
        if (_dragging && Hit(e.GetPosition(_canvas)) is { } cell && cell != _current) SelectCore(cell, true, false);
    }
    private void HandleKey(object? sender, KeyEventArgs e)
    {
        if (_snapshot is null || IsComposing) return;
        if (e.Key == Key.F2) { e.Handled = BeginEdit(); return; }
        if (e.Key == Key.Escape && _editing) { Cancel(); e.Handled = true; return; }
        if (e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta) || e.KeyModifiers.HasFlag(KeyModifiers.Alt)) return;
        if (_editing && e.Key is not (Key.Enter or Key.Tab)) return;
        var delta = e.Key switch { Key.Up => (-1, 0), Key.Down => (1, 0), Key.Left => (0, -1), Key.Right => (0, 1),
            Key.Enter => (e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? -1 : 1, 0), Key.Tab => (0, e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? -1 : 1), _ => (0, 0) };
        if (delta == (0, 0)) return;
        e.Handled = true;
        if (!Commit() || !(BeforeSelection?.Invoke() ?? true)) return;
        SelectCore(new(_rows.Neighbor(_current.Row, delta.Item1), _columns.Neighbor(_current.Column, delta.Item2)),
            e.Key is not (Key.Enter or Key.Tab) && e.KeyModifiers.HasFlag(KeyModifiers.Shift), true);
    }
    private void InvalidateContent() { _artwork.Release(); _artwork.InvalidateVisual(); _overlay.InvalidateVisual(); }
    private void Paint(SKCanvas canvas)
    {
        if (_snapshot is null) return;
        canvas.Clear(SKColors.White); canvas.Scale((float)_zoom);
        var rows = _rows.Visible(ViewHeight); var columns = _columns.Visible(ViewWidth);
        foreach (var row in rows)
        foreach (var col in columns)
        {
            var rect = new RenderRect(HeaderWidth + col.Start, HeaderHeight + row.Start, col.Size, row.Size);
            canvas.Save(); canvas.ClipRect(SKRect.Create((float)(HeaderWidth + (col.Frozen ? 0 : _columns.FrozenSize)),
                (float)(HeaderHeight + (row.Frozen ? 0 : _rows.FrozenSize)), (float)ViewWidth, (float)ViewHeight));
            var cell = Sheet.GetCell(row.ModelIndex, col.ModelIndex);
            var style = ResolveStyle?.Invoke(_snapshot, Sheet, row.ModelIndex, col.ModelIndex) ?? _snapshot.Styles[Sheet.GetStyleId(row.ModelIndex, col.ModelIndex)];
            var display = CellDisplayFormatter.FormatDisplay(cell?.Value ?? CellValue.Blank, style, _snapshot.DateSystem);
            if (display.Color is not null) style = style with { FontColor = display.Color };
            if (style.Horizontal == 0 && cell?.Value.Kind == CellValueKind.Number) style = style with { Horizontal = 3 };
            _renderer.DrawCell(canvas, new(rect, new(display.Text, style, RichText: cell?.Content.RichText, DefaultFontSize: 14,
                FillOffset: display.FillOffset, FillCharacter: display.FillCharacter), Sheet.ShowGridLines,
                Sheet.Annotations.Notes.ContainsKey(new(row.ModelIndex, col.ModelIndex))), RenderOutput.Screen);
            canvas.Restore();
        }
        var header = new StyleDefinition { FillColor = "FFF2F5F4", FontColor = "FF42534E", Horizontal = 2, Vertical = 1 };
        foreach (var col in columns) _renderer.DrawCell(canvas, new(new(HeaderWidth + col.Start, 0, col.Size, HeaderHeight), new(ColumnName(col.ModelIndex), header, DefaultFontSize: 12), true), RenderOutput.Screen);
        foreach (var row in rows) _renderer.DrawCell(canvas, new(new(0, HeaderHeight + row.Start, HeaderWidth, row.Size), new((row.ModelIndex + 1).ToString(), header, DefaultFontSize: 12), true), RenderOutput.Screen);
        _renderer.DrawCell(canvas, new(new(0, 0, HeaderWidth, HeaderHeight), new("", header), true), RenderOutput.Screen);
    }
    public static string ColumnName(int column)
    {
        var text = ""; for (column++; column > 0; column = (column - 1) / 26) text = (char)('A' + (column - 1) % 26) + text; return text;
    }
    public void Dispose() { if (_disposed) return; _disposed = true; _artwork.Release(); _renderer.Dispose(); }
    private sealed class GridArtwork(WorkbookGrid owner) : Control
    {
        private Bitmap? _bitmap;
        private double _scale;
        public void Release() { _bitmap?.Dispose(); _bitmap = null; }
        public override void Render(DrawingContext context)
        {
            var scale = TopLevel.GetTopLevel(this)?.RenderScaling ?? 1;
            if (scale != _scale) { Release(); _scale = scale; }
            if (_bitmap is null && Bounds.Width > 0 && Bounds.Height > 0)
            {
                var width = (int)Math.Ceiling(Bounds.Width * scale); var height = (int)Math.Ceiling(Bounds.Height * scale);
                if ((long)width * height > 16_000_000) throw new InvalidOperationException("网格视口超过像素预算。");
                using var pixels = new SKBitmap(width, height, SKColorType.Bgra8888, SKAlphaType.Premul);
                using var canvas = new SKCanvas(pixels); canvas.Scale((float)scale); owner.Paint(canvas); canvas.Flush();
                _bitmap = new(PixelFormat.Bgra8888, AlphaFormat.Premul, pixels.GetPixels(), new(width, height), new(96 * scale, 96 * scale), pixels.RowBytes);
            }
            // Bitmap.Draw passes source rectangles to the backend in physical pixels;
            // using bitmap.Size here crops a high-DPI raster and scales it a second time.
            if (_bitmap is not null) context.DrawImage(_bitmap, new Rect(0, 0, _bitmap.PixelSize.Width, _bitmap.PixelSize.Height), new Rect(Bounds.Size));
        }
    }
    private sealed class SelectionOverlay(WorkbookGrid owner) : Control
    {
        public override void Render(DrawingContext context)
        {
            if (owner._snapshot is null) return;
            using var clip = context.PushClip(new Rect(HeaderWidth * owner._zoom, HeaderHeight * owner._zoom, owner.ViewWidth * owner._zoom, owner.ViewHeight * owner._zoom));
            foreach (var range in owner._selection.Ranges)
            {
                var rows = owner._rows.Visible(owner.ViewHeight).Where(r => r.ModelIndex >= range.Top && r.ModelIndex <= range.Bottom).ToArray();
                var cols = owner._columns.Visible(owner.ViewWidth).Where(c => c.ModelIndex >= range.Left && c.ModelIndex <= range.Right).ToArray();
                if (rows.Length == 0 || cols.Length == 0) continue;
                var rect = new Rect((HeaderWidth + cols.Min(c => c.Start)) * owner._zoom, (HeaderHeight + rows.Min(r => r.Start)) * owner._zoom,
                    (cols.Max(c => c.Start + c.Size) - cols.Min(c => c.Start)) * owner._zoom, (rows.Max(r => r.Start + r.Size) - rows.Min(r => r.Start)) * owner._zoom);
                context.DrawRectangle(new SolidColorBrush(Color.FromArgb(22, 0, 128, 110)), new Pen(Brushes.Teal, 1.5), rect);
            }
        }
    }
}
