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
public sealed partial class WorkbookGrid : UserControl, IGridSurface, IDisposable
{
    private const double HeaderHeight = 28;
    private double HeaderWidth => 52 + OutlineWidth;
    private readonly ViewportAxis _rows = new(), _columns = new();
    private readonly Canvas _canvas = new() { ClipToBounds = true };
    private readonly GridArtwork _artwork;
    private readonly SelectionOverlay _overlay;
    private readonly ScrollBar _vertical = new() { Orientation = global::Avalonia.Layout.Orientation.Vertical, Width = 16, AllowAutoHide = false };
    private readonly ScrollBar _horizontal = new() { Orientation = global::Avalonia.Layout.Orientation.Horizontal, Height = 16, AllowAutoHide = false };
    private SkiaRenderer? _renderer;
    private SkiaRenderer Renderer => _renderer ??= new();
    private WorkbookSnapshot? _snapshot;
    private int _sheetIndex;
    private CellPosition _current, _anchor, _draftOwner;
    private GridCellRange[] _merges = [];
    private readonly GridRangeSelection _selection = new();
    private bool _settingText, _editing, _scrolling, _dragging, _disposed;
    private double _zoom = 1;
    public CompositionTextBox CellEditor { get; } = new() { AcceptsReturn = true, NewLine = "\n", Padding = new(2, 0), MinHeight = 0, MinWidth = 0,
        BorderThickness = new(1), BorderBrush = Brushes.Teal, Background = Brushes.White, Opacity = 0, IsHitTestVisible = false };
    public Func<int, CellPosition, string, bool>? CommitCell { get; set; }
    public Func<int, CellPosition, string, Task<bool>>? CommitCellAsync { get; set; }
    private Task<bool>? _pendingCommit;
    public bool IsCommitPending => _pendingCommit is { IsCompleted: false };
    public Func<bool>? BeforeSelection { get; set; }
    public Func<WorkbookSnapshot, WorksheetSnapshot, int, int, StyleDefinition>? ResolveStyle { get; set; }
    public Func<WorkbookSnapshot, WorksheetSnapshot, int, int, ConditionalCellVisual?>? ResolveConditionalVisual { get; set; }
    public event Action? SelectionChanged;
    public event Action? DraftEnded;
    public event Action? DraftStarted;
    public event Action? CellClicked;
    public event Action<int>? ZoomRequested;
    public event Action<Exception>? RenderingFailed;
    public event Action? ViewportChanged;
    public CellPosition Current => _current;
    public bool IsEditing => _editing;
    public bool IsReadOnly { get; private set; }
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
        InstallChartControls();
        var layout = new global::Avalonia.Controls.Grid { RowDefinitions = new("*,16"), ColumnDefinitions = new("*,16") };
        layout.Children.Add(_canvas); layout.Children.Add(_vertical); layout.Children.Add(_horizontal);
        global::Avalonia.Controls.Grid.SetColumn(_vertical, 1); global::Avalonia.Controls.Grid.SetRow(_horizontal, 1);
        Content = layout;
        _canvas.SizeChanged += (_, _) => { _artwork.Width = _overlay.Width = _canvas.Bounds.Width; _artwork.Height = _overlay.Height = _canvas.Bounds.Height; RestoreScrollViewport(); UpdateScrollbars(); InvalidateContent(); PositionEditor(); };
        _vertical.ValueChanged += (_, _) => { if (!_scrolling) Scroll(_columns.Offset, _vertical.Value); };
        _horizontal.ValueChanged += (_, _) => { if (!_scrolling) Scroll(_horizontal.Value, _rows.Offset); };
        _canvas.PointerPressed += PointerDown; _canvas.PointerMoved += PointerMove;
        _canvas.PointerReleased += PointerUp;
        _canvas.PointerCaptureLost += (_, _) => CancelGesture();
        _canvas.PointerWheelChanged += (_, e) =>
        {
            if (IsDrawingObjectDragging) { e.Handled = true; return; }
            if (e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta))
            {
                e.Handled = true; if (IsComposing || e.Delta.Y == 0) return;
                int[] stops = [10, 25, 50, 75, 100, 125, 150, 175, 200, 300, 400];
                ZoomRequested?.Invoke(e.Delta.Y > 0 ? stops.FirstOrDefault(s => s > _zoom * 100, 400) : stops.LastOrDefault(s => s < _zoom * 100, 10)); return;
            }
            if (!Commit()) return;
            Scroll(_columns.Offset - (e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? e.Delta.Y : e.Delta.X) * 60, _rows.Offset - (e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? 0 : e.Delta.Y) * 60);
            e.Handled = true;
        };
        CellEditor.PropertyChanged += (_, e) => { if (e.Property == TextBox.TextProperty && !_settingText) { if (!_editing) DraftStarted?.Invoke(); _editing = true; CellEditor.Opacity = 1; CellEditor.IsHitTestVisible = true; } };
        CellEditor.CompositionChanged += () => { if (IsComposing) { if (!_editing) DraftStarted?.Invoke(); _editing = true; CellEditor.Opacity = 1; CellEditor.IsHitTestVisible = true; } };
        AddHandler(KeyDownEvent, HandleKey, RoutingStrategies.Tunnel);
        DetachedFromVisualTree += (_, _) => { CancelGesture(); ReleaseRendering(); };
    }
    public void Bind(WorkbookSnapshot snapshot, int sheetIndex, int rows, int columns) => Bind(snapshot, sheetIndex, rows, columns, snapshot.DocumentId);
    public void Bind(WorkbookSnapshot snapshot, int sheetIndex, int rows, int columns, Guid sessionId, bool readOnly = false)
    {
        IsReadOnly = readOnly;
        SetFormulaAuditArrows(null);
        ClearDrawingObjectSelection();
        Cancel(); ReleaseRendering(); _snapshot = snapshot; _sheetIndex = sheetIndex; _zoom = Sheet.View.Zoom / 100d;
        BindScrollExtent(sessionId);
        Configure(rows, columns); _rows.ScrollTo(0, ViewHeight); _columns.ScrollTo(0, ViewWidth);
        SelectCore(Sheet.View.ActiveCell, false, false);
        if (Sheet.View.Selection is { } selection)
            _selection.Set(ExpandSelection(new(selection.Range.Top, selection.Range.Left, selection.Range.Bottom, selection.Range.Right)));
        IncludeSelectionExtent();
        _pendingScrollCell = Sheet.View.HasFreeze ? Sheet.View.PaneTopLeftCell : Sheet.View.TopLeftCell;
        RestoreScrollViewport();
        InvalidateContent(); UpdateScrollbars(); PositionEditor();
    }
    public void Refresh(WorkbookSnapshot snapshot)
    {
        if (IsDrawingObjectDragging && snapshot.Revision != _objectRevision) CancelObjectGesture();
        if (_snapshot is { } previous && previous.Revision == snapshot.Revision &&
            ReferenceEquals(previous.Sheets[_sheetIndex].Cells, snapshot.Sheets[_sheetIndex].Cells) &&
            previous.Sheets[_sheetIndex].View.Zoom == snapshot.Sheets[_sheetIndex].View.Zoom)
        { _snapshot = snapshot; if (!_editing) PrepareEditor(false); return; }
        var selectedMerge = _selection.Ranges.Count == 1 && _merges.Contains(_selection.Ranges[0])
            ? _selection.Ranges[0] : (GridCellRange?)null;
        if (!ReferenceEquals(_snapshot, snapshot)) ClearScene();
        _snapshot = snapshot; _zoom = Sheet.View.Zoom / 100d; Configure(_rows.Count, _columns.Count);
        // Insert/undo/redo can resize the selected merged cell without moving its anchor.
        if (selectedMerge is { } oldMerge)
            foreach (var merge in _merges)
                if (merge.Top == oldMerge.Top && merge.Left == oldMerge.Left) { _selection.Set(merge); break; }
        if (_selectedObject is { } selected && ObjectAnchor(selected) is null) ClearDrawingObjectSelection();
        UpdateScrollbars(); InvalidateContent(); if (!_editing) PrepareEditor(false); else PositionEditor();
    }
    private void Configure(int rows, int columns)
    {
        _rowOutline = Sheet.Data.Outline?.ShowSymbols == false ? [] : WorksheetOutline.Groups(Sheet, SheetAxis.Row);
        _outlineDepth = _rowOutline.Count == 0 ? 0 : _rowOutline.Max(g => g.Level);
        var view = new SheetViewState(); view.Synchronize(_snapshot!, Sheet.Id, _current); view.RefreshFilter(_snapshot!);
        var hiddenRows = Sheet.Rows.Where(p => p.Value.Hidden).Select(p => p.Key).ToHashSet();
        foreach (var filter in Sheet.Data.AllFilters)
            for (var r = filter.Range.Top; r <= filter.Range.Bottom; r++) if (view.IsFilteredOut(_snapshot!, _sheetIndex, r)) hiddenRows.Add(r);
        _rows.Configure(rows, Math.Clamp((int)(Sheet.DefaultRowHeight ?? 31), 1, 10000),
            Sheet.Rows.Where(p => p.Value.Height.HasValue).Select(p => new KeyValuePair<int, int>(p.Key, (int)Math.Clamp(p.Value.Height!.Value, 1, 10000))),
            hiddenRows, Enumerable.Range(0, Math.Min(rows, Sheet.View.FrozenRows)).ToHashSet());
        _columns.Configure(columns, (int)Math.Clamp(Sheet.Data.DefaultColumnWidth is { } width ? Math.Floor(width * 7 + 5) : 112, 1, 10000),
            Sheet.Columns.Keys.Append(0).Distinct().Select(c => new KeyValuePair<int, int>(c, (int)Math.Clamp(DrawingGeometry.ColumnSize(Sheet, c), 1, 10000))),
            Sheet.Columns.Where(p => p.Value.Hidden).Select(p => p.Key).ToHashSet(), Enumerable.Range(0, Math.Min(columns, Sheet.View.FrozenColumns)).ToHashSet());
        _merges = Sheet.MergedCells.Select(CellRectangle.Parse).Select(m => new GridCellRange(m.Top, m.Left, m.Bottom, m.Right)).ToArray();
        IncludeUsedExtent();
    }
    public void Select(CellPosition position, bool extend = false)
    { if (Commit() && (BeforeSelection?.Invoke() ?? true)) SelectCore(position, extend, true); }
    private void SelectCore(CellPosition position, bool extend, bool focus)
    {
        ClearDrawingObjectSelection();
        var row = _rows.NearestVisible(position.Row); var col = _columns.NearestVisible(position.Column);
        if (row < 0 || col < 0) return;
        var anchor = CellAnchor(new(row, col));
        if (!ProtectionSafety.CanSelect(_snapshot!, Sheet, anchor.Row, anchor.Column)) return;
        _current = anchor;
        if (!extend) _anchor = _current;
        var range = GridRange.Between(_anchor, _current);
        _selection.Set(GridMergedSelection.Expand(new(range.Top, range.Left, range.Bottom, range.Right), _merges));
        IncludeSelectionExtent();
        var offsets = (_rows.Offset, _columns.Offset);
        _rows.EnsureVisible(_current.Row, ViewHeight); _columns.EnsureVisible(_current.Column, ViewWidth);
        UpdateScrollbars();
        if (offsets != (_rows.Offset, _columns.Offset)) InvalidateContent(); else _overlay.InvalidateVisual();
        PrepareEditor(focus); SelectionChanged?.Invoke();
    }
    public void SelectRange(GridRange range, CellPosition current)
    {
        if (!CanSelectRanges([range], current)) return;
        if (!Commit() || !(BeforeSelection?.Invoke() ?? true)) return;
        ClearDrawingObjectSelection();
        _selection.Set(ExpandSelection(new(range.Top, range.Left, range.Bottom, range.Right)));
        _current = current; _anchor = new(range.Top, range.Left); IncludeSelectionExtent(); _overlay.InvalidateVisual(); PrepareEditor(false); SelectionChanged?.Invoke();
    }
    private GridCellRange ExpandSelection(GridCellRange range) =>
        // Whole-axis selections retain their explicit header boundaries, just like header drags.
        range.Left == 0 && range.Right == _columns.Count - 1 || range.Top == 0 && range.Bottom == _rows.Count - 1
            ? range : GridMergedSelection.Expand(range, _merges);
    public bool Commit()
    {
        var task = CommitAsync();
        return task.IsCompletedSuccessfully && task.Result;
    }
    public Task<bool> CommitAsync()
    {
        if (_pendingCommit is { } pending) return pending;
        if (IsComposing) return Task.FromResult(false);
        if (!_editing) return Task.FromResult(true);
        var task = CommitDraftAsync();
        if (!task.IsCompleted) _pendingCommit = task;
        return task;
    }
    private async Task<bool> CommitDraftAsync()
    {
        CellEditor.IsReadOnly = true;
        try
        {
            var accepted = CommitCellAsync is { } commit
                ? await commit(_sheetIndex, _draftOwner, CellEditor.Text ?? "")
                : CommitCell?.Invoke(_sheetIndex, _draftOwner, CellEditor.Text ?? "") == true;
            if (!accepted) return false;
            _editing = false; CellEditor.Opacity = 0; CellEditor.IsHitTestVisible = false; DraftEnded?.Invoke(); return true;
        }
        finally
        {
            _pendingCommit = null;
            if (_snapshot is not null) CellEditor.IsReadOnly = IsCellReadOnly(_draftOwner);
        }
    }
    public void Cancel() { if (IsCommitPending) return; _editing = false; if (_snapshot is not null) PrepareEditor(false); DraftEnded?.Invoke(); }
    public bool IsCellReadOnly(CellPosition position) => IsReadOnly || _snapshot is null || _snapshot.FileState?.FormulaReadOnly == true ||
        PivotReport.HasUnmodeledDefinitions(_snapshot) ||
        Sheet.GetCell(position.Row, position.Column)?.SpillOwner is not null ||
        Sheet.Data.Pivots.Any(p=>p.Range.Contains(position.Row,position.Column)) ||
        Sheet.GetCell(position.Row, position.Column)?.Content.RichText is not null || !ProtectionSafety.CanEdit(_snapshot, Sheet, position.Row, position.Column);
    public bool BeginEdit()
    {
        if (_selectedObject is not null) return false;
        if (_snapshot is null || IsComposing || IsCommitPending || IsCellReadOnly(_current)) return false;
        if (!_editing) DraftStarted?.Invoke();
        _editing = true; _draftOwner = _current; CellEditor.Opacity = 1; CellEditor.IsHitTestVisible = true; CellEditor.Focus(); CellEditor.CaretIndex = CellEditor.Text?.Length ?? 0; return true;
    }
    private void PrepareEditor(bool focus)
    {
        if (_snapshot is null || _editing || IsComposing || _selectedObject is not null) return;
        _draftOwner = _current; _settingText = true;
        CellEditor.Text = CellInput.EditText(_snapshot, Sheet, _current.Row, _current.Column);
        CellEditor.IsReadOnly = IsCellReadOnly(_current);
        _settingText = false; CellEditor.Opacity = 0; CellEditor.IsHitTestVisible = false;
        PositionEditor(); if (focus) { CellEditor.Focus(); CellEditor.SelectAll(); }
    }
    private Rect? CellBounds(CellPosition position)
    {
        foreach (var fragment in Fragments(CellRange(position)))
        {
            var visible = fragment.Bounds.Intersect(fragment.Clip);
            if (visible.Width > 0 && visible.Height > 0) return new(visible.X * _zoom, visible.Y * _zoom, visible.Width * _zoom, visible.Height * _zoom);
        }
        return null;
    }
    /// <summary>Visible range bounds in grid coordinates, clipped to frozen/scrolling panes.</summary>
    public Rect? GetVisibleRangeBounds(GridRange range)
    {
        Rect? result = null;
        foreach (var fragment in Fragments(new(range.Top, range.Left, range.Bottom, range.Right)))
        {
            var visible = fragment.Bounds.Intersect(fragment.Clip);
            if (visible.Width <= 0 || visible.Height <= 0) continue;
            var scaled = new Rect(visible.X * _zoom, visible.Y * _zoom, visible.Width * _zoom, visible.Height * _zoom);
            result = result is { } previous ? previous.Union(scaled) : scaled;
        }
        return result;
    }
    private void PositionEditor()
    {
        if (_snapshot is null || _rows.Count == 0 || _columns.Count == 0) { CellEditor.IsVisible = false; return; }
        if (CellBounds(_draftOwner) is not { } rect || !rect.Intersects(new Rect(_canvas.Bounds.Size))) { CellEditor.IsVisible = false; return; }
        CellEditor.IsVisible = !_dragging && _selectedObject is null;
        Canvas.SetLeft(CellEditor, rect.X); Canvas.SetTop(CellEditor, rect.Y);
        CellEditor.Width = Math.Max(1, Math.Min(rect.Width, _canvas.Bounds.Width - rect.X)); CellEditor.Height = Math.Max(1, Math.Min(rect.Height, _canvas.Bounds.Height - rect.Y));
        var anchor = CellAnchor(_draftOwner);
        var style = ResolveStyle?.Invoke(_snapshot, Sheet, anchor.Row, anchor.Column)
            ?? TablePresentation.GetStyle(_snapshot, Sheet, anchor.Row, anchor.Column);
        var value = Sheet.GetCell(anchor.Row, anchor.Column)?.Value ?? CellValue.Blank;
        var display = CellDisplayFormatter.FormatDisplay(value, style, _snapshot.DateSystem);
        var font = Renderer.ResolveFont(new(style.FontName ?? "Microsoft YaHei UI", style.EastAsianFontName,
            style.FontSize ?? 14 * 72d / 96, style.Bold, style.Italic));
        CellEditor.FontFamily = new FontFamily(font.Family!);
        CellEditor.FontSize = font.Points * 96 / 72 * _zoom;
        CellEditor.FontWeight = font.Bold ? FontWeight.Bold : FontWeight.Normal;
        CellEditor.FontStyle = font.Italic ? FontStyle.Italic : FontStyle.Normal;
        var foreground = SkiaRenderer.Color(display.Color ?? style.FontColor, SKColors.Black);
        CellEditor.Foreground = new SolidColorBrush(Color.FromArgb(foreground.Alpha, foreground.Red, foreground.Green, foreground.Blue));
        CellEditor.CaretBrush = CellEditor.Foreground;
        var background = SkiaRenderer.Color(style.FillColor, SKColors.White);
        CellEditor.Background = new SolidColorBrush(Color.FromArgb(background.Alpha, background.Red, background.Green, background.Blue));
        CellEditor.TextWrapping = style.WrapText ? TextWrapping.Wrap : TextWrapping.NoWrap;
        ScrollViewer.SetHorizontalScrollBarVisibility(CellEditor, style.WrapText ? ScrollBarVisibility.Disabled : ScrollBarVisibility.Hidden);
        ScrollViewer.SetVerticalScrollBarVisibility(CellEditor, ScrollBarVisibility.Hidden);
        // The renderer reserves two pixels on each horizontal edge and one vertically.
        // Include the editor border in that inset so both layouts wrap at the same width.
        CellEditor.BorderThickness = new(_zoom);
        CellEditor.Padding = new(_zoom, 0);
        // Match the grid's effective style and General alignment, including merged anchors
        // and inherited row/column formatting. Keep the native presenter responsible for caret/IME layout.
        CellEditor.TextAlignment = style.Horizontal switch
        {
            2 => TextAlignment.Center,
            3 => TextAlignment.Right,
            0 when value.Kind == CellValueKind.Number => TextAlignment.Right,
            _ => TextAlignment.Left
        };
        CellEditor.VerticalContentAlignment = style.Vertical switch
        {
            0 => global::Avalonia.Layout.VerticalAlignment.Top,
            1 => global::Avalonia.Layout.VerticalAlignment.Center,
            _ => global::Avalonia.Layout.VerticalAlignment.Bottom
        };
    }
    private CellPosition? Hit(Point point)
    {
        var row = _rows.HitTest(point.Y / _zoom - HeaderHeight, ViewHeight); var col = _columns.HitTest(point.X / _zoom - HeaderWidth, ViewWidth);
        return row >= 0 && col >= 0 ? new(row, col) : null;
    }
    private CellPosition CellAnchor(CellPosition position)
    {
        var range = CellRange(position);
        return new(range.Top, range.Left);
    }
    private void InvalidateContent() { _artwork.Release(); _artwork.InvalidateVisual(); _overlay.InvalidateVisual(); PositionChartControls(); PositionOutlineControls(); }
    public static string ColumnName(int column)
    {
        var text = ""; for (column++; column > 0; column = (column - 1) / 26) text = (char)('A' + (column - 1) % 26) + text; return text;
    }
    public void Dispose() { if (_disposed) return; _disposed = true; ReleaseRendering(); }
    private sealed class GridArtwork(WorkbookGrid owner) : Control
    {
        private Bitmap? _bitmap;
        private double _scale;
        private bool _failed;
        public void Release() { _bitmap?.Dispose(); _bitmap = null; _failed = false; }
        public override void Render(DrawingContext context)
        {
            if (owner._disposed) return;
            var scale = TopLevel.GetTopLevel(this)?.RenderScaling ?? 1;
            if (scale != _scale) { Release(); _scale = scale; }
            if (!_failed && _bitmap is null && Bounds.Width > 0 && Bounds.Height > 0)
            {
                try
                {
                var width = (int)Math.Ceiling(Bounds.Width * scale); var height = (int)Math.Ceiling(Bounds.Height * scale);
                if ((long)width * height > 16_000_000) throw new InvalidOperationException(UiText.Get("网格视口超过像素预算。"));
                using var pixels = new SKBitmap(width, height, SKColorType.Bgra8888, SKAlphaType.Premul);
                using var canvas = new SKCanvas(pixels); canvas.Scale((float)scale); owner.Paint(canvas); canvas.Flush();
                _bitmap = new(PixelFormat.Bgra8888, AlphaFormat.Premul, pixels.GetPixels(), new(width, height), new(96 * scale, 96 * scale), pixels.RowBytes);
                }
                catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or IOException or OutOfMemoryException)
                { owner.ReleaseRendering(); _failed = true; owner.RenderingFailed?.Invoke(ex); }
            }
            if (_failed) context.DrawText(new FormattedText(UiText.Get("网格暂时无法绘制，请调整窗口或重新打开工作簿。"), System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight, Typeface.Default, 14, Brushes.DarkRed), new(8, 8));
            // Bitmap.Draw passes source rectangles to the backend in physical pixels;
            // using bitmap.Size here crops a high-DPI raster and scales it a second time.
            if (_bitmap is not null) context.DrawImage(_bitmap, new Rect(0, 0, _bitmap.PixelSize.Width, _bitmap.PixelSize.Height), new Rect(Bounds.Size));
            owner.RenderingTrace?.Invoke("content-submit", System.Diagnostics.Stopwatch.GetTimestamp());
        }
    }
    private sealed class SelectionOverlay(WorkbookGrid owner) : Control
    {
        public override void Render(DrawingContext context) => owner.PaintOverlay(context);
    }
}
