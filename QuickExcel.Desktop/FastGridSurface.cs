using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using FastWpfGrid;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>Snapshot projection plus a native editor. Commit is an injected controller operation, never a model setter.</summary>
internal sealed partial class FastGridSurface : IGridSurface
{
    internal FastGridControl Control { get; } = new() { IsReadOnly = true, UseSelectionOverlay = true, UseSmoothScrolling = true, AllowColumnResize = false, UseClearType = false, CellFontName = "Microsoft YaHei UI", CellFontSize = 14 };
    public Grid Element { get; } = new();
    private readonly Canvas _overlay = new() { ClipToBounds = true };
    internal TextBox Editor { get; } = new() { BorderThickness = new(1), Padding = new(3, 1, 3, 1), Visibility = Visibility.Hidden, FocusVisualStyle = null };
    internal FastGridArtwork Artwork { get; }
    private readonly SnapshotModel _model = new();
    private readonly Func<CellPosition, string, bool> _commit;
    private readonly Dictionary<int, Button> _headers = [];
    private bool _loading, _editing, _selecting, _selectionPresentationQueued;
    private CellPosition _current, _anchor;
    private WorkbookSnapshot _snapshot = null!;
    private int _sheetIndex;
    internal bool HandlesKeyboard { get; init; } = true;
    internal Func<bool>? BeforeNavigate { get; set; }
    internal Func<bool>? BeforeBeginEdit { get; set; }
    internal Func<WorkbookSnapshot>? LatestSnapshot { get; set; }
    internal bool IsEditing => _editing;
    internal WorkbookSnapshot Snapshot => _snapshot;
    internal WorksheetSnapshot Sheet => _snapshot.Sheets[_sheetIndex];
    internal IReadOnlyCollection<Button> HeaderButtons => _headers.Values;
    internal int RowCount => _model.Rows;
    internal int ColumnCount => _model.Columns;
    internal event Action? EditingChanged;
    public bool IsComposing { get; private set; }
    public SheetViewState View { get; set; } = new();
    public CellPosition Current => _current;
    internal GridCellRange CurrentRange => Control.GetMergedModelRange(_current.Row, _current.Column);
    private CellPosition MergeOwner(int row, int column)
    {
        var range = Control.GetMergedModelRange(row, column);
        return new(range.Top, range.Left);
    }
    public long ReadCount => _model.ReadCount;
    public event Action<int, FrameworkElement>? HeaderRequested;
    public event Action? SelectionChanged;
    public event Action<Key, ModifierKeys>? CommandRequested;
    public IReadOnlyList<GridRange> SelectionRanges => Control.GetSelectedModelRanges().Select(r => new GridRange(r.Top, r.Left, r.Bottom, r.Right)).ToArray();
    public IReadOnlyList<CellPosition> Selection
    {
        get
        {
            if (Control.SelectedCellCount > 100_000) throw new ArgumentException("操作超过 100,000 格，请缩小选区。");
            return Control.GetSelectedModelCells().Select(c => new CellPosition(c.Row!.Value, c.Column!.Value)).ToArray();
        }
    }
    public FastGridSurface(Func<CellPosition, string, bool> commit)
    {
        _commit = commit;
        Artwork = new(this);
        Control.SetContentDecoration(Artwork);
        Element.Children.Add(Control); Element.Children.Add(_overlay); _overlay.Children.Add(Editor);
        Control.SelectedCellsChanged += (_, _) =>
        {
            if (_selecting || Control.CurrentModelCell is not { IsCell: true } cell) return;
            _current = MergeOwner(cell.Row!.Value, cell.Column!.Value);
            if (!Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) _anchor = _current;
            QueueSelectionPresentation();
        };
        Control.LostMouseCapture += (_, _) => QueueSelectionPresentation();
        Control.ViewportChanged += (_, _) => { PositionOverlay(); Artwork.InvalidateVisual(); };
        Element.SizeChanged += (_, _) => PositionOverlay();
        Element.PreviewMouseDown += (_, e) =>
        {
            if (e.OriginalSource is DependencyObject target && (ReferenceEquals(target, Editor) || Editor.IsAncestorOf(target))) return;
            if (!Commit() || BeforeNavigate?.Invoke() == false) { e.Handled = true; Editor.Focus(); }
        };
        Control.PreviewMouseDoubleClick += (_, e) =>
        {
            var cell = Control.HitTestModel(e.GetPosition(Control));
            if (cell.IsCell && e.ChangedButton == MouseButton.Left && BeginEdit()) e.Handled = true;
        };
        Editor.TextChanged += (_, _) =>
        {
            if (_loading) return;
            ActivateDraft();
        };
        Editor.AddHandler(TextCompositionManager.PreviewTextInputStartEvent, new TextCompositionEventHandler((_, _) => IsComposing = true));
        Editor.AddHandler(TextCompositionManager.PreviewTextInputUpdateEvent, new TextCompositionEventHandler((_, _) => IsComposing = true));
        Editor.AddHandler(TextCompositionManager.PreviewTextInputEvent, new TextCompositionEventHandler((_, _) =>
            Element.Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() => IsComposing = false))));
        Element.PreviewKeyDown += HandleKey;
        Control.Model = _model;
    }

    private void QueueSelectionPresentation()
    {
        if (_selectionPresentationQueued) return;
        _selectionPresentationQueued = true;
        Element.Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() =>
        {
            _selectionPresentationQueued = false;
            if (!Element.IsLoaded) return;
            SelectionChanged?.Invoke();
            // Dragging owns mouse capture. Preparing/focusing a TextBox for every mouse move
            // repeats text layout and input-context work before the selection has settled.
            if (Control.IsSelectionDragging) { Editor.Visibility = Visibility.Hidden; return; }
            PrepareEditor(true);
        }));
    }

    public void Bind(WorkbookSnapshot snapshot, int sheetIndex, int rows, int columns)
    {
        _editing = false; IsComposing = false; EditingChanged?.Invoke();
        // Dimension notifications can lay out the editor before the new selection is installed.
        _anchor = _current = new(0, 0);
        _snapshot = snapshot; _sheetIndex = sheetIndex;
        _model.Rows = rows; _model.Columns = columns; UpdateProjection();
        _selecting = true;
        try { Control.SetMergedModelRanges([]); Control.NotifyRefresh(); ApplyDimensions(); UpdateMergedRanges(); }
        finally { _selecting = false; }
        SelectRangeCore(new(0, 0, 0, 0), _current, false);
    }
    public void Refresh(WorkbookSnapshot snapshot)
    {
        var oldRows = _model.HiddenRows; var oldColumns = _model.HiddenColumns;
        var oldFrozenRows = _model.FrozenRows; var oldFrozenColumns = _model.FrozenColumns;
        var previous = _snapshot;
        _snapshot = snapshot; UpdateProjection();
        var rearrange = !oldRows.SetEquals(_model.HiddenRows) || !oldColumns.SetEquals(_model.HiddenColumns) ||
            !oldFrozenRows.SetEquals(_model.FrozenRows) || !oldFrozenColumns.SetEquals(_model.FrozenColumns);
        _selecting = true;
        try
        {
            if (rearrange) { Control.NotifyRowArrangeChanged(); Control.NotifyColumnArrangeChanged(); }
            if (rearrange || !ReferenceEquals(previous, snapshot)) ApplyDimensions();
            UpdateMergedRanges();
        }
        finally { _selecting = false; }
        if (!_editing)
        {
            var owner = MergeOwner(_current.Row, _current.Column);
            if (owner != _current) { _current = owner; SelectionChanged?.Invoke(); }
            _anchor = MergeOwner(_anchor.Row, _anchor.Column);
        }
        var row = View.IsFilteredOut(snapshot, _sheetIndex, _current.Row) ? View.FilterRange?.Top ?? 0 : _current.Row;
        var column = _current.Column;
        if (_model.HiddenRows.Contains(row)) row = NearestVisible(row, _model.Rows, _model.HiddenRows);
        if (_model.HiddenColumns.Contains(column)) column = NearestVisible(column, _model.Columns, _model.HiddenColumns);
        if (rearrange && !_editing && row >= 0 && column >= 0)
            SelectRangeCore(new(row, column, row, column), new(row, column), false);
        if (!ReferenceEquals(previous, snapshot)) Control.InvalidateAll();
        PrepareEditor(false); PositionOverlay();
    }
    internal void RefreshDimensions(WorkbookSnapshot snapshot) => Refresh(snapshot);
    internal bool IsFiltered(int row) => View.IsFilteredOut(_snapshot, _sheetIndex, row);
    internal string DisplayText(int row, int column) => _model.GetCellText(row, column);
    private static int NearestVisible(int index, int count, HashSet<int> hidden)
    {
        for (var i = index; i < count; i++) if (!hidden.Contains(i)) return i;
        for (var i = index - 1; i >= 0; i--) if (!hidden.Contains(i)) return i;
        return -1;
    }
    internal void ApplyDimensions(SheetAxis? previewAxis = null, int previewIndex = 0, double previewPixels = 0)
    {
        var sheet = _snapshot.Sheets[_sheetIndex];
        var rows = sheet.Rows.Where(p => p.Value.Height.HasValue && !p.Value.Hidden).ToDictionary(p => p.Key, p => Math.Clamp(p.Value.Height!.Value, 1, 10000));
        var columns = Enumerable.Range(0, _model.Columns).Where(c => sheet.Columns.GetValueOrDefault(c)?.Hidden != true).ToDictionary(c => c,
            c => sheet.Columns.GetValueOrDefault(c)?.Width is { } width ? Math.Clamp(width * 7 + 5, 1, 10000) : c == 0 ? 160d : 112d);
        if (previewAxis == SheetAxis.Row) rows[previewIndex] = previewPixels;
        if (previewAxis == SheetAxis.Column) columns[previewIndex] = previewPixels;
        Control.SetModelDimensions(rows, columns, Math.Clamp(sheet.DefaultRowHeight ?? 31, 1, 10000));
        Control.GridLineColor = Sheet.ShowGridLines ? Color.FromRgb(231, 236, 232) : Colors.White;
    }
    private void UpdateProjection()
    {
        View.RefreshFilter(_snapshot); _model.Snapshot = _snapshot; _model.SheetIndex = _sheetIndex;
        var sheet = _snapshot.Sheets[_sheetIndex];
        _model.HiddenRows = Enumerable.Range(0, _model.Rows).Where(r => sheet.Rows.GetValueOrDefault(r)?.Hidden == true || View.IsFilteredOut(_snapshot, _sheetIndex, r)).ToHashSet();
        _model.HiddenColumns = sheet.Columns.Where(p => p.Value.Hidden).Select(p => p.Key).Where(c => c < _model.Columns).ToHashSet();
        _model.FrozenRows = Enumerable.Range(0, Math.Min(View.FrozenRows, _model.Rows)).Where(r => !_model.HiddenRows.Contains(r)).ToHashSet();
        _model.FrozenColumns = Enumerable.Range(0, Math.Min(View.FrozenColumns, _model.Columns)).Where(c => !_model.HiddenColumns.Contains(c)).ToHashSet();
        Artwork.Refresh();
    }
    private void UpdateMergedRanges() => Control.SetMergedModelRanges(Sheet.MergedCells.Select(text =>
    {
        var range = CellRectangle.Parse(text);
        return new GridCellRange(range.Top, range.Left, range.Bottom, range.Right);
    }));
    public void Select(CellPosition position, bool extend = false)
    {
        if (!Commit() || BeforeNavigate?.Invoke() == false) return;
        var row = _model.HiddenRows.Contains(position.Row) ? NearestVisible(position.Row, _model.Rows, _model.HiddenRows) : position.Row;
        var column = _model.HiddenColumns.Contains(position.Column) ? NearestVisible(position.Column, _model.Columns, _model.HiddenColumns) : position.Column;
        if (row < 0 || column < 0) return;
        position = new(row, column);
        if (!extend) _anchor = MergeOwner(row, column);
        SelectRange(GridRange.Between(_anchor, position), position);
    }
    public void SelectRange(GridRange range, CellPosition current)
    {
        if (!Commit() || BeforeNavigate?.Invoke() == false) return;
        SelectRangeCore(range, current, true);
    }
    private void SelectRangeCore(GridRange range, CellPosition current, bool focus)
    {
        if (range.Top < 0 || range.Left < 0 || range.Bottom >= _model.Rows || range.Right >= _model.Columns) throw new ArgumentException("选区超出显示范围。");
        _selecting = true;
        try
        {
            if (!Control.SelectModelRange(new(range.Top, range.Left, range.Bottom, range.Right), new(current.Row, current.Column))) return;
            var selected = Control.CurrentModelCell;
            _current = MergeOwner(selected.Row!.Value, selected.Column!.Value);
        }
        finally { _selecting = false; }
        SelectionChanged?.Invoke(); PositionOverlay(); PrepareEditor(focus);
    }
    public bool Commit()
    {
        if (IsComposing) return false;
        if (!_editing) return true;
        if (!_commit(_current, Editor.Text)) return false;
        _editing = false; EditingChanged?.Invoke();
        if (LatestSnapshot is { } latest)
        {
            var snapshot = latest();
            if (!ReferenceEquals(snapshot, _snapshot))
            {
                _snapshot = snapshot; _model.Snapshot = snapshot;
                Control.InvalidateModelCell(_current.Row, _current.Column); Artwork.Refresh();
            }
        }
        PrepareEditor(false); return true;
    }
    public void Cancel() { if (IsComposing) return; _editing = false; EditingChanged?.Invoke(); PrepareEditor(true); }
    public bool BeginEdit()
    {
        if (IsComposing || BeforeBeginEdit?.Invoke() == false) return false;
        if (LatestSnapshot is { } latest) { _snapshot = latest(); _model.Snapshot = _snapshot; }
        _editing = true; SetText(CellInput.EditText(_snapshot.Sheets[_sheetIndex].GetCell(_current.Row, _current.Column)?.Content ?? CellContent.FromLiteral(CellValue.Blank)));
        EditingChanged?.Invoke(); ShowEditorDraft(); PositionOverlay(); Editor.Focus(); Editor.SelectAll(); return true;
    }
    internal void ActivateDraft()
    {
        if (_loading) return;
        _editing = true; EditingChanged?.Invoke(); ShowEditorDraft();
    }
    private void ShowEditorDraft()
    {
        Editor.Opacity = 1;
        var style = _snapshot.Styles[Sheet.GetStyleId(_current.Row, _current.Column)];
        Editor.FontFamily = FastGridArtwork.Font(style.FontName); Editor.FontSize = (style.FontSize ?? 11) * 96 / 72;
        Editor.FontWeight = style.Bold ? FontWeights.SemiBold : FontWeights.Normal;
        Editor.FontStyle = style.Italic ? FontStyles.Italic : FontStyles.Normal;
        Editor.TextWrapping = style.WrapText ? TextWrapping.Wrap : TextWrapping.NoWrap;
        Editor.TextAlignment = style.Horizontal switch { 2 => TextAlignment.Center, 3 => TextAlignment.Right, _ => TextAlignment.Left };
        Editor.VerticalContentAlignment = style.Vertical switch { 0 => VerticalAlignment.Top, 1 => VerticalAlignment.Center, _ => VerticalAlignment.Bottom };
        Editor.IsHitTestVisible = true; Editor.CaretBrush = Brushes.Black; Editor.Foreground = Brushes.Black;
        Editor.Background = Brushes.White; Editor.BorderBrush = Brushes.SeaGreen; Editor.SelectionOpacity = .4;
        Editor.Foreground = FastGridArtwork.Brush(style.FontColor, Brushes.Black);
    }
    private void SetText(string text) { _loading = true; try { CellEditorText.InitializeText(Editor, text); } finally { _loading = false; } }
    private void PrepareEditor(bool focus)
    {
        if (_snapshot is null || _editing || IsComposing) return;
        SetText(_model.GetCellText(_current.Row, _current.Column));
        Editor.IsHitTestVisible = false; Editor.CaretBrush = Brushes.Transparent; Editor.SelectionOpacity = 0;
        Editor.Foreground = Editor.Background = Editor.BorderBrush = Brushes.Transparent;
        // Keep the native input focused, but hide theme-provided focus/hover chrome too.
        // Transparent brushes alone do not suppress the Windows TextBox template's border.
        Editor.Opacity = 0;
        PositionOverlay(); if (focus && Editor.Visibility == Visibility.Visible) { Editor.Focus(); Editor.SelectAll(); }
    }
    private void PositionOverlay()
    {
        if (_snapshot is null) return;
        Rect? VisibleClip(int row, int column)
        {
            var fragments = Control.GetModelRangeFragments(new(row, column, row, column));
            return fragments.Count == 0 ? null : Rect.Intersect(fragments[0].Bounds, fragments[0].Clip);
        }
        static void ClipToCell(FrameworkElement element, Rect bounds, Rect clip)
        {
            var local = Rect.Intersect(bounds, clip);
            if (local.IsEmpty) { element.Clip = Geometry.Empty; return; }
            local.Offset(-bounds.Left, -bounds.Top);
            element.Clip = new RectangleGeometry(local);
        }
        var editorFragments = Control.GetModelRangeFragments(CurrentRange);
        if (!Control.IsSelectionDragging && editorFragments.Count > 0 && editorFragments[0].Bounds is var rect && rect.Width > 0 && rect.Height > 0)
        {
            Editor.Visibility = Visibility.Visible; Editor.Width = rect.Width; Editor.Height = rect.Height;
            Canvas.SetLeft(Editor, rect.Left); Canvas.SetTop(Editor, rect.Top);
            ClipToCell(Editor, rect, editorFragments[0].Clip);
        }
        else Editor.Visibility = Visibility.Hidden;
        var visible = new HashSet<int>();
        if (View.Region is { HasHeader: true } region)
            for (var c = region.Range.Left; c <= region.Range.Right; c++)
                if (Control.GetModelCellBounds(region.Range.Top, c) is { } bounds && bounds.Width >= 20 && bounds.Height >= 20)
                {
                    var marker = View.HeaderMarker(c);
                    var width = marker.Length > 2 ? 34d : 23d;
                    if (bounds.Width < width) continue;
                    visible.Add(c);
                    if (!_headers.TryGetValue(c, out var button))
                    {
                        var column = c; button = new Button { Width = 19, Height = 20, Padding = new(0), FontSize = 10 };
                        button.Click += (_, _) => { if (Commit()) HeaderRequested?.Invoke(column, button); };
                        _headers.Add(c, button); _overlay.Children.Add(button);
                    }
                    button.Width = width; button.Content = marker;
                    button.ToolTip = $"{MainWindow.ColumnName(c)} 列筛选和排序" + (View.Filters.ContainsKey(c) ? "（已筛选）" : "");
                    System.Windows.Automation.AutomationProperties.SetName(button, button.ToolTip.ToString());
                    Canvas.SetLeft(button, bounds.Right - width); Canvas.SetTop(button, bounds.Top);
                    ClipToCell(button, new(bounds.Right - width, bounds.Top, width, button.Height), VisibleClip(region.Range.Top, c) ?? Rect.Empty);
                }
        foreach (var column in _headers.Keys.Where(c => !visible.Contains(c)).ToArray()) { _overlay.Children.Remove(_headers[column]); _headers.Remove(column); }
    }
    private void HandleKey(object sender, KeyEventArgs e)
    {
        if (!HandlesKeyboard || e.Handled) return;
        if (IsComposing || e.Key == Key.ImeProcessed) return;
        var modifiers = Keyboard.Modifiers; var ctrl = modifiers.HasFlag(ModifierKeys.Control);
        var key = e.Key == Key.System ? e.SystemKey : e.Key;
        if (modifiers.HasFlag(ModifierKeys.Alt) && key == Key.Down) { e.Handled = true; if (Commit()) HeaderRequested?.Invoke(_current.Column, Editor); return; }
        if (key == Key.Escape) { e.Handled = true; Cancel(); return; }
        if (_editing && key is not (Key.Enter or Key.Tab)) return;
        if (ctrl && key is Key.C or Key.V or Key.Z or Key.Y || key == Key.Delete)
        { e.Handled = true; if (Commit()) CommandRequested?.Invoke(key, modifiers); return; }
        if (ctrl && key == Key.A) { e.Handled = true; SelectRange(new(0, 0, _model.Rows - 1, _model.Columns - 1), new(0, 0)); return; }
        if (key == Key.F2) { e.Handled = true; BeginEdit(); return; }
        if (key is not (Key.Enter or Key.Tab or Key.Left or Key.Right or Key.Up or Key.Down or Key.Home or Key.End)) return;
        e.Handled = true; if (!Commit()) return;
        var row = _current.Row; var column = _current.Column; var shift = modifiers.HasFlag(ModifierKeys.Shift);
        var delta = shift && key is Key.Enter or Key.Tab ? -1 : 1;
        var currentRange = CurrentRange;
        if (key == Key.Enter) row = delta > 0 ? currentRange.Bottom + 1 : currentRange.Top - 1;
        else if (key == Key.Tab) column = delta > 0 ? currentRange.Right + 1 : currentRange.Left - 1;
        else if (key == Key.Up) row = currentRange.Top - 1; else if (key == Key.Down) row = currentRange.Bottom + 1;
        else if (key == Key.Left) column = currentRange.Left - 1; else if (key == Key.Right) column = currentRange.Right + 1;
        else if (key == Key.Home) { column = 0; if (ctrl) row = 0; }
        else { column = _model.Columns - 1; if (ctrl) row = _model.Rows - 1; }
        row = Math.Clamp(row, 0, _model.Rows - 1); column = Math.Clamp(column, 0, _model.Columns - 1);
        var rd = row < _current.Row ? -1 : 1; var cd = column < _current.Column ? -1 : 1;
        while (_model.HiddenRows.Contains(row) && row + rd >= 0 && row + rd < _model.Rows) row += rd;
        while (_model.HiddenColumns.Contains(column) && column + cd >= 0 && column + cd < _model.Columns) column += cd;
        if (_model.HiddenRows.Contains(row) || _model.HiddenColumns.Contains(column)) return;
        Select(new(row, column), shift && key is not (Key.Tab or Key.Enter));
    }

    private sealed class SnapshotModel : FastGridModelBase
    {
        public WorkbookSnapshot? Snapshot; public int SheetIndex, Rows, Columns;
        public long ReadCount;
        public HashSet<int> HiddenRows = [], HiddenColumns = [], FrozenRows = [], FrozenColumns = [];
        public override int RowCount => Rows;
        public override int ColumnCount => Columns;
        public override string GetColumnHeaderText(int column) => MainWindow.ColumnName(column);
        public override string GetCellText(int row, int column)
        {
            ReadCount++;
            if (Snapshot is null) return "";
            var sheet = Snapshot.Sheets[SheetIndex];
            return CellDisplayFormatter.Format(sheet.GetCell(row, column)?.Value ?? CellValue.Blank, Snapshot.Styles[sheet.GetStyleId(row, column)], Snapshot.DateSystem);
        }
        public override IFastGridCell GetCell(IFastGridView view, int row, int column)
        {
            if (Snapshot is null) return new FastGridCellImpl();
            var sheet = Snapshot.Sheets[SheetIndex]; var style = Snapshot.Styles[sheet.GetStyleId(row, column)];
            var cell = new FastGridCellImpl { BackgroundColor = Color(style.FillColor) ?? Colors.White };
            // Artwork owns the opaque styled cell, including text. Rasterizing
            // the same text underneath it doubles formatting/glyph work and
            // creates bitmaps that cannot be seen.
            if (FastGridArtwork.PaintsStyle(style)) return cell;
            var block = cell.AddTextBlock(GetCellText(row, column)); block.IsBold = style.Bold; block.IsItalic = style.Italic; block.FontColor = Color(style.FontColor);
            return cell;
        }
        private static Color? Color(string? value) { try { return value is null ? null : (Color)ColorConverter.ConvertFromString("#" + value.TrimStart('#')); } catch (FormatException) { return null; } }
        public override HashSet<int> GetHiddenRows(IFastGridView view) => HiddenRows;
        public override HashSet<int> GetHiddenColumns(IFastGridView view) => HiddenColumns;
        public override HashSet<int> GetFrozenRows(IFastGridView view) => FrozenRows;
        public override HashSet<int> GetFrozenColumns(IFastGridView view) => FrozenColumns;
    }
}
