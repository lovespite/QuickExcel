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
internal sealed class FastGridSurface : IGridSurface
{
    internal FastGridControl Control { get; } = new() { IsReadOnly = true, UseSelectionOverlay = true, UseSmoothScrolling = true, AllowColumnResize = false, UseClearType = false, CellFontName = "Microsoft YaHei UI", CellFontSize = 14 };
    public Grid Element { get; } = new();
    private readonly Canvas _overlay = new() { ClipToBounds = true };
    internal TextBox Editor { get; } = new() { BorderThickness = new(1), Padding = new(3, 1, 3, 1), Visibility = Visibility.Hidden };
    private readonly SnapshotModel _model = new();
    private readonly Func<CellPosition, string, bool> _commit;
    private readonly Dictionary<int, Button> _headers = [];
    private bool _loading, _editing, _selecting, _selectionPresentationQueued;
    private CellPosition _current, _anchor;
    private WorkbookSnapshot _snapshot = null!;
    private int _sheetIndex;
    public bool IsComposing { get; private set; }
    public SheetViewState View { get; set; } = new();
    public CellPosition Current => _current;
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
        Element.Children.Add(Control); Element.Children.Add(_overlay); _overlay.Children.Add(Editor);
        Control.SelectedCellsChanged += (_, _) =>
        {
            if (_selecting || Control.CurrentModelCell is not { IsCell: true } cell) return;
            _current = new(cell.Row!.Value, cell.Column!.Value);
            if (!Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) _anchor = _current;
            QueueSelectionPresentation();
        };
        Control.LostMouseCapture += (_, _) => QueueSelectionPresentation();
        Control.ViewportChanged += (_, _) => PositionOverlay();
        Element.SizeChanged += (_, _) => PositionOverlay();
        Element.PreviewMouseDown += (_, e) =>
        {
            if (e.OriginalSource is DependencyObject target && (ReferenceEquals(target, Editor) || Editor.IsAncestorOf(target))) return;
            if (!Commit()) { e.Handled = true; Editor.Focus(); }
        };
        Editor.TextChanged += (_, _) =>
        {
            if (_loading) return;
            _editing = true; ShowEditorDraft();
        };
        Editor.AddHandler(TextCompositionManager.PreviewTextInputStartEvent, new TextCompositionEventHandler((_, _) => IsComposing = true), true);
        Editor.AddHandler(TextCompositionManager.PreviewTextInputUpdateEvent, new TextCompositionEventHandler((_, _) => IsComposing = true), true);
        Editor.AddHandler(TextCompositionManager.PreviewTextInputEvent, new TextCompositionEventHandler((_, _) =>
            Element.Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() => IsComposing = false))), true);
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
        _snapshot = snapshot; _sheetIndex = sheetIndex;
        _model.Rows = rows; _model.Columns = columns; UpdateProjection();
        Control.NotifyRefresh(); ApplyDimensions(); Select(new(0, 0));
    }
    public void Refresh(WorkbookSnapshot snapshot)
    {
        _snapshot = snapshot; UpdateProjection();
        _selecting = true;
        try { Control.NotifyRowArrangeChanged(); Control.NotifyColumnArrangeChanged(); ApplyDimensions(); }
        finally { _selecting = false; }
        var row = View.IsFilteredOut(snapshot, _sheetIndex, _current.Row) ? View.Region?.Range.Top ?? 0 : _current.Row;
        if (!_editing) Select(new(row, _current.Column));
        Control.InvalidateAll(); PrepareEditor(false); PositionOverlay();
    }
    private void ApplyDimensions()
    {
        var sheet = _snapshot.Sheets[_sheetIndex];
        Control.SetModelDimensions(sheet.Rows.Where(p => p.Value.Height.HasValue && !p.Value.Hidden).ToDictionary(p => p.Key, p => Math.Clamp(p.Value.Height!.Value, 1, 400)),
            Enumerable.Range(0, _model.Columns).Where(c => sheet.Columns.GetValueOrDefault(c)?.Hidden != true).ToDictionary(c => c,
                c => sheet.Columns.GetValueOrDefault(c)?.Width is { } width ? Math.Clamp(width * 7 + 5, 12, 400) : c == 0 ? 160d : 112d));
    }
    private void UpdateProjection()
    {
        View.RefreshFilter(_snapshot); _model.Snapshot = _snapshot; _model.SheetIndex = _sheetIndex;
        var sheet = _snapshot.Sheets[_sheetIndex];
        _model.HiddenRows = Enumerable.Range(0, _model.Rows).Where(r => sheet.Rows.GetValueOrDefault(r)?.Hidden == true || View.IsFilteredOut(_snapshot, _sheetIndex, r)).ToHashSet();
        _model.HiddenColumns = sheet.Columns.Where(p => p.Value.Hidden).Select(p => p.Key).Where(c => c < _model.Columns).ToHashSet();
        _model.FrozenRows = Enumerable.Range(0, Math.Min(View.FrozenRows, _model.Rows)).Where(r => !_model.HiddenRows.Contains(r)).ToHashSet();
        _model.FrozenColumns = Enumerable.Range(0, Math.Min(View.FrozenColumns, _model.Columns)).Where(c => !_model.HiddenColumns.Contains(c)).ToHashSet();
    }
    public void Select(CellPosition position, bool extend = false)
    {
        if (!Commit()) return;
        if (!extend) _anchor = position;
        SelectRange(GridRange.Between(_anchor, position), position);
    }
    public void SelectRange(GridRange range, CellPosition current)
    {
        if (!Commit()) return;
        if (range.Top < 0 || range.Left < 0 || range.Bottom >= _model.Rows || range.Right >= _model.Columns) throw new ArgumentException("选区超出显示范围。");
        _selecting = true;
        try
        {
            if (!Control.SelectModelRange(new(range.Top, range.Left, range.Bottom, range.Right), new(current.Row, current.Column))) return;
            var selected = Control.CurrentModelCell;
            _current = new(selected.Row!.Value, selected.Column!.Value);
        }
        finally { _selecting = false; }
        SelectionChanged?.Invoke(); PositionOverlay(); PrepareEditor(true);
    }
    public bool Commit()
    {
        if (IsComposing) return false;
        if (!_editing) return true;
        if (!_commit(_current, Editor.Text)) return false;
        _editing = false; PrepareEditor(false); return true;
    }
    public void Cancel() { if (IsComposing) return; _editing = false; PrepareEditor(true); }
    public bool BeginEdit()
    {
        if (IsComposing) return false;
        _editing = true; SetText(CellInput.EditText(_snapshot.Sheets[_sheetIndex].GetCell(_current.Row, _current.Column)?.Content ?? CellContent.FromLiteral(CellValue.Blank)));
        ShowEditorDraft(); PositionOverlay(); Editor.Focus(); Editor.SelectAll(); return true;
    }
    private void ShowEditorDraft()
    {
        Editor.IsHitTestVisible = true; Editor.CaretBrush = Brushes.Black; Editor.Foreground = Brushes.Black;
        Editor.Background = Brushes.White; Editor.BorderBrush = Brushes.SeaGreen; Editor.SelectionOpacity = .4;
    }
    private void SetText(string text) { _loading = true; try { GridInputColumn.InitializeText(Editor, text); } finally { _loading = false; } }
    private void PrepareEditor(bool focus)
    {
        if (_snapshot is null || _editing || IsComposing) return;
        SetText(_model.GetCellText(_current.Row, _current.Column));
        Editor.IsHitTestVisible = false; Editor.CaretBrush = Brushes.Transparent; Editor.SelectionOpacity = 0;
        Editor.Foreground = Editor.Background = Editor.BorderBrush = Brushes.Transparent;
        PositionOverlay(); if (focus && Editor.Visibility == Visibility.Visible) { Editor.Focus(); Editor.SelectAll(); }
    }
    private void PositionOverlay()
    {
        if (_snapshot is null) return;
        if (!Control.IsSelectionDragging && Control.GetModelCellBounds(_current.Row, _current.Column) is { } rect && rect.Width > 0 && rect.Height > 0)
        {
            Editor.Visibility = Visibility.Visible; Editor.Width = rect.Width; Editor.Height = rect.Height;
            Canvas.SetLeft(Editor, rect.Left); Canvas.SetTop(Editor, rect.Top);
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
                }
        foreach (var column in _headers.Keys.Where(c => !visible.Contains(c)).ToArray()) { _overlay.Children.Remove(_headers[column]); _headers.Remove(column); }
    }
    private void HandleKey(object sender, KeyEventArgs e)
    {
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
        if (key == Key.Enter) row += delta; else if (key == Key.Tab) column += delta;
        else if (key == Key.Up) row--; else if (key == Key.Down) row++;
        else if (key == Key.Left) column--; else if (key == Key.Right) column++;
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
