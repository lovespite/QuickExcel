using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop;

/// <summary>View-only boundary. All writes remain in EditorController/DocumentSession.</summary>
internal interface IGridSurface
{
    void Bind(WorkbookSnapshot snapshot, int sheetIndex, int rows, int columns);
    void Refresh(WorkbookSnapshot snapshot);
    IReadOnlyList<CellPosition> Selection { get; }
    void Select(CellPosition position, bool extend = false);
    void SelectRange(GridRange range, CellPosition current);
    bool Commit();
    void Cancel();
    bool BeginEdit();
}

internal sealed class WpfGridSurface : IGridSurface
{
    private readonly DataGrid _grid;
    private readonly HashSet<DataGridRow> _realized = [];
    private CellPosition? _anchor;
    private int _sheetIndex;
    private SheetId? _sheetId;
    public WorkbookSnapshot Snapshot { get; private set; } = null!;
    public WorksheetSnapshot Sheet => Snapshot.Sheets[_sheetIndex];
    public VirtualRowList<SheetRow> Rows { get; private set; } = new(0, _ => throw new InvalidOperationException(), row => row.Index);
    public int RealizedRows => _realized.Count;
    public int LastRefreshRows { get; private set; }
    public SheetViewState View { get; set; } = new();
    public WpfGridSurface(DataGrid grid)
    {
        _grid = grid;
        grid.LoadingRow += (_, e) => { _realized.Add(e.Row); RefreshRow(e.Row); };
        grid.UnloadingRow += (_, e) => _realized.Remove(e.Row);
    }

    public void Bind(WorkbookSnapshot snapshot, int sheetIndex, int rows, int columns)
    {
        Snapshot = snapshot;
        _sheetIndex = sheetIndex;
        var newSheet = _sheetId != Sheet.Id;
        var rebuildColumns = newSheet || _grid.Columns.Count != columns;
        _sheetId = Sheet.Id;
        _grid.GridLinesVisibility = Sheet.ShowGridLines ? DataGridGridLinesVisibility.All : DataGridGridLinesVisibility.None;
        if (newSheet || Rows.Count != rows)
        {
            _anchor = null;
            Rows = new(rows, r => new SheetRow(this, r), row => row.Index);
            _grid.ItemsSource = Rows;
        }
        if (rebuildColumns)
        {
            _grid.Columns.Clear();
            for (var column = 0; column < columns; column++)
            {
                var style = new Style(typeof(TextBlock));
                style.Setters.Add(new Setter(TextBlock.PaddingProperty, new Thickness(8, 4, 5, 3)));
                style.Setters.Add(new Setter(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis));
                style.Setters.Add(new Setter(TextBlock.ForegroundProperty, new Binding { Converter = new AppearanceConverter(column, "foreground") }));
                style.Setters.Add(new Setter(TextBlock.BackgroundProperty, new Binding { Converter = new AppearanceConverter(column, "background") }));
                style.Setters.Add(new Setter(TextBlock.FontWeightProperty, new Binding { Converter = new AppearanceConverter(column, "weight") }));
                style.Setters.Add(new Setter(TextBlock.FontStyleProperty, new Binding { Converter = new AppearanceConverter(column, "italic") }));
                style.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, new Binding { Converter = new AppearanceConverter(column, "alignment") }));
                var cellStyle = new Style(typeof(DataGridCell), (Style)_grid.FindResource(typeof(DataGridCell)));
                cellStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Binding { Converter = new AppearanceConverter(column, "border") }));
                cellStyle.Setters.Add(new Setter(Control.BorderBrushProperty, new Binding { Converter = new AppearanceConverter(column, "borderColor") }));
                cellStyle.Triggers.Add(new Trigger { Property = DataGridCell.IsSelectedProperty, Value = true,
                    Setters = { new Setter(Control.BorderThicknessProperty, new Thickness(1)),
                        new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(24, 117, 83))) } });
                cellStyle.Triggers.Add(new Trigger { Property = DataGridCell.IsSelectedProperty, Value = true,
                    Setters = { new Setter(Control.BorderThicknessProperty, new Thickness(1)),
                        new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(24, 117, 83))) } });
                Sheet.Columns.TryGetValue(column, out var metadata);
                _grid.Columns.Add(new GridInputColumn(column)
                {
                    Header = MainWindow.ColumnName(column),
                    Width = metadata?.Width is { } width ? Math.Clamp(width * 7 + 5, 55, 400) : column == 0 ? 160 : 112,
                    Visibility = metadata?.Hidden == true ? Visibility.Collapsed : Visibility.Visible,
                    Binding = new Binding($"[{column}]") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.Explicit },
                    ElementStyle = style,
                    CellStyle = cellStyle,
                    EditingElementStyle = new Style(typeof(TextBox))
                    { Setters = { new Setter(TextBox.PaddingProperty, new Thickness(5, 3, 4, 2)),
                        new Setter(TextBox.ForegroundProperty, new Binding { Converter = new AppearanceConverter(column, "foreground") }),
                        new Setter(TextBox.BackgroundProperty, new Binding { Converter = new AppearanceConverter(column, "background") }),
                        new Setter(TextBox.FontWeightProperty, new Binding { Converter = new AppearanceConverter(column, "weight") }),
                        new Setter(TextBox.FontStyleProperty, new Binding { Converter = new AppearanceConverter(column, "italic") }),
                        new Setter(TextBox.TextAlignmentProperty, new Binding { Converter = new AppearanceConverter(column, "alignment") }) },
                      Triggers = { new Trigger { Property = GridInputColumn.IsPreparedProperty, Value = true,
                          Setters = { new Setter(UIElement.IsHitTestVisibleProperty, false),
                              new Setter(TextBox.CaretBrushProperty, Brushes.Transparent),
                              new Setter(TextBox.SelectionOpacityProperty, 0d),
                              new Setter(TextBox.SelectionTextBrushProperty, new Binding("Foreground") { RelativeSource = RelativeSource.Self }),
                              new Setter(Control.BorderThicknessProperty, new Thickness(0)),
                              new Setter(Control.BackgroundProperty, Brushes.Transparent),
                              new Setter(FrameworkElement.FocusVisualStyleProperty, null) } } } }
                });
            }
        }
        Refresh(snapshot);
    }

    public void Refresh(WorkbookSnapshot snapshot)
    {
        Snapshot = snapshot;
        _grid.FrozenColumnCount = Math.Min(View.FrozenColumns, _grid.Columns.Count);
        for (var column = 0; column < _grid.Columns.Count; column++)
        {
            Sheet.Columns.TryGetValue(column, out var metadata);
            _grid.Columns[column].Visibility = metadata?.Hidden == true ? Visibility.Collapsed : Visibility.Visible;
            _grid.Columns[column].Width = metadata?.Width is { } width ? Math.Clamp(width * 7 + 5, 55, 400) : column == 0 ? 160 : 112;
        }
        LastRefreshRows = 0;
        foreach (var container in _realized.ToArray())
        {
            RefreshRow(container);
            LastRefreshRows++;
        }
    }

    private void RefreshRow(DataGridRow container)
    {
        if (container.Item is not SheetRow row) return;
        container.Header = (row.Index + 1).ToString(CultureInfo.InvariantCulture);
        Sheet.Rows.TryGetValue(row.Index, out var metadata);
        container.Height = metadata?.Hidden == true || View.IsFilteredOut(Snapshot, _sheetIndex, row.Index) ? 0 : Math.Clamp(metadata?.Height ?? Sheet.DefaultRowHeight ?? 31, 0, 400);
        row.Refresh();
    }

    public IReadOnlyList<CellPosition> Selection => _grid.SelectedCells.Where(c => c.Item is SheetRow)
        // Keep one overflow sentinel; every consuming operation rejects it before writing.
        .Take(100_001).Select(c => new CellPosition(((SheetRow)c.Item).Index, c.Column.DisplayIndex)).ToArray();

    public void ResetAnchor() => _anchor = null;
    public void CaptureAnchor()
    {
        if (_anchor is null && _grid.CurrentCell.Item is SheetRow row && _grid.CurrentColumn is { } column)
            _anchor = new(row.Index, column.DisplayIndex);
    }
    public void Select(CellPosition position, bool extend = false)
    {
        var previous = _grid.CurrentCell.Item is SheetRow row && _grid.CurrentColumn is { } column
            ? new CellPosition(row.Index, column.DisplayIndex) : position;
        _anchor = extend ? _anchor ?? previous : position;
        SelectRange(GridRange.Between(_anchor.Value, position), position);
    }

    public void SelectRange(GridRange range, CellPosition current)
    {
        if (range.Top < 0 || range.Left < 0 || range.Bottom >= Rows.Count || range.Right >= _grid.Columns.Count ||
            range.Top > range.Bottom || range.Left > range.Right)
            throw new ArgumentException("选区超出当前显示范围。");
        if (range.Count > GridClipboard.MaxEditCells) throw new ArgumentException("键盘扩展选区上限为 20,000 格，请缩小范围。");
        _grid.CurrentCell = new(Rows[current.Row], _grid.Columns[current.Column]);
        // A rejected pending formula edit may have restored CurrentCell.
        if (_grid.CurrentCell.Item != Rows[current.Row] || _grid.CurrentColumn != _grid.Columns[current.Column]) return;
        _grid.SelectedCells.Clear();
        for (var r = range.Top; r <= range.Bottom; r++)
            for (var c = range.Left; c <= range.Right; c++)
                _grid.SelectedCells.Add(new(Rows[r], _grid.Columns[c]));
        _grid.ScrollIntoView(Rows[current.Row], _grid.Columns[current.Column]);
        _grid.UpdateLayout();
        if (_grid.Columns[current.Column].GetCellContent(Rows[current.Row]) is FrameworkElement content && content.Parent is DataGridCell cell)
            cell.Focus();
        else _grid.Focus();
    }

    public bool Commit() => _grid.CommitEdit(DataGridEditingUnit.Cell, true) && _grid.CommitEdit(DataGridEditingUnit.Row, true);
    public void Cancel() { _grid.CancelEdit(DataGridEditingUnit.Cell); _grid.CancelEdit(DataGridEditingUnit.Row); }
    public bool BeginEdit() => _grid.BeginEdit();
}

internal sealed class SheetRow(WpfGridSurface owner, int index) : INotifyPropertyChanged
{
    public int Index { get; } = index;
    public string GetInput(int column) => owner.Sheet.GetCell(Index, column) is { } cell ? CellInput.EditText(cell.Content) : "";
    public string this[int column]
    {
        get => CellDisplayFormatter.Format(owner.Sheet.GetCell(Index, column)?.Value ?? CellValue.Blank,
            GetStyle(column), owner.Snapshot.DateSystem);
        set { } // Explicit controller commit is the only write path.
    }
    public StyleDefinition GetStyle(int column) => owner.Snapshot.Styles[owner.Sheet.GetStyleId(Index, column)];
    public event PropertyChangedEventHandler? PropertyChanged;
    public void Refresh() => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(string.Empty));
}

internal sealed class GridInputColumn(int column) : DataGridTextColumn
{
    // Keep the native input context alive without exposing a caret, text selection or mouse
    // editor until the user explicitly edits. Hit testing stays on the DataGridCell so the
    // first click and drag/modified clicks continue to belong to grid selection.
    internal static readonly DependencyProperty IsPreparedProperty = DependencyProperty.RegisterAttached(
        "IsPrepared", typeof(bool), typeof(GridInputColumn), new PropertyMetadata(false));

    internal static void SetPrepared(TextBox editor, bool prepared) => editor.SetValue(IsPreparedProperty, prepared);

    protected override FrameworkElement GenerateEditingElement(DataGridCell cell, object dataItem)
    {
        var editor = (TextBox)base.GenerateEditingElement(cell, dataItem);
        // Load raw input before WPF focuses the TextBox and forwards the initiating key/composition.
        // PreparingCellForEdit is too late: resetting Text there discards the first typed character
        // and can reset the native IME composition. Writes still go through CellEditEnding only.
        BindingOperations.ClearBinding(editor, TextBox.TextProperty);
        InitializeText(editor, ((SheetRow)dataItem).GetInput(column));
        return editor;
    }

    internal static void InitializeText(TextBox editor, string text)
    {
        // Programmatic display/raw-value changes are not user edits and must not be reachable
        // through Ctrl+Z in this or a later cell's editor.
        editor.IsUndoEnabled = false;
        editor.Text = text;
        editor.IsUndoEnabled = true;
    }
}

internal sealed class AppearanceConverter(int column, string property) : IValueConverter
{
    private static readonly Brush Ink = FrozenBrush(Color.FromRgb(34, 53, 46));
    private readonly Dictionary<string, Brush> _brushes = [];
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var style = (value as SheetRow)?.GetStyle(column);
        return property switch
        {
            "weight" => style?.Bold == true ? FontWeights.SemiBold : FontWeights.Normal,
            "italic" => style?.Italic == true ? FontStyles.Italic : FontStyles.Normal,
            "alignment" => style?.Horizontal switch { 2 => TextAlignment.Center, 3 => TextAlignment.Right, _ => TextAlignment.Left },
            "border" => new Thickness(style?.LeftBorder > 0 ? 1 : 0, style?.TopBorder > 0 ? 1 : 0, style?.RightBorder > 0 ? 1 : 0, style?.BottomBorder > 0 ? 1 : 0),
            "borderColor" => ColorBrush(style?.BottomBorderColor, Brushes.Gray),
            "background" => ColorBrush(style?.FillColor, Brushes.Transparent),
            _ => ColorBrush(style?.FontColor, Ink)
        };
    }
    private Brush ColorBrush(string? color, Brush fallback)
    {
        if (string.IsNullOrEmpty(color)) return fallback;
        if (_brushes.TryGetValue(color, out var brush)) return brush;
        try
        {
            brush = (Brush)new BrushConverter().ConvertFromString("#" + color.TrimStart('#'))!;
            brush.Freeze();
            if (_brushes.Count < 1024) _brushes[color] = brush;
            return brush;
        }
        catch { return fallback; }
    }
    private static Brush FrozenBrush(Color color) { var brush = new SolidColorBrush(color); brush.Freeze(); return brush; }
    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;
}
