using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.Media;
using LayoutGrid = Avalonia.Controls.Grid;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class TextSplitPreview : Border
{
    internal const int SampleLimit = 200;
    private readonly LayoutGrid _headers = new();
    private readonly LayoutGrid _columns = new();
    private readonly ScrollViewer _vertical;
    private readonly ScrollViewer _horizontal;
    private readonly List<(Button Header, Border Body, TextBlock Text)> _selectable = [];
    public event Action<int>? ColumnSelected;

    public TextSplitPreview()
    {
        BorderBrush = Brushes.Gray;
        BorderThickness = new(1);
        Background = Brushes.White;
        _vertical = new ScrollViewer { Content = _columns, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto, VerticalContentAlignment = VerticalAlignment.Stretch };
        var table = new LayoutGrid { RowDefinitions = new("Auto,*") };
        table.Children.Add(_headers);
        LayoutGrid.SetRow(_vertical, 1);
        table.Children.Add(_vertical);
        _horizontal = new ScrollViewer { Content = table, HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility = ScrollBarVisibility.Disabled, HorizontalContentAlignment = HorizontalAlignment.Stretch };
        Child = _horizontal;
    }

    public void ShowRows(string[][] rows, int fieldCount, int? firstSheetRow = null, bool selectColumns = false)
    {
        _headers.Children.Clear();
        _headers.ColumnDefinitions.Clear();
        _columns.Children.Clear();
        _columns.ColumnDefinitions.Clear();
        _selectable.Clear();
        _horizontal.Offset = _vertical.Offset = default;
        var numbered = firstSheetRow.HasValue;
        for (var index = numbered ? -1 : 0; index < fieldCount; index++)
        {
            var field = index;
            var values = rows.Select((row, r) => field < 0 ? (firstSheetRow.GetValueOrDefault() + r + 1).ToString() :
                field < row.Length ? SingleLine(row[field]) : "").ToArray();
            var width = field < 0 ? 52 : Math.Clamp(20 + values.Select(v => v.Length * 11).DefaultIfEmpty().Max(), 104, 460);
            var slot = _columns.ColumnDefinitions.Count;
            // The last field fills the viewport for short rows, as in the import wizard.
            var length = field == fieldCount - 1 ? GridLength.Star : new GridLength(width);
            _columns.ColumnDefinitions.Add(new(length) { MinWidth = width });
            _headers.ColumnDefinitions.Add(new(length) { MinWidth = width });
            var text = new TextBlock { Text = string.Join("\n", values), Foreground = Brushes.Black, LineHeight = 22,
                FontFamily = new FontFamily("Consolas, Microsoft YaHei UI"), FontSize = 13,
                TextAlignment = field < 0 ? TextAlignment.Right : TextAlignment.Left };
            var body = new Border { Child = text, Padding = new(6, 2), BorderBrush = Brushes.Silver,
                BorderThickness = new(0, 0, 1, 0), ClipToBounds = true, Background = field < 0 ? Brushes.WhiteSmoke : Brushes.White };
            LayoutGrid.SetColumn(body, slot);
            _columns.Children.Add(body);
            if (!selectColumns) continue;
            var header = new Button { HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Left,
                Padding = new(6, 3), MinHeight = 26, CornerRadius = new(0), BorderThickness = new(0, 0, 1, 1), BorderBrush = Brushes.Silver };
            header.Click += (_, _) => ColumnSelected?.Invoke(field);
            body.PointerPressed += (_, _) => ColumnSelected?.Invoke(field);
            LayoutGrid.SetColumn(header, slot);
            _headers.Children.Add(header);
            _selectable.Add((header, body, text));
        }
    }

    public void SetFormats(int selected, Func<int, string> label)
    {
        for (var i = 0; i < _selectable.Count; i++)
        {
            var (header, body, text) = _selectable[i];
            header.Content = label(i);
            AutomationProperties.SetName(header, UiText.Format($"第 {i + 1} 列，{header.Content}"));
            body.Background = i == selected ? Brushes.Black : Brushes.White;
            text.Foreground = i == selected ? Brushes.White : Brushes.Black;
            header.FontWeight = i == selected ? FontWeight.Bold : FontWeight.Normal;
        }
    }

    private static string SingleLine(string value)
    {
        var display = value.Length <= 200 ? value : value[..200] + "…";
        return display.Replace("\r\n", "↵").Replace('\r', '↵').Replace('\n', '↵').Replace('\t', '⇥')
            .Replace('\u0085', '↵').Replace('\u2028', '↵').Replace('\u2029', '↵');
    }
}
