using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using QuickExcel.Core;
using SkiaSharp;

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    private IReadOnlyList<OutlineGroup> _rowOutline = [];
    private int _outlineDepth;
    private readonly Dictionary<string, Button> _outlineButtons = [];
    private double OutlineWidth => _outlineDepth == 0 ? 0 : (_outlineDepth + 1) * 22 + 6;
    public event Action<OutlineGroup>? OutlineToggleRequested;
    public event Action<int>? OutlineLevelRequested;

    private void PositionOutlineControls()
    {
        var used = new HashSet<string>();
        void Button(string key, string text, string label, double x, double y, Action action)
        {
            used.Add(key);
            if (!_outlineButtons.TryGetValue(key, out var button))
            {
                button = new Button { MinWidth = 0, MinHeight = 0, Padding = new(0), CornerRadius = new(0), BorderThickness = new(1),
                    Background = new SolidColorBrush(Color.Parse("#F1F2F3")), BorderBrush = new SolidColorBrush(Color.Parse("#9299A4")),
                    Foreground = Brushes.DimGray, HorizontalContentAlignment = HorizontalAlignment.Center, VerticalContentAlignment = VerticalAlignment.Center };
                button.Click += (_, _) => { if (!IsComposing && button.Tag is Action invoke) invoke(); };
                _outlineButtons.Add(key, button); _canvas.Children.Add(button);
            }
            button.Tag = action; button.Content = text; button.FontSize = 12 * _zoom;
            button.Width = button.Height = 19 * _zoom;
            AutomationProperties.SetName(button, label); ToolTip.SetTip(button, label);
            Canvas.SetLeft(button, x * _zoom); Canvas.SetTop(button, y * _zoom);
        }
        if (_rowOutline.Count > 0)
        {
            for (var level = 1; level <= _outlineDepth + 1; level++)
            { var value = level; Button("level:" + level, level.ToString(), UiText.Format($"显示大纲级别 {level}"), 3 + (level - 1) * 22, 4, () => OutlineLevelRequested?.Invoke(value)); }
            foreach (var group in _rowOutline)
            {
                if (_rows.Cell(group.Summary) is not { } row) continue;
                var y = HeaderHeight + row.Start + (row.Size - 19) / 2;
                if (y < HeaderHeight + (row.Frozen ? 0 : _rows.FrozenSize) || y + 19 > HeaderHeight + ViewHeight) continue;
                Button($"group:{group.Level}:{group.First}:{group.Last}", group.Collapsed ? "+" : "−",
                    UiText.Format($"{(group.Collapsed ? UiText.Get("展开") : UiText.Get("折叠"))}第 {group.First + 1}–{group.Last + 1} 行（{group.Level} 级）"),
                    3 + (group.Level - 1) * 22, y, () => OutlineToggleRequested?.Invoke(group));
            }
        }
        foreach (var key in _outlineButtons.Keys.Where(k => !used.Contains(k)).ToArray())
        { _canvas.Children.Remove(_outlineButtons[key]); _outlineButtons.Remove(key); }
    }

    private void PaintOutline(SKCanvas canvas)
    {
        if (OutlineWidth == 0) return;
        using var background = new SKPaint { Color = new(248, 249, 249) };
        using var line = new SKPaint { Color = new(151, 155, 160), StrokeWidth = 1, IsAntialias = false };
        canvas.DrawRect(SKRect.Create(0, 0, (float)OutlineWidth, (float)(HeaderHeight + ViewHeight)), background);
        canvas.DrawLine((float)OutlineWidth, 0, (float)OutlineWidth, (float)(HeaderHeight + ViewHeight), line);
        canvas.Save(); canvas.ClipRect(SKRect.Create(0, (float)HeaderHeight, (float)OutlineWidth, (float)ViewHeight));
        foreach (var group in _rowOutline.Where(g => !g.Collapsed))
        {
            var x = (float)(12.5 + (group.Level - 1) * 22);
            foreach (var fragment in _rows.Project(group.First, group.Last, ViewHeight))
            {
                canvas.Save(); canvas.ClipRect(SKRect.Create(0, (float)(HeaderHeight + fragment.ClipStart), (float)OutlineWidth, (float)fragment.ClipSize));
                var top = (float)(HeaderHeight + fragment.Start); var bottom = (float)(top + fragment.Size);
                canvas.DrawLine(x, top, x, bottom, line);
                var edge = Sheet.Data.Outline?.SummaryBelow == false ? bottom : top;
                canvas.DrawLine(x, edge, x + 6, edge, line); canvas.Restore();
            }
        }
        foreach (var row in _rows.Visible(ViewHeight))
        {
            var level = Sheet.Rows.GetValueOrDefault(row.ModelIndex)?.OutlineLevel ?? 0;
            if (level <= 0) continue;
            var y = HeaderHeight + row.Start + row.Size / 2;
            if (y < HeaderHeight + (row.Frozen ? 0 : _rows.FrozenSize)) continue;
            canvas.DrawRect((float)(12 + level * 22), (float)y - 1, 2, 2, line);
        }
        canvas.Restore();
    }
}
