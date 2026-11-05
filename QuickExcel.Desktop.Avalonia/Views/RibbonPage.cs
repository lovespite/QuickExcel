using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.LogicalTree;
using Avalonia.VisualTree;

namespace QuickExcel.Desktop.Avalonia;

/// <summary>Tries labeled commands before compact commands, then lets the viewer scroll.</summary>
internal sealed class RibbonPage : ScrollViewer
{
    private readonly StackPanel _groups = new() { Orientation = Orientation.Horizontal };
    private readonly List<RibbonCommandContent> _commands = [];
    private readonly bool _allowLabels;

    protected override Type StyleKeyOverride => typeof(ScrollViewer);
    public bool CommandLabelsVisible { get; private set; }

    public RibbonPage(bool allowLabels)
    {
        _allowLabels = allowLabels;
        Content = _groups;
        VerticalAlignment = VerticalAlignment.Stretch;
        HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
        VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
    }

    public void AddGroup(Control group)
    {
        _groups.Children.Add(group);
        _commands.AddRange(group.GetLogicalDescendants().OfType<RibbonCommandContent>());
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        if (_allowLabels)
        {
            // Measure the full labels every time, independently of the previous compact
            // extent. This restores labels on growth and avoids threshold oscillation.
            SetLabels(true);
            var expanded = base.MeasureOverride(availableSize);
            var width = Math.Max(0, availableSize.Width - Padding.Left - Padding.Right - BorderThickness.Left - BorderThickness.Right);
            if (_groups.DesiredSize.Width <= width) return expanded;
            SetLabels(false);
        }
        return base.MeasureOverride(availableSize);
    }

    private void SetLabels(bool visible)
    {
        foreach (var command in _commands)
        {
            if (command.Label.IsVisible == visible) continue;
            command.Label.IsVisible = visible;
            // Avalonia normally propagates changed desired sizes in later layout passes.
            // Both measurements here must see fresh button, grid and presenter sizes now.
            command.InvalidateMeasure();
            foreach (var ancestor in command.GetVisualAncestors().TakeWhile(v => !ReferenceEquals(v, this)).OfType<Layoutable>())
                ancestor.InvalidateMeasure();
        }
        CommandLabelsVisible = visible;
    }
}

/// <summary>The same icon and title are retained in both Ribbon layouts.</summary>
internal sealed class RibbonCommandContent : StackPanel
{
    public TextBlock Label { get; }

    public RibbonCommandContent(RibbonIcon icon, string title)
    {
        Orientation = Orientation.Horizontal;
        Spacing = 6;
        Children.Add(RibbonIcons.Create(icon));
        Label = new TextBlock { Text = title, IsVisible = false, VerticalAlignment = VerticalAlignment.Center };
        Children.Add(Label);
    }
}
