using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;

namespace QuickExcel.Desktop.Avalonia;

internal static class RibbonIcons
{
    // Created on the UI thread and reused; individual buttons always own separate controls.
    private static readonly Dictionary<RibbonIcon, Geometry> Geometries = [];
    private static readonly FuncControlTemplate<PathIcon> Template = new((icon, _) => new global::Avalonia.Controls.Shapes.Path
    {
        Width = 20, Height = 20, Stretch = Stretch.None, Data = icon.Data,
        [!Shape.FillProperty] = icon[!PathIcon.ForegroundProperty]
    });

    public static PathIcon Create(RibbonIcon icon)
    {
        if (!Geometries.TryGetValue(icon, out var geometry))
            Geometries.Add(icon, geometry = Geometry.Parse(RibbonIconData.Path(icon)));
        // Keep the SVG's complete 20x20 viewBox, including its optical margins. The default
        // PathIcon template stretches tight geometry bounds, making narrow glyphs too large.
        return new PathIcon
        {
            Classes = { "fluent-ribbon-icon" },
            Data = geometry, Width = 20, Height = 20, Template = Template,
            VerticalAlignment = VerticalAlignment.Center, IsHitTestVisible = false, Focusable = false
        };
    }
}
