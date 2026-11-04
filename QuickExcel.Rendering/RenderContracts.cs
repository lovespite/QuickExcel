using QuickExcel.Domain;

namespace QuickExcel.Rendering;

/// <summary>All reusable geometry is in 1/96 inch units. Only the host applies screen DPI/zoom.</summary>
public readonly record struct RenderRect(double X, double Y, double Width, double Height)
{
    public double Right => X + Width;
    public double Bottom => Y + Height;
}

public enum RenderOutput { Screen, Print }
public sealed record FontRequest(string? Family, string? EastAsianFamily, double Points, bool Bold = false, bool Italic = false);
public sealed record RenderDiagnostic(string Code, string Location, string Message);
public sealed record TextLayoutRequest(string Text, StyleDefinition Style, double Width = double.PositiveInfinity,
    RichCellText? RichText = null, double DefaultFontSize = 14, int FillOffset = -1, char? FillCharacter = null);
public sealed record TextMetrics(double Width, double Height, double RotatedWidth, double RotatedHeight, int Lines);
public sealed record CellRenderInput(RenderRect Bounds, TextLayoutRequest Text, bool GridLines = false, bool Note = false, ConditionalCellVisual? Conditional = null);

/// <summary>Budgets apply to each renderer/task, including native pixel memory.</summary>
public sealed record RenderBudget
{
    public long MaxPdfBytes { get; init; } = 128L * 1024 * 1024;
    public int MaxImagePixels { get; init; } = 16_000_000;
    public long MaxImageCacheBytes { get; init; } = 64L * 1024 * 1024;
    public long MaxPreviewBytes { get; init; } = 64L * 1024 * 1024;
    public int MaxPageCells { get; init; } = 100_000;
    public int MaxFontFaces { get; init; } = 128;
    public int MaxDiagnostics { get; init; } = 1024;
}
