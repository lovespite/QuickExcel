namespace QuickExcel.Rendering;

public static class BorderMetrics
{
    public static float Width(int style) => style switch { 7 => .5f, 2 or 8 or 10 or 12 or 13 => 2, 5 => 3, _ => 1 };
    public static float[]? Dashes(int style) => style switch
    {
        3 => [4, 3], 4 => [1, 2], 8 => [8, 4], 9 => [6, 3, 1, 3], 10 => [8, 3, 2, 3],
        11 => [6, 3, 1, 3, 1, 3], 12 => [8, 3, 2, 3, 2, 3], 13 => [7, 3, 2, 1], _ => null
    };
}
