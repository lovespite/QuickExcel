namespace QuickExcel.Presentation;

public sealed record TextRotationPreset(string Name, int Rotation);

/// <summary>UI angles are positive counterclockwise; stored angles follow XLSX alignment.</summary>
public static class TextRotationOptions
{
    public static IReadOnlyList<TextRotationPreset> Presets { get; } = Array.AsReadOnly(new[]
    {
        new TextRotationPreset("逆时针角度", 45),
        new TextRotationPreset("顺时针角度", 135),
        new TextRotationPreset("竖排文字", 255),
        new TextRotationPreset("向上旋转文字", 90),
        new TextRotationPreset("向下旋转文字", 180),
        new TextRotationPreset("水平文字", 0)
    });

    public static int ToDegrees(int rotation) => rotation == 255 ? 0 : rotation <= 90 ? rotation : 90 - rotation;
    public static int FromDegrees(int degrees)
    {
        if (degrees is < -90 or > 90) throw new ArgumentOutOfRangeException(nameof(degrees));
        return degrees < 0 ? 90 - degrees : degrees;
    }
}
