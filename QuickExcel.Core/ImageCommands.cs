using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Insert, replace (including move/resize), or delete one floating image as one history entry.</summary>
public sealed class SetImageCommand(SheetId sheetId, int index, SheetImage? image) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        if (index < 0 || index > sheet.Images.Length || image is null && index == sheet.Images.Length) throw new ArgumentOutOfRangeException(nameof(index));
        if (image?.PlaceInCell == true) throw new NotSupportedException("此命令仅支持浮动图片。");
        if (image is not null && image.Anchor.Bounds(sheet) is { Width: > 0, Height: > 0 } bounds)
            _ = DrawingGeometry.FromBounds(sheet, bounds, DrawingAnchorKind.TwoCell);
        var images = image is null ? sheet.Images.RemoveAt(index) : index == sheet.Images.Length ? sheet.Images.Add(image) : sheet.Images.SetItem(index, image);
        if (images.Length > 1024 || images.Sum(i => (long)i.Data.Length) > 64 * 1024 * 1024) throw new ArgumentException("每表图片最多 1024 张、共 64 MiB。");
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? s.WithImages(images) : s)), [], true, CalculationUnchanged: true);
    }
}
