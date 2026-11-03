using QuickExcel.Domain;

namespace QuickExcel.Core;

public enum DrawingObjectKind { Image, Chart }
public readonly record struct DrawingObjectId(DrawingObjectKind Kind, int Index);

/// <summary>One completed pointer gesture is one atomic, protection-checked history entry.</summary>
public sealed class TransformDrawingObjectCommand(SheetId sheetId, DrawingObjectId target, DrawingAnchor original, DrawingAnchor placement) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        placement.Validate(); var sheet = snapshot.GetSheet(sheetId);
        if (original == placement) return new(snapshot, [], false);
        if (target.Index < 0) throw new ArgumentOutOfRangeException(nameof(target));
        if (target.Kind == DrawingObjectKind.Image && target.Index < sheet.Images.Length)
        {
            var image = sheet.Images[target.Index];
            if (image.Anchor != original) throw new InvalidOperationException("图片已变化，请重新操作。");
            return new SetImageCommand(sheetId, target.Index, image.WithAnchor(placement)).Prepare(snapshot);
        }
        if (target.Kind == DrawingObjectKind.Chart && target.Index < sheet.Data.Charts.Length)
        {
            var chart = sheet.Data.Charts[target.Index];
            if (chart.Anchor != original) throw new InvalidOperationException("图表已变化，请重新操作。");
            return new SetChartCommand(sheetId, target.Index, chart with { Anchor = placement }).Prepare(snapshot);
        }
        throw new InvalidOperationException("绘图对象已不存在。");
    }
}
