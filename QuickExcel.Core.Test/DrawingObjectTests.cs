using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class DrawingObjectTests
{
    [Theory]
    [InlineData(0)] [InlineData(45)] [InlineData(90)] [InlineData(270)]
    public void ResizeKeepsTheOppositeRotatedCornerFixed(double rotation)
    {
        var before = new DrawingPlacement(new(200, 200, 120, 80), rotation);
        var start = DrawingObjectInteraction.HandlePoint(before, DrawingHandle.BottomRight);
        var anchor = DrawingObjectInteraction.HandlePoint(before, DrawingHandle.TopLeft);
        var after = DrawingObjectInteraction.Update(before, DrawingHandle.BottomRight, start, new(start.X + 30, start.Y + 20));
        var fixedCorner = DrawingObjectInteraction.HandlePoint(after, DrawingHandle.TopLeft);
        Assert.Equal(anchor.X, fixedCorner.X, 8); Assert.Equal(anchor.Y, fixedCorner.Y, 8);
        Assert.True(DrawingObjectInteraction.Contains(after, DrawingObjectInteraction.HandlePoint(after, DrawingHandle.Move)));
    }
    [Fact]
    public void ConstrainedResizeAndRotationUseAspectRatioAndFifteenDegreeSteps()
    {
        var original = new DrawingPlacement(new(100, 100, 200, 100), 0);
        var start = DrawingObjectInteraction.HandlePoint(original, DrawingHandle.BottomRight);
        var after = DrawingObjectInteraction.Update(original, DrawingHandle.BottomRight, start, new(start.X + 10, start.Y + 80), true);
        Assert.Equal(2, after.Bounds.Width / after.Bounds.Height, 9);
        var knob = DrawingObjectInteraction.HandlePoint(original, DrawingHandle.Rotate); var center = DrawingObjectInteraction.HandlePoint(original, DrawingHandle.Move);
        var rotation = DrawingObjectInteraction.Update(original, DrawingHandle.Rotate, knob, DrawingObjectInteraction.Rotate(knob, center, 37), true);
        Assert.Equal(30, rotation.Rotation);
    }
    [Theory]
    [InlineData(DrawingObjectKind.Image, DrawingAnchorKind.TwoCell)] [InlineData(DrawingObjectKind.Image, DrawingAnchorKind.OneCell)]
    [InlineData(DrawingObjectKind.Chart, DrawingAnchorKind.Absolute)] [InlineData(DrawingObjectKind.Chart, DrawingAnchorKind.TwoCell)]
    public void TransformIsAtomicUndoableProtectedAndRoundTrips(DrawingObjectKind kind, DrawingAnchorKind anchorKind)
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Objects"); var original = DrawingGeometry.FromBounds(sheet, new(100, 100, 200, 100), anchorKind);
        var image = new SheetImage(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/mWQAAAAASUVORK5CYII="), "png", 0, 0, anchor: original);
        var chart = SetChartCommand.FromRange(sheet, new(0, 0, 2, 1), WorksheetChartKind.Column, original);
        sheet = kind == DrawingObjectKind.Image ? sheet.WithImages([image]) : sheet.WithData(new() { Charts = [chart] });
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]); using var session = new DocumentSession(book);
        var changed = DrawingGeometry.FromBounds(sheet, new(180, 130, 280, 140), anchorKind, 37.5); var target = new DrawingObjectId(kind, 0);
        session.Execute(new TransformDrawingObjectCommand(sheet.Id, target, original, changed));
        DrawingAnchor Anchor(WorkbookSnapshot b) => kind == DrawingObjectKind.Image ? b.Sheets[0].Images[0].Anchor : b.Sheets[0].Data.Charts[0].Anchor;
        Assert.Equal(changed, Anchor(session.Snapshot)); session.Undo(); Assert.Equal(original, Anchor(session.Snapshot)); Assert.False(session.CanUndo); session.Redo();
        var before = session.Snapshot; Assert.Throws<InvalidOperationException>(() => session.Execute(new TransformDrawingObjectCommand(sheet.Id, target, original, original with { Rotation = 90 }))); Assert.Same(before, session.Snapshot);
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../validation/drawing-objects", Guid.NewGuid().ToString("N"))); Directory.CreateDirectory(root);
        try
        {
            var storage = new TypedWorkbookStorage();
            foreach (var extension in new[] { ".xlsx", ".xlsq" }) { var path = Path.Combine(root, "rotated" + extension); storage.Save(path, session.Snapshot); var read = storage.Open(path); Assert.Empty(read.FileState!.CompatibilityIssues); Assert.Equal(changed, Anchor(read)); }
        }
        finally { Directory.Delete(root, true); }
        session.Execute(new SetProtectionCommand(sheet.Id, new() { Objects = true })); before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new TransformDrawingObjectCommand(sheet.Id, target, changed, original))); Assert.Same(before, session.Snapshot);
    }
}
