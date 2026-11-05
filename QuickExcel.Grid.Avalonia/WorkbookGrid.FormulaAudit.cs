using Avalonia;
using Avalonia.Media;
using QuickExcel.Domain;

namespace QuickExcel.Grid.Avalonia;

public sealed record FormulaAuditArrow(CellAddress From, CellAddress To);

public sealed partial class WorkbookGrid
{
    private WorkbookSnapshot? _auditSnapshot;
    private FormulaAuditArrow[] _auditArrows = [];
    public int FormulaAuditArrowCount => ReferenceEquals(_auditSnapshot, _snapshot) ? _auditArrows.Length : 0;

    public void SetFormulaAuditArrows(WorkbookSnapshot? snapshot, IEnumerable<FormulaAuditArrow>? arrows = null)
    {
        _auditSnapshot = snapshot; _auditArrows = arrows?.Take(50).ToArray() ?? [];
        _overlay.InvalidateVisual();
    }

    // Bounds come from the same clipped frozen/scrolling pane geometry as selection.
    // Cross-sheet and offscreen endpoints remain in the navigable audit lists.
    private void PaintFormulaAuditArrows(DrawingContext context)
    {
        if (!ReferenceEquals(_auditSnapshot, _snapshot)) return;
        var pen = new Pen(Brushes.DodgerBlue, 2 / _zoom);
        foreach (var arrow in _auditArrows)
        {
            if (arrow.From.SheetId != Sheet.Id || arrow.To.SheetId != Sheet.Id ||
                CellBounds(new(arrow.From.Row, arrow.From.Column)) is not { } source ||
                CellBounds(new(arrow.To.Row, arrow.To.Column)) is not { } target) continue;
            var from = new Point(source.Center.X / _zoom, source.Center.Y / _zoom);
            var to = new Point(target.Center.X / _zoom, target.Center.Y / _zoom);
            var delta = new Vector(to.X - from.X, to.Y - from.Y); var length = delta.Length;
            if (length < 1) { context.DrawEllipse(null, pen, to, 7 / _zoom, 7 / _zoom); continue; }
            var unit = delta / length; var end = to - unit * (8 / _zoom);
            var wing = new Vector(-unit.Y, unit.X) * (4 / _zoom);
            context.DrawEllipse(Brushes.DodgerBlue, null, from, 3 / _zoom, 3 / _zoom);
            context.DrawLine(pen, from, to); context.DrawLine(pen, to, end + wing); context.DrawLine(pen, to, end - wing);
        }
    }
}
