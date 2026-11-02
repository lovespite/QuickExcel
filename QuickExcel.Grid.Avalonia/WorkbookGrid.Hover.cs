using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    private CellPosition? _hoverCell;
    private static readonly Cursor ResizeRowCursor = new(StandardCursorType.SizeNorthSouth), ResizeColumnCursor = new(StandardCursorType.SizeWestEast);
    private void UpdateHover(Point point)
    {
        _canvas.Cursor = ResizeHit(point) is { } hit ? hit.Axis == SheetAxis.Row ? ResizeRowCursor : ResizeColumnCursor : Cursor.Default;
        var cell = Hit(point); if (cell == _hoverCell) return; _hoverCell = cell;
        ToolTip.SetTip(_canvas, cell is { } position ? CellAnnotation(position) : null);
    }
    internal string? CellAnnotation(CellPosition position)
    {
        if (_snapshot is null) return null;
        var notes = new List<string>();
        if (Sheet.Annotations.Notes.TryGetValue(position, out var note)) notes.Add(note.Author + "：" + note.Text.Text);
        if (Sheet.Annotations.Links.TryGetValue(position, out var link)) notes.Add("链接（不会自动打开）：" + (link.Tooltip ?? link.Display ?? "") + "\n" + (link.Target ?? "") + (link.Location is null ? "" : "#" + link.Location));
        if (notes.Count == 0) return null; var text = string.Join("\n", notes); return text.Length <= 2048 ? text : text[..2048] + "…";
    }
}
