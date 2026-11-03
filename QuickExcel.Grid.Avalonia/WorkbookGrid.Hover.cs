using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Grid.Avalonia;

public sealed partial class WorkbookGrid
{
    private string? _hoverText;
    private static readonly Cursor ResizeRowCursor = new(StandardCursorType.SizeNorthSouth), ResizeColumnCursor = new(StandardCursorType.SizeWestEast);
    private void UpdateHover(Point point)
    {
        _canvas.Cursor = ResizeHit(point) is { } hit ? hit.Axis == SheetAxis.Row ? ResizeRowCursor : ResizeColumnCursor : Cursor.Default;
        var text = HeaderAnnotation(point) ?? (Hit(point) is { } position ? CellAnnotation(position) : null);
        if (text == _hoverText) return; _hoverText = text;
        ToolTip.SetTip(_canvas, text);
    }
    internal string? HeaderAnnotation(Point point)
    {
        if (HeaderHit(point) is not { } header) return null;
        var row = header.Axis == SheetAxis.Row;
        if (!(row ? Sheet.Printing.RowBreaks : Sheet.Printing.ColumnBreaks).Any(b => b.Manual && b.Id == header.Index)) return null;
        return row ? $"第 {header.Index + 1} 行前有手动分页符；右键行标可删除行分页符。"
            : $"{ColumnName(header.Index)} 列前有手动分页符；右键列标可删除列分页符。";
    }
    internal string? CellAnnotation(CellPosition position)
    {
        if (_snapshot is null) return null;
        var notes = new List<string>();
        if (Sheet.Annotations.Notes.TryGetValue(position, out var note)) notes.Add(note.Author + "：" + note.Text.Text);
        if (Sheet.Annotations.Links.TryGetValue(position, out var link)) notes.Add("链接（不会自动打开）：" + (link.Tooltip ?? link.Display ?? "") + "\n" + (link.Target ?? "") + (link.Location is null ? "" : "#" + link.Location));
        if (notes.Count == 0) return null; var text = string.Join("\n", notes); return text.Length <= 2048 ? text : text[..2048] + "";
    }
}
