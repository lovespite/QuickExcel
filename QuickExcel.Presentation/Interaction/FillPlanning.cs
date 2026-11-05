using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Presentation;

/// <summary>Double-click fill follows the left adjacent run, then the right, without overwriting content.</summary>
public static class FillPlanning
{
    public static GridRange? DownExtent(WorksheetSnapshot sheet, GridRange source)
    {
        if (source.Top < 0 || source.Left < 0 || source.Bottom < source.Top || source.Right < source.Left ||
            source.Bottom >= WorkbookLimits.MaxRows || source.Right >= WorkbookLimits.MaxColumns || source.Count > WorkbookClipboard.MaxCells)
            throw new ArgumentException(UiText.Get("填充区域必须是不超过 20,000 格的连续矩形。"));
        bool HasContent(int row, int column) => sheet.GetCell(row, column) is { } cell &&
            (cell.Content.Kind == CellContentKind.Formula || cell.Value.Kind != CellValueKind.Blank);
        foreach (var column in new[] { source.Left - 1, source.Right + 1 })
        {
            if (column < 0 || column >= WorkbookLimits.MaxColumns || !HasContent(source.Bottom, column)) continue;
            var bottom = source.Bottom;
            while (bottom + 1 < WorkbookLimits.MaxRows && HasContent(bottom + 1, column))
            {
                var row = bottom + 1;
                if (Enumerable.Range(source.Left, source.Columns).Any(c => HasContent(row, c))) break;
                if ((long)(row - source.Top + 1) * source.Columns > WorkbookClipboard.MaxCells)
                    throw new ArgumentException(UiText.Get("自动填充范围超过 20,000 格，请改为拖动填充较小范围。"));
                bottom = row;
            }
            if (bottom > source.Bottom) return source with { Bottom = bottom };
        }
        return null;
    }
}
