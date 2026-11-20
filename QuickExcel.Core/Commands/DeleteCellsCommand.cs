using System.Collections.Immutable;
using QuickExcel.Domain;
using QuickExcel.Formulas;

namespace QuickExcel.Core;

public enum CellDeleteShift { Left, Up }

/// <summary>Deletes a rectangle and closes its gap, preserving cut-style references and one history entry.</summary>
public sealed class DeleteCellsCommand(SheetId sheetId, CellRectangle range, CellDeleteShift shift) : DocumentCommand
{
    private bool Horizontal => shift == CellDeleteShift.Left;
    private int Count => Horizontal ? range.Right - range.Left + 1 : range.Bottom - range.Top + 1;
    private CellRectangle Affected => Horizontal ? range with { Right = WorkbookLimits.MaxColumns - 1 }
        : range with { Bottom = WorkbookLimits.MaxRows - 1 };

    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorkbookLimits.ValidatePosition(range.Top, range.Left);
        WorkbookLimits.ValidatePosition(range.Bottom, range.Right);
        if (!Enum.IsDefined(shift) || range.Top > range.Bottom || range.Left > range.Right)
            throw new ArgumentException(MessageText.Get("删除区域或移动方向无效。"));
        var source = snapshot.GetSheet(sheetId);
        Validate(snapshot, source);

        FormulaNode Transform(FormulaNode node, SheetId context)
        {
            if (node is not ReferenceNode reference ||
                !string.Equals(reference.Sheet ?? snapshot.GetSheet(context).Name, source.Name, StringComparison.OrdinalIgnoreCase)) return node;
            // Partial cell deletion does not remove worksheet columns. A:A still denotes the complete column.
            if (reference.WholeColumns && (!Horizontal || range.Top != 0 || range.Bottom != WorkbookLimits.MaxRows - 1)) return node;
            var mapped = MapRectangle(new(reference.Row, reference.Column, reference.EndRow, reference.EndColumn));
            if (mapped is null) return new ValueNode(CellValue.FromError(CellErrorCode.Reference));
            return reference with { Row = mapped.Value.Top, Column = mapped.Value.Left, EndRow = mapped.Value.Bottom, EndColumn = mapped.Value.Right };
        }

        var merges = new List<string>();
        foreach (var text in source.MergedCells)
        {
            var original = CellRectangle.Parse(text);
            if (MapRectangle(original) is not { } mapped) continue;
            var owner = Position(new(original.Top, original.Left));
            if (owner is null || owner.Value != new CellPosition(mapped.Top, mapped.Left))
                throw new NotSupportedException(MessageText.Get("操作会移除或移动合并区域的主单元格，请先取消合并。"));
            // A one-cell remainder no longer needs a merge record.
            if (mapped.Count > 1) merges.Add(mapped.ToString());
        }
        var ignored = source.Data.IgnoredErrors.SelectMany(e => Pieces(e.Range).Select(r => e with { Range = r }))
            .Distinct().Take(4097).ToImmutableArray();
        if (ignored.Length > 4096)
            throw new NotSupportedException(MessageText.Get("操作会使忽略错误提示的区域超过 4096 个，整批操作未应用。"));

        var transformed = StructureEdits.Transform(snapshot, Transform, (sheet, cells) =>
        {
            if (sheet.Id != sheetId) return sheet.WithCells(cells);
            var relocated = ImmutableDictionary.CreateBuilder<CellPosition, CellSnapshot>();
            foreach (var (position, cell) in cells)
            {
                if (Position(position) is not { } destination) continue;
                relocated.Add(destination, cell);
            }
            MoveInheritedStyles(source, relocated);
            return new(sheet.Id, sheet.Name, relocated, sheet.Rows, sheet.Columns, merges, sheet.Images,
                sheet.ShowGridLines, sheet.DefaultRowHeight, sheet.TabColor, sheet.View, sheet.Visibility, sheet.Printing,
                sheet.Annotations, sheet.Data with { IgnoredErrors = ignored });
        });
        return StructureEdits.Draft(snapshot, transformed);
    }

    private void Validate(WorkbookSnapshot book, WorksheetSnapshot sheet)
    {
        if (sheet.Data.Protection?.Sheet == true)
            throw new InvalidOperationException(MessageText.Get("请先取消工作表保护，再删除单元格并移动区域。"));
        WorkbookClipboard.RequireVisibleDimensions(sheet, Affected);
        WorkbookClipboard.RequireNoDrawings(sheet, Affected);
        if (sheet.Printing.HasReferences)
            throw new NotSupportedException(MessageText.Get("含打印区域或重复标题的工作表暂不支持删除单元格并移动区域。"));
        if (sheet.Data.Tables.Any(t => ValidationSafety.Overlaps(t.Range, Affected)) ||
            sheet.Data.Filter is { } filter && ValidationSafety.Overlaps(filter.Range, Affected) ||
            sheet.Data.Sort is { } sort && ValidationSafety.Overlaps(sort.Range, Affected))
            throw new NotSupportedException(MessageText.Get("删除移动区域与表格、筛选或排序区域相交；请先调整或清除该结构。"));
        if (book.Sheets.Any(s => !s.Data.Validations.IsEmpty || !s.Data.ConditionalFormats.IsEmpty) ||
            book.FileState?.CompatibilityDiagnostics.Any(d => d.Feature.Contains("conditionalFormatting", StringComparison.OrdinalIgnoreCase) ||
                d.Feature.Contains("dataValidation", StringComparison.OrdinalIgnoreCase) || d.Feature.Contains("sortState", StringComparison.OrdinalIgnoreCase)) == true ||
            SubtotalFunctions.HasUnmodeledFilters(book))
            throw new NotSupportedException(MessageText.Get("工作簿含验证、条件格式或尚未建模的筛选排序规则，暂不能安全删除单元格并移动区域。"));
        if (WorksheetFilterEngine.Evaluate(book, sheet.Id, WorksheetFilterEngine.CanUseSourceCaches(book)).Any(r => r >= Affected.Top && r <= Affected.Bottom))
            throw new InvalidOperationException(MessageText.Get("删除移动区域含筛选隐藏行，请先清除筛选条件。"));
        if (book.Sheets.SelectMany(s => s.Data.Charts).SelectMany(c => c.Series)
            .SelectMany(s => new[] { s.NameReference, s.Categories, s.Values }).OfType<string>()
            .Any(f => string.Equals(WorksheetChart.ParseReference(f).Sheet, sheet.Name, StringComparison.OrdinalIgnoreCase)))
            throw new NotSupportedException(MessageText.Get("受影响工作表含图表源引用，暂不删除单元格并移动区域；请先调整图表。"));
        foreach (var (position, cell) in sheet.Cells)
        {
            if (cell.Calculation?.Array is not { } array) continue;
            var spill = new CellRectangle(position.Row, position.Column, position.Row + array.Rows - 1, position.Column + array.Columns - 1);
            if (ValidationSafety.Overlaps(spill, Affected) && !range.Contains(spill))
                throw new NotSupportedException(MessageText.Get("删除移动区域与数组溢出区域相交，请先移除整个数组公式。"));
        }
    }

    private CellPosition? Position(CellPosition position)
    {
        if (!Affected.Contains(position.Row, position.Column)) return position;
        if (range.Contains(position.Row, position.Column)) return null;
        return Horizontal ? position with { Column = position.Column - Count } : position with { Row = position.Row - Count };
    }

    // Split along the affected strip before mapping. The pieces stay disjoint, so their areas
    // prove whether a formula/merge remains a single rectangle without enumerating its cells.
    private IEnumerable<CellRectangle> Pieces(CellRectangle original)
    {
        if (!ValidationSafety.Overlaps(original, Affected)) { yield return original; yield break; }
        var band = Horizontal ? original : Transpose(original);
        var deleted = Horizontal ? range : Transpose(range);
        if (band.Top < deleted.Top) yield return Restore(band with { Bottom = deleted.Top - 1 });
        if (band.Bottom > deleted.Bottom) yield return Restore(band with { Top = deleted.Bottom + 1 });
        var top = Math.Max(band.Top, deleted.Top); var bottom = Math.Min(band.Bottom, deleted.Bottom);
        if (band.Left < deleted.Left)
            yield return Restore(new(top, band.Left, bottom, Math.Min(band.Right, deleted.Left - 1)));
        if (band.Right > deleted.Right)
            yield return Restore(new(top, Math.Max(band.Left, deleted.Right + 1) - Count, bottom, band.Right - Count));
    }

    private static CellRectangle Transpose(CellRectangle value) => new(value.Left, value.Top, value.Right, value.Bottom);
    private CellRectangle Restore(CellRectangle value) => Horizontal ? value : Transpose(value);

    private CellRectangle? MapRectangle(CellRectangle original)
    {
        var pieces = Pieces(original).ToArray();
        if (pieces.Length == 0) return null;
        var result = new CellRectangle(pieces.Min(r => r.Top), pieces.Min(r => r.Left), pieces.Max(r => r.Bottom), pieces.Max(r => r.Right));
        if (result.Count != pieces.Sum(r => r.Count))
            throw new NotSupportedException(MessageText.Get("删除会把引用或合并区域拆成不连续区域，操作未应用。请调整选区或先取消合并。"));
        return result;
    }

    private void MoveInheritedStyles(WorksheetSnapshot sheet, ImmutableDictionary<CellPosition, CellSnapshot>.Builder cells)
    {
        var indexes = Horizontal ? sheet.Columns.Where(p => p.Value.StyleId != 0).Select(p => p.Key)
            : sheet.Rows.Where(p => p.Value.ApplyStyle).Select(p => p.Key);
        var begin = Horizontal ? range.Left : range.Top;
        var end = (Horizontal ? WorkbookLimits.MaxColumns : WorkbookLimits.MaxRows) - Count;
        var candidates = indexes.SelectMany(i => new[] { i, i - Count }).Where(i => i >= begin && i < end).Distinct().ToArray();
        var perpendicularStart = Horizontal ? range.Top : range.Left;
        var perpendicularEnd = Horizontal ? range.Bottom : range.Right;
        // This bound concerns materializing implicit formatting, not the size of a sparse deletion.
        if ((long)candidates.Length * (perpendicularEnd - perpendicularStart + 1) > WorkbookLimits.MaxCells)
            throw new NotSupportedException(MessageText.Get("删除移动区域的继承格式过多，请缩小选区。"));
        var added = 0;
        foreach (var index in candidates)
        for (var cross = perpendicularStart; cross <= perpendicularEnd; cross++)
        {
            var destination = Horizontal ? new CellPosition(cross, index) : new(index, cross);
            var source = Horizontal ? destination with { Column = index + Count } : destination with { Row = index + Count };
            if (sheet.Cells.ContainsKey(source)) continue;
            var style = sheet.GetInheritedStyleId(source.Row, source.Column);
            if (style == sheet.GetInheritedStyleId(destination.Row, destination.Column)) continue;
            if (++added > WorkbookClipboard.MaxCells)
                throw new NotSupportedException(MessageText.Get("删除移动区域的继承格式过多，请缩小选区。"));
            cells[destination] = new(CellContent.FromLiteral(CellValue.Blank, style));
        }
    }
}
