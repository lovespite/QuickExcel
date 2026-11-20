using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>A whole-axis copy lease. Only the occupied rectangle is materialized.</summary>
public sealed class AxisClipboard
{
    private AxisClipboard(WorkbookClipboard cells, SheetAxis axis, int first, int count)
    { Cells = cells; Axis = axis; First = first; Count = count; }
    public WorkbookClipboard Cells { get; }
    public SheetAxis Axis { get; }
    public int First { get; }
    public int Count { get; }

    public static AxisClipboard Capture(WorkbookSnapshot snapshot, SheetId sheetId, SheetAxis axis, int first, int count)
    {
        var row = axis == SheetAxis.Row;
        var limit = row ? WorkbookLimits.MaxRows : WorkbookLimits.MaxColumns;
        if (!Enum.IsDefined(axis) || first < 0 || count < 1 || first > limit - count) throw new ArgumentOutOfRangeException(nameof(first));
        var sheet = snapshot.GetSheet(sheetId);
        if (sheet.Data.AllFilters.Any(f => f.Columns.Count > 0)) throw new InvalidOperationException(MessageText.Get("请先清除筛选条件，再复制完整行列。"));
        var full = row ? new CellRectangle(first, 0, first + count - 1, WorkbookLimits.MaxColumns - 1)
            : new CellRectangle(0, first, WorkbookLimits.MaxRows - 1, first + count - 1);
        ExtendedObjectSafety.ValidateCopy(snapshot, sheetId, full);
        WorkbookClipboard.RequireCompleteMerges(sheet, full);
        WorkbookClipboard.RequireNoDrawings(sheet, full);
        WorkbookClipboard.RequireVisibleDimensions(sheet, full);
        if (sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => full.Contains(p.Row, p.Column)) ||
            sheet.Data.Validations.Any(v => v.Ranges.Any(r => ValidationSafety.Overlaps(r, full))) ||
            sheet.Data.ConditionalFormats.Any(v => v.Ranges.Any(r => ValidationSafety.Overlaps(r, full))) ||
            sheet.Data.IgnoredErrors.Any(e => ValidationSafety.Overlaps(e.Range, full)) ||
            sheet.Data.Tables.Any(t => ValidationSafety.Overlaps(t.Range, full)) || WorksheetOutline.HasGroups(sheet))
            throw new NotSupportedException(MessageText.Get("复制整行列涉及链接、批注、验证、条件格式、忽略错误提示、表格或分组，尚不能完整复制其附件；操作未应用。"));
        var extent = 0;
        foreach (var position in sheet.Cells.Keys.Where(p => full.Contains(p.Row, p.Column)))
            extent = Math.Max(extent, row ? position.Column : position.Row);
        foreach (var merge in sheet.MergedCells.Select(CellRectangle.Parse).Where(m => ValidationSafety.Overlaps(m, full)))
            extent = Math.Max(extent, row ? merge.Right : merge.Bottom);
        // Blank cells can still inherit formatting from the perpendicular axis.
        var styled = row ? sheet.Columns.Where(p => p.Value.StyleId != 0).Select(p => p.Key)
            : sheet.Rows.Where(p => p.Value.ApplyStyle == true).Select(p => p.Key);
        extent = Math.Max(extent, styled.DefaultIfEmpty(0).Max());
        var bounded = row ? full with { Right = extent } : full with { Bottom = extent };
        return new(WorkbookClipboard.Capture(snapshot, sheetId, bounded), axis, first, count);
    }
}

/// <summary>Insertion, formula translation, copied dimensions and contents share one undo entry.</summary>
public sealed class InsertCopiedAxisCommand(AxisClipboard source, SheetId sheetId, int index) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        ArgumentNullException.ThrowIfNull(source);
        var copied = source.Cells;
        if (snapshot.DocumentId != copied.DocumentId || snapshot.Revision != copied.Revision)
            throw new InvalidOperationException(MessageText.Get("复制后文档已改变，请重新复制行列。"));
        var row = source.Axis == SheetAxis.Row;
        var original = snapshot.GetSheet(copied.SheetId);
        var target = snapshot.GetSheet(sheetId);
        if (target.Data.Protection?.Sheet == true)
            throw new InvalidOperationException(MessageText.Get("请先取消工作表保护，再插入复制的行列。"));
        if (!target.Data.Tables.IsEmpty || target.Data.AllFilters.Any(f => f.Columns.Count > 0))
            throw new NotSupportedException(MessageText.Get("请先清除筛选条件，且在不含表格的工作表插入复制的行列。"));
        if (sheetId == original.Id && index > source.First && index < source.First + source.Count)
            throw new InvalidOperationException(MessageText.Get("插入位置不能拆开复制区域，请选择复制区域之前或之后的行列。"));
        var shift = new ShiftAxisCommand(sheetId, source.Axis, index, source.Count);
        var shifted = shift.Prepare(snapshot).Snapshot;
        ProtectionSafety.Validate(snapshot, shifted, shift);
        // Read the source after insertion so absolute and relative references both follow the shift.
        var movedFirst = source.First + (sheetId == original.Id && index <= source.First ? source.Count : 0);
        var capture = AxisClipboard.Capture(shifted, original.Id, source.Axis, movedFirst, source.Count).Cells;
        var destination = shifted.GetSheet(sheetId);
        bool Inserted(CellPosition p) => (row ? p.Row : p.Column) >= index && (row ? p.Row : p.Column) < index + source.Count;
        var cells = destination.Cells.Where(p => !Inserted(p.Key)).ToImmutableDictionary();
        var rows = destination.Rows.ToBuilder(); var columns = destination.Columns.ToBuilder();
        for (var offset = 0; offset < source.Count; offset++)
        {
            if (row)
            { var metadata = original.Rows.GetValueOrDefault(source.First + offset) ?? new(); rows[index + offset] = metadata with { Height = metadata.Height ?? original.DefaultRowHeight, Descent = metadata.Descent ?? original.Data.DefaultRowDescent }; }
            else
            { var metadata = original.Columns.GetValueOrDefault(source.First + offset) ?? new(); columns[index + offset] = metadata with { Width = metadata.Width ?? original.Data.DefaultColumnWidth ?? 8.43 }; }
        }
        destination = new(destination.Id, destination.Name, cells, rows, columns, destination.MergedCells, destination.Images,
            destination.ShowGridLines, destination.DefaultRowHeight, destination.TabColor, destination.View, destination.Visibility,
            destination.Printing, destination.Annotations, destination.Data);
        shifted = shifted.WithSheets(shifted.Sheets.Select(s => s.Id == sheetId ? destination : s));
        var output = row ? new CellRectangle(index, 0, index + source.Count - 1, capture.Range.Right)
            : new CellRectangle(0, index, capture.Range.Bottom, index + source.Count - 1);
        var pasted = new PasteCellsCommand(capture, sheetId, output, snapshot.Revision).Prepare(shifted).Snapshot;
        return StructureEdits.Draft(snapshot, pasted);
    }
}
