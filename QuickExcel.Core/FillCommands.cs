using System.Collections.Immutable;
using QuickExcel.Domain;

namespace QuickExcel.Core;

/// <summary>Extends a captured rectangle in one direction, repeating content/styles with A1 copy semantics.</summary>
public sealed class FillCellsCommand(SheetId sheetId, CellRectangle source, CellRectangle extent) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorkbookClipboard.ValidateRange(source); WorkbookClipboard.ValidateRange(extent);
        var vertical = source.Left == extent.Left && source.Right == extent.Right &&
            (source.Top == extent.Top && source.Bottom <= extent.Bottom || source.Bottom == extent.Bottom && source.Top >= extent.Top);
        var horizontal = source.Top == extent.Top && source.Bottom == extent.Bottom &&
            (source.Left == extent.Left && source.Right <= extent.Right || source.Right == extent.Right && source.Left >= extent.Left);
        if (!vertical && !horizontal) throw new ArgumentException("填充只能沿一个方向扩展原矩形，不能修改原选区以外的另一方向。");
        if (extent == source) return new(snapshot, [], false);
        var sheet = snapshot.GetSheet(sheetId); var captured = WorkbookClipboard.Capture(snapshot, sheetId, source);
        WorkbookClipboard.RequirePlainGeometry(sheet, extent); WorkbookClipboard.RequireVisibleDimensions(sheet, extent);
        if (captured.HasConditionalFormats || sheet.Data.ConditionalFormats.Any(r => r.Ranges.Any(a => ValidationSafety.Overlaps(a, extent))))
            throw new NotSupportedException("填充涉及条件格式，尚不能复制或拆分规则；操作未应用。请使用只贴值/公式或单独设置规则。");
        if (sheet.Annotations.Links.Keys.Concat(sheet.Annotations.Notes.Keys).Any(p => extent.Contains(p.Row, p.Column)))
            throw new NotSupportedException("填充区域含超链接/批注，尚不能安全复制附件；整批操作未应用。");
        static int Mod(int value, int divisor) => (value % divisor + divisor) % divisor;
        var edits = new List<CellEdit>();
        for (var row = extent.Top; row <= extent.Bottom; row++)
        for (var column = extent.Left; column <= extent.Right; column++)
        {
            if (source.Contains(row, column)) continue;
            var sr = Mod(row - source.Top, captured.Rows); var sc = Mod(column - source.Left, captured.Columns);
            var content = captured.Cells[sr * captured.Columns + sc].Content;
            if (content.Kind == CellContentKind.Formula) content = CellContent.FromFormula(PasteCellsCommand.CopyFormula(snapshot, sheetId, sheetId,
                content.FormulaText, new(source.Top + sr, source.Left + sc), new(row, column)), content.StyleId);
            edits.Add(new(new(sheetId, row, column), content));
        }
        return new SetCellsCommand(edits).Prepare(snapshot);
    }
}

public enum CellSeriesKind { Numbers, Days, Months, Years }

/// <summary>Explicit linear/date series from the first numeric literal. Never infers numbers or dates from text IDs.</summary>
public sealed class FillSeriesCommand(SheetId sheetId, CellRectangle range, CellSeriesKind kind, double step) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        WorkbookClipboard.ValidateRange(range);
        if (!Enum.IsDefined(kind) || !double.IsFinite(step)) throw new ArgumentException("序列类型或步长无效。");
        if (range.Top != range.Bottom && range.Left != range.Right) throw new ArgumentException("序列需要单行或单列选区。");
        if (kind != CellSeriesKind.Numbers && (step != Math.Truncate(step) || Math.Abs(step) > 100_000))
            throw new ArgumentException("日期序列步长须为 -100,000 至 100,000 的整数。");
        var sheet = snapshot.GetSheet(sheetId);
        WorkbookClipboard.RequirePlainGeometry(sheet, range); WorkbookClipboard.RequireVisibleDimensions(sheet, range);
        if (sheet.Data.ConditionalFormats.Any(r => r.Ranges.Any(a => ValidationSafety.Overlaps(a, range))))
            throw new NotSupportedException("序列填充涉及条件格式，尚不能复制或拆分规则；操作未应用。");
        var first = sheet.GetCell(range.Top, range.Left)?.Content;
        if (first?.Kind != CellContentKind.Literal || first.LiteralValue.Kind != CellValueKind.Number)
            throw new ArgumentException("序列首格必须是数字或日期序号的字面值；文本 ID、公式和空白不会被自动转换。");
        var count = Math.Max(range.Bottom - range.Top, range.Right - range.Left) + 1;
        var edits = new List<CellEdit>(); var start = first.LiteralValue.NumberValue;
        for (var index = 1; index < count; index++)
        {
            var value = SeriesValue(start, step, index, kind, snapshot.DateSystem);
            edits.Add(new(new(sheetId, range.Top + (range.Top == range.Bottom ? 0 : index), range.Left + (range.Left == range.Right ? 0 : index)),
                CellContent.FromLiteral(CellValue.FromNumber(value), sheet.GetStyleId(range.Top, range.Left))));
        }
        return new SetCellsCommand(edits).Prepare(snapshot);
    }

    public static double SeriesValue(double start, double step, int index, CellSeriesKind kind, ExcelDateSystem system)
    {
        if (!Enum.IsDefined(kind) || !Enum.IsDefined(system) || !double.IsFinite(start) || !double.IsFinite(step) || index < 0 || index >= WorkbookClipboard.MaxCells)
            throw new ArgumentException("序列参数无效。");
        if (kind == CellSeriesKind.Numbers)
        {
            var number = start + index * step;
            if (!double.IsFinite(number)) throw new ArgumentException("序列超出有限数字范围，整批操作未应用。");
            return number;
        }
        if (step != Math.Truncate(step) || Math.Abs(step) > 100_000) throw new ArgumentException("日期步长须为有界整数。");
        bool Valid(double serial) => serial >= 0 && (new ExcelSerialDate(serial, system).IsFictitiousLeapDay || new ExcelSerialDate(serial, system).TryToDateTime(out _));
        if (!Valid(start)) throw new ArgumentException("日期序列首值超出有效范围。");
        if (kind == CellSeriesKind.Days)
        {
            var day = start + step * index;
            if (!Valid(day)) throw new ArgumentException("日期序列越界，整批操作未应用。");
            return day;
        }
        var serialDate = new ExcelSerialDate(start, system);
        int year, month, originalDay;
        if (serialDate.IsFictitiousLeapDay) { year = 1900; month = 2; originalDay = 29; }
        else
        {
            serialDate.TryToDateTime(out var date); year = date.Year; month = date.Month; originalDay = date.Day;
            if (system == ExcelDateSystem.Windows1900 && start < 1) throw new ArgumentException("按月/年序列需要有效日历日期，1900 日期序号 0 仅支持按日序列。");
        }
        var months = (long)(year - 1) * 12 + month - 1 + checked((long)step * index * (kind == CellSeriesKind.Years ? 12 : 1));
        if (months is < 0 or >= 9999 * 12) throw new ArgumentException("日期序列越界，整批操作未应用。");
        year = (int)(months / 12) + 1; month = (int)(months % 12) + 1;
        var fakeFebruary = system == ExcelDateSystem.Windows1900 && year == 1900 && month == 2;
        var dayOfMonth = Math.Min(originalDay, fakeFebruary ? 29 : DateTime.DaysInMonth(year, month));
        var result = (fakeFebruary && dayOfMonth == 29 ? 60 : ExcelSerialDate.FromDateTime(new(year, month, dayOfMonth), system).Serial) + (start - Math.Floor(start));
        if (!Valid(result)) throw new ArgumentException("日期序列越界，整批操作未应用。");
        return result;
    }
}

public sealed class ClearCellFormatsCommand : DocumentCommand
{
    private readonly ImmutableArray<CellAddress> _addresses;
    public ClearCellFormatsCommand(IEnumerable<CellAddress> addresses)
    {
        _addresses = addresses.Take(WorkbookClipboard.MaxCells + 1).ToImmutableArray();
        if (_addresses.Length > WorkbookClipboard.MaxCells) throw new ArgumentException("单次清除格式上限为 20,000 格。");
    }
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var edits = new List<CellEdit>();
        foreach (var address in _addresses.Distinct())
        {
            var sheet = snapshot.GetSheet(address.SheetId); var range = new CellRectangle(address.Row, address.Column, address.Row, address.Column);
            if (sheet.Data.ConditionalFormats.Any(r => r.Contains(address.Row, address.Column)))
                throw new NotSupportedException("清除格式涉及条件格式；请先在条件格式管理器中调整完整规则，操作未应用。");
            WorkbookClipboard.RequirePlainGeometry(sheet, range); WorkbookClipboard.RequireVisibleDimensions(sheet, range);
            var content = sheet.GetCell(address.Row, address.Column)?.Content ?? CellContent.FromLiteral(CellValue.Blank);
            // Excel ClearFormats resets the cell XF, retaining explicit rich-text run fonts.
            edits.Add(new(address, content.WithStyle(0)));
        }
        return new SetCellsCommand(edits).Prepare(snapshot) with { CalculationUnchanged = true };
    }
}
