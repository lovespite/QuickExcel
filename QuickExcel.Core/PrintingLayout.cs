using System.Collections.Immutable;
using System.Globalization;
using System.Text;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed record PrintPage(ImmutableArray<int> Rows, ImmutableArray<int> Columns);
public sealed record PrintLayout(double PaperWidth, double PaperHeight, PageMargins Margins, double Scale, ImmutableArray<PrintPage> Pages);

/// <summary>Bounded, DPI-independent pagination. All geometry uses 1/96-inch units.</summary>
public static class PrintingLayout
{
    public const int MaxPages = 2000;
    public const long MaxPrintedCells = 2_000_000;
    public static WorksheetSnapshot VisibleSheet(WorkbookSnapshot book, WorksheetSnapshot sheet)
    {
        var hidden = WorksheetFilterEngine.Evaluate(book, sheet.Id);
        if (hidden.IsEmpty) return sheet;
        var rows = sheet.Rows.ToBuilder();
        foreach (var r in hidden) rows[r] = (rows.GetValueOrDefault(r) ?? new RowMetadata()) with { Hidden = true };
        return new(sheet.Id, sheet.Name, sheet.Cells, rows, sheet.Columns, sheet.MergedCells, sheet.Images,
            sheet.ShowGridLines, sheet.DefaultRowHeight, sheet.TabColor, sheet.View, sheet.Visibility, sheet.Printing, sheet.Annotations, sheet.Data);
    }
    public static PrintLayout Create(WorksheetSnapshot sheet, CancellationToken cancellationToken = default)
    {
        var printing = sheet.Printing; var setup = printing.Setup ?? new();
        var (width, height) = setup.PaperSize switch
        {
            null or 9 => (210 / 25.4 * 96, 297 / 25.4 * 96), 1 => (816d, 1056d), 5 => (816d, 1344d),
            8 => (297 / 25.4 * 96, 420 / 25.4 * 96), 11 => (148 / 25.4 * 96, 210 / 25.4 * 96),
            _ when setup.PaperWidth is not null && setup.PaperHeight is not null => (Measure(setup.PaperWidth), Measure(setup.PaperHeight)),
            _ => throw new NotSupportedException("此纸张尚不能预览；请在页面设置中选择 A4、A3、A5、Letter 或 Legal。")
        };
        if (setup.Orientation == "landscape") (width, height) = (height, width);
        var margins = printing.Margins ?? new(.7, .7, .75, .75, .3, .3);
        var usableWidth = width - 96 * (margins.Left + margins.Right); var usableHeight = height - 96 * (margins.Top + margins.Bottom);
        if (usableWidth <= 1 || usableHeight <= 1) throw new ArgumentException("页边距未留下可打印区域。");
        if (setup.CellComments is "asDisplayed" or "atEnd") throw new NotSupportedException("批注打印尚未支持，请将页面批注设置为不打印。");
        var extentRows = Math.Max(1, sheet.UsedRowCount); var extentColumns = Math.Max(1, sheet.UsedColumnCount);
        foreach (var merge in sheet.MergedCells.Select(CellRectangle.Parse)) { extentRows = Math.Max(extentRows, merge.Bottom + 1); extentColumns = Math.Max(extentColumns, merge.Right + 1); }
        foreach (var anchor in sheet.Images.Select(i => i.Anchor).Concat(sheet.Data.Charts.Select(c => c.Anchor)))
        {
            var bounds = anchor.Bounds(sheet); if (bounds.Width <= 0 || bounds.Height <= 0) continue;
            var end = DrawingGeometry.FromBounds(sheet, bounds, DrawingAnchorKind.TwoCell);
            extentRows = Math.Max(extentRows, end.EndRow + (end.EndRowOffset > 0 ? 1 : 0)); extentColumns = Math.Max(extentColumns, end.EndColumn + (end.EndColumnOffset > 0 ? 1 : 0));
        }
        var areas = printing.Areas.IsEmpty ? [new PrintRange(PrintRangeKind.Cells, 0, 0, extentRows - 1, extentColumns - 1)] : printing.Areas;
        var normalized = areas.Select(a => new CellRectangle(a.Kind == PrintRangeKind.Columns ? 0 : a.Top, a.Kind == PrintRangeKind.Rows ? 0 : a.Left,
            a.Kind == PrintRangeKind.Columns ? extentRows - 1 : a.Bottom, a.Kind == PrintRangeKind.Rows ? extentColumns - 1 : a.Right)).ToArray();
        foreach (var a in normalized)
        {
            WorkbookLimits.ValidatePosition(a.Top, a.Left); WorkbookLimits.ValidatePosition(a.Bottom, a.Right);
            if (a.Count > MaxPrintedCells) throw new NotSupportedException($"单个打印区域超过 {MaxPrintedCells:N0} 格，请缩小打印区域。");
        }
        var repeatRows = printing.Titles.FirstOrDefault(t => t.Kind == PrintRangeKind.Rows);
        var repeatCols = printing.Titles.FirstOrDefault(t => t.Kind == PrintRangeKind.Columns);
        ImmutableArray<int> Visible(int first, int last, bool row) => Enumerable.Range(first, last - first + 1).Where(i => (row ? DrawingGeometry.RowSize(sheet, i) : DrawingGeometry.ColumnSize(sheet, i)) > 0).ToImmutableArray();
        var titleRows = repeatRows is null ? [] : Visible(repeatRows.Top, Math.Min(repeatRows.Bottom, WorkbookLimits.MaxRows - 1), true);
        var titleCols = repeatCols is null ? [] : Visible(repeatCols.Left, Math.Min(repeatCols.Right, WorkbookLimits.MaxColumns - 1), false);
        if (titleRows.Length > 256 || titleCols.Length > 64) throw new NotSupportedException("重复标题最多 256 行、64 列。");

        List<ImmutableArray<int>> Split(int first, int last, bool row, double available, ImmutableArray<int> titles, IEnumerable<PrintPageBreak> breaks)
        {
            var titleSize = titles.Sum(i => row ? DrawingGeometry.RowSize(sheet, i) : DrawingGeometry.ColumnSize(sheet, i));
            if (titleSize >= available) throw new ArgumentException("重复标题超过可打印页大小，请缩小标题或缩放比例。");
            var cuts = breaks.Where(b => b.Manual).Select(b => b.Id).ToHashSet();
            var result = new List<ImmutableArray<int>>(); var current = new List<int>(); var size = 0d;
            IEnumerable<int> Prefix(int firstBody) => titles.Where(t => t < firstBody);
            double PrefixSize(int firstBody) => Prefix(firstBody).Sum(t => row ? DrawingGeometry.RowSize(sheet, t) : DrawingGeometry.ColumnSize(sheet, t));
            void Finish() => result.Add(Prefix(current[0]).Concat(current).ToImmutableArray());
            foreach (var i in Visible(first, last, row))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var dimension = row ? DrawingGeometry.RowSize(sheet, i) : DrawingGeometry.ColumnSize(sheet, i);
                var extraTitles = PrefixSize(current.Count == 0 ? i : current[0]);
                if (current.Count > 0 && (size + dimension + extraTitles > available + .001 || cuts.Any(k => k > current[^1] && k <= i)))
                {
                    Finish(); current.Clear(); size = 0;
                }
                // A single oversize row/column is explicitly rejected instead of silently clipping data.
                if (dimension + PrefixSize(current.Count == 0 ? i : current[0]) > available + .001)
                    throw new ArgumentException("单行/列连同重复标题大于一页，请缩小缩放比例。");
                current.Add(i); size += dimension;
                if (result.Count > MaxPages) throw new ArgumentException("打印页数超过 2000 页。");
            }
            if (current.Count > 0) Finish();
            return result;
        }
        List<PrintPage> Paginate(double scale, bool fitting)
        {
            var result = new List<PrintPage>(); long cells = 0;
            foreach (var a in normalized)
            {
                if (printing.RowBreaks.Any(b => b.Manual && b.Id > a.Top && b.Id <= a.Bottom && (b.Minimum > a.Left || b.Maximum < a.Right)) ||
                    printing.ColumnBreaks.Any(b => b.Manual && b.Id > a.Left && b.Id <= a.Right && (b.Minimum > a.Top || b.Maximum < a.Bottom)))
                    throw new NotSupportedException("局部跨度的手动分页尚不能预览；请移除该分页或调整打印区域。");
                var rows = Split(a.Top, a.Bottom, true, usableHeight / scale - (printing.Options?.Headings == true ? 24 : 0), titleRows, fitting ? [] : printing.RowBreaks);
                var cols = Split(a.Left, a.Right, false, usableWidth / scale - (printing.Options?.Headings == true ? 40 : 0), titleCols, fitting ? [] : printing.ColumnBreaks);
                if (fitting && ((setup.FitToWidth ?? 1) > 0 && cols.Count > (setup.FitToWidth ?? 1) || (setup.FitToHeight ?? 1) > 0 && rows.Count > (setup.FitToHeight ?? 1)))
                    throw new ArgumentException("缩放尚未适页。");
                var pairs = setup.PageOrder == "overThenDown" ? rows.SelectMany(r => cols.Select(c => new PrintPage(r, c))) : cols.SelectMany(c => rows.Select(r => new PrintPage(r, c)));
                foreach (var page in pairs)
                {
                    cells += (long)page.Rows.Length * page.Columns.Length;
                    if (result.Count >= MaxPages || cells > MaxPrintedCells) throw new ArgumentException("打印任务超过 2000 页或 2,000,000 格输出预算。");
                    result.Add(page);
                }
            }
            return result;
        }
        var scale = (setup.Scale ?? 100) / 100d;
        if (printing.Properties?.FitToPage == true)
        {
            // Excel fit-to-page shrinks when needed; it does not enlarge a report above 100%.
            try { _ = Paginate(1, true); scale = 1; }
            catch (ArgumentException)
            {
                var low = .1; var high = 1d; _ = Paginate(low, true);
                for (var i = 0; i < 24; i++) { var mid = (low + high) / 2; try { _ = Paginate(mid, true); low = mid; } catch (ArgumentException) { high = mid; } }
                scale = Math.Floor(low * 100) / 100;
            }
        }
        var pages = Paginate(scale, printing.Properties?.FitToPage == true).ToImmutableArray();
        if (pages.IsEmpty) throw new InvalidOperationException("打印区域中没有可见单元格。");
        return new(width, height, margins, scale, pages);
    }

    private static double Measure(string measure)
    {
        var value = double.Parse(measure[..^2], CultureInfo.InvariantCulture);
        return value * (measure[^2..] switch { "mm" => 96 / 25.4, "cm" => 96 / 2.54, "in" => 96, "pt" => 96 / 72d, "pc" or "pi" => 16, _ => throw new NotSupportedException("纸张单位尚未支持。") });
    }

    public static (string Left, string Center, string Right) Header(string? source, string sheet, string file, int page, int pages, DateTime timestamp)
    {
        var parts = new[] { new StringBuilder(), new StringBuilder(), new StringBuilder() }; var section = 1;
        for (var i = 0; i < (source?.Length ?? 0); i++)
        {
            if (source![i] != '&' || i == source.Length - 1) { parts[section].Append(source[i]); continue; }
            var code = source[++i];
            switch (code)
            {
                case 'L': section = 0; break; case 'C': section = 1; break; case 'R': section = 2; break;
                case '&': parts[section].Append('&'); break;
                case 'P': parts[section].Append(page); break; case 'N': parts[section].Append(pages); break;
                case 'A': parts[section].Append(sheet); break; case 'F': parts[section].Append(file); break;
                case 'D': parts[section].Append(timestamp.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)); break;
                case 'T': parts[section].Append(timestamp.ToString("HH:mm", CultureInfo.InvariantCulture)); break;
                default: throw new NotSupportedException($"页眉/页脚指令 &{code} 尚未支持预览，请改用普通文本、&L/&C/&R、&P/&N、&A/&F、&D/&T。");
            }
        }
        return (parts[0].ToString(), parts[1].ToString(), parts[2].ToString());
    }
}

public sealed class SetPrintingCommand(SheetId sheetId, WorksheetPrinting printing) : DocumentCommand
{
    internal override CommandDraft Prepare(WorkbookSnapshot snapshot)
    {
        var sheet = snapshot.GetSheet(sheetId);
        if (sheet.Printing.ContentEquals(printing)) return new(snapshot, [], false);
        var updated = sheet.WithPrinting(printing);
        return new(snapshot.WithSheets(snapshot.Sheets.Select(s => s.Id == sheetId ? updated : s)), [], true, CalculationUnchanged: true);
    }
}
