using System.Collections.Immutable;

namespace QuickExcel.Domain;

public sealed record PageMargins(double Left, double Right, double Top, double Bottom, double Header, double Footer);
public sealed record PageSetup
{
    public uint? PaperSize { get; init; }
    public string? PaperHeight { get; init; }
    public string? PaperWidth { get; init; }
    public uint? Scale { get; init; }
    public uint? FirstPageNumber { get; init; }
    public uint? FitToWidth { get; init; }
    public uint? FitToHeight { get; init; }
    public string? PageOrder { get; init; }
    public string? Orientation { get; init; }
    public bool? UsePrinterDefaults { get; init; }
    public bool? BlackAndWhite { get; init; }
    public bool? Draft { get; init; }
    public string? CellComments { get; init; }
    public bool? UseFirstPageNumber { get; init; }
    public string? Errors { get; init; }
    public uint? HorizontalDpi { get; init; }
    public uint? VerticalDpi { get; init; }
    public uint? Copies { get; init; }
}
public sealed record PageSetupProperties(bool? AutoPageBreaks = null, bool? FitToPage = null);
public sealed record PrintOptions(bool? HorizontalCentered = null, bool? VerticalCentered = null,
    bool? Headings = null, bool? GridLines = null, bool? GridLinesSet = null);
public sealed record PageHeaderFooter
{
    public bool? DifferentOddEven { get; init; }
    public bool? DifferentFirst { get; init; }
    public bool? ScaleWithDoc { get; init; }
    public bool? AlignWithMargins { get; init; }
    public string? OddHeader { get; init; }
    public string? OddFooter { get; init; }
    public string? EvenHeader { get; init; }
    public string? EvenFooter { get; init; }
    public string? FirstHeader { get; init; }
    public string? FirstFooter { get; init; }
}
public enum PrintRangeKind { Cells, Rows, Columns }
/// <summary>Sheet-local, absolute printing reference. Metadata may span Excel's full sheet without allocating cells.</summary>
public sealed record PrintRange(PrintRangeKind Kind, int Top, int Left, int Bottom, int Right)
{
    public const int MaxRows = 1048576, MaxColumns = 16384;
    internal void Validate()
    {
        if (!Enum.IsDefined(Kind) || Top < 0 || Left < 0 || Bottom < Top || Right < Left || Bottom >= MaxRows || Right >= MaxColumns ||
            Kind == PrintRangeKind.Rows && (Left != 0 || Right != MaxColumns - 1) ||
            Kind == PrintRangeKind.Columns && (Top != 0 || Bottom != MaxRows - 1)) throw new ArgumentException("打印引用范围无效。");
    }
}
/// <summary>Boundary before row/column Id (zero-based), with the perpendicular span inclusive.</summary>
public sealed record PrintPageBreak(int Id, int Minimum, int Maximum, bool Manual = false, bool Pivot = false);

public sealed record WorksheetPrinting
{
    public PageMargins? Margins { get; init; }
    public PageSetup? Setup { get; init; }
    public PageSetupProperties? Properties { get; init; }
    public PrintOptions? Options { get; init; }
    public PageHeaderFooter? HeaderFooter { get; init; }
    public ImmutableArray<PrintRange> Areas { get; init; } = [];
    public ImmutableArray<PrintRange> Titles { get; init; } = [];
    public ImmutableArray<PrintPageBreak> RowBreaks { get; init; } = [];
    public ImmutableArray<PrintPageBreak> ColumnBreaks { get; init; } = [];
    // Validated relationship to an opaque printer configuration; never executed or used to contact a printer.
    public ImmutableArray<byte> PrinterSettings { get; init; } = [];
    public bool HasReferences => !Areas.IsEmpty || !Titles.IsEmpty;
    public bool IsEmpty => Margins is null && Setup is null && Properties is null && Options is null && HeaderFooter is null &&
        !HasReferences && RowBreaks.IsEmpty && ColumnBreaks.IsEmpty && PrinterSettings.IsEmpty;
    public bool ContentEquals(WorksheetPrinting other) => Margins == other.Margins && Setup == other.Setup && Properties == other.Properties &&
        Options == other.Options && HeaderFooter == other.HeaderFooter && Areas.SequenceEqual(other.Areas) && Titles.SequenceEqual(other.Titles) &&
        RowBreaks.SequenceEqual(other.RowBreaks) && ColumnBreaks.SequenceEqual(other.ColumnBreaks) && PrinterSettings.SequenceEqual(other.PrinterSettings);
    internal void Validate()
    {
        if (Areas.IsDefault || Titles.IsDefault || RowBreaks.IsDefault || ColumnBreaks.IsDefault || PrinterSettings.IsDefault ||
            Areas.Length > 1024 || Titles.Length > 2 || RowBreaks.Length > 1026 || ColumnBreaks.Length > 1026 || PrinterSettings.Length > 1024 * 1024)
            throw new ArgumentException("打印元数据超过预算。");
        if (Margins is { } m && new[] { m.Left, m.Right, m.Top, m.Bottom, m.Header, m.Footer }.Any(v => !double.IsFinite(v) || v < 0))
            throw new ArgumentException("页边距必须为非负有限英寸数。");
        if (Setup is { } s)
        {
            if (s.Scale is < 10 or > 400 || s.Orientation is not (null or "default" or "portrait" or "landscape") ||
                s.PageOrder is not (null or "downThenOver" or "overThenDown") || s.CellComments is not (null or "none" or "asDisplayed" or "atEnd") ||
                s.Errors is not (null or "displayed" or "blank" or "dash" or "NA")) throw new ArgumentException("页面设置参数无效。");
            foreach (var measure in new[] { s.PaperHeight, s.PaperWidth })
                if (measure is not null && !System.Text.RegularExpressions.Regex.IsMatch(measure, @"^[0-9]+(?:\.[0-9]+)?(?:mm|cm|in|pt|pc|pi)$"))
                    throw new ArgumentException("自定义纸张尺寸无效。");
        }
        if (!PrinterSettings.IsEmpty && Setup is null) throw new ArgumentException("打印机配置需要页面设置。");
        foreach (var r in Areas.Concat(Titles)) r.Validate();
        if (Titles.Any(r => r.Kind == PrintRangeKind.Cells) || Titles.Select(r => r.Kind).Distinct().Count() != Titles.Length)
            throw new ArgumentException("重复标题必须是单组行和/或列。");
        void Breaks(ImmutableArray<PrintPageBreak> breaks, int limit, int perpendicular)
        {
            if (breaks.Select(b => b.Id).Distinct().Count() != breaks.Length || breaks.Any(b => b.Id < 1 || b.Id > limit || b.Minimum < 0 || b.Maximum < b.Minimum || b.Maximum >= perpendicular))
                throw new ArgumentException("分页位置无效或重复。");
        }
        Breaks(RowBreaks, PrintRange.MaxRows, PrintRange.MaxColumns); Breaks(ColumnBreaks, PrintRange.MaxColumns, PrintRange.MaxRows);
        if (HeaderFooter is { } h)
            foreach (var text in new[] { h.OddHeader, h.OddFooter, h.EvenHeader, h.EvenFooter, h.FirstHeader, h.FirstFooter })
                if (text is not null) WorkbookLimits.ValidateText(text);
    }
}
