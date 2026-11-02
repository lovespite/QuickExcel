using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Rendering;

/// <summary>A frozen report revision. Projection is supplied by the composition root, never by a calculator in a draw callback.</summary>
public sealed class PrintJob
{
    public WorkbookSnapshot Book { get; }
    public WorksheetSnapshot Sheet { get; }
    public PrintLayout Layout { get; }
    public string FileName { get; }
    public DateTime Timestamp { get; }
    public IReadOnlyList<string> Warnings { get; }
    private readonly Func<WorkbookSnapshot, WorksheetSnapshot, int, int, StyleDefinition> _style;
    public PrintJob(WorkbookSnapshot book, WorksheetSnapshot outputSheet, string file, DateTime timestamp, IEnumerable<string> warnings,
        Func<WorkbookSnapshot, WorksheetSnapshot, int, int, StyleDefinition>? style = null, CancellationToken cancellationToken = default)
    {
        Book = book; Sheet = outputSheet; FileName = Path.GetFileName(file); Timestamp = timestamp;
        Warnings = Array.AsReadOnly(warnings.Distinct().ToArray());
        _style = style ?? ((b, s, r, c) => b.Styles[s.GetStyleId(r, c)]);
        Layout = PrintingLayout.Create(outputSheet, cancellationToken);
    }
    public CellRenderInput Cell(int row, int column, RenderRect rect)
    {
        var cell = Sheet.GetCell(row, column); var style = _style(Book, Sheet, row, column);
        var display = CellDisplayFormatter.FormatDisplay(cell?.Value ?? CellValue.Blank, style, Book.DateSystem);
        if (display.Color is not null) style = style with { FontColor = display.Color };
        if (style.Horizontal == 0 && cell?.Value.Kind == CellValueKind.Number) style = style with { Horizontal = 3 };
        var text = cell?.Value.Kind == CellValueKind.Error ? Sheet.Printing.Setup?.Errors switch
        { "blank" => "", "dash" => "--", "NA" => "#N/A", _ => display.Text } : display.Text;
        return new(rect, new(text, style, RichText: cell?.Content.RichText, FillOffset: display.FillOffset, FillCharacter: display.FillCharacter), Sheet.Printing.Options?.GridLines == true);
    }

    public static WorksheetPrinting PreparePrinting(WorksheetPrinting printing, string file, DateTime timestamp, ICollection<string> warnings)
    {
        if (printing.Setup?.CellComments is "asDisplayed" or "atEnd")
        { warnings.Add("批注打印暂未支持，本次不输出批注。"); printing = printing with { Setup = printing.Setup with { CellComments = "none" } }; }
        string? Header(string? text, string label)
        {
            try { _ = PrintingLayout.Header(text, "", Path.GetFileName(file), 1, 1, timestamp); return text; }
            catch (NotSupportedException ex) { warnings.Add($"{label}包含未支持的内容，本次不输出该页眉或页脚：{ex.Message}"); return null; }
        }
        if (printing.HeaderFooter is { } h)
            printing = printing with { HeaderFooter = h with {
                OddHeader = Header(h.OddHeader, "默认/奇数页页眉"), OddFooter = Header(h.OddFooter, "默认/奇数页页脚"),
                EvenHeader = Header(h.EvenHeader, "偶数页页眉"), EvenFooter = Header(h.EvenFooter, "偶数页页脚"),
                FirstHeader = Header(h.FirstHeader, "首页页眉"), FirstFooter = Header(h.FirstFooter, "首页页脚") } };
        return printing;
    }
}
