using System.Xml;
using QuickExcel.Domain;

namespace QuickExcel.Core;

public sealed record TextImportResult(WorkbookSnapshot Workbook, IReadOnlyList<string> Notices, int Rows, int Columns);

public static class TextImport
{
    public static TextImportResult Convert(IReadOnlyList<string[]> rows, IReadOnlyList<CleaningColumnType> types,
        NumericTextOptions numeric, ExcelDateSystem dates = ExcelDateSystem.Windows1900, bool header = false, CancellationToken cancellation = default)
    {
        numeric.Validate();
        var columns = rows.Count == 0 ? 0 : rows.Max(r => r.Length);
        if (rows.Count > WorkbookLimits.MaxRows || columns > WorkbookLimits.MaxColumns || (long)rows.Count * columns > WorkbookLimits.MaxCells)
            throw new ArgumentException(MessageText.Get("文本交换范围超过 1,000,000 格，请缩小范围。"));
        if (types.Count != columns || types.Any(t => !Enum.IsDefined(t) || t == CleaningColumnType.General)) throw new ArgumentException(nameof(types));
        var styles = new List<StyleDefinition> { new() }; var notices = new HashSet<string>();
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        for (var row = 0; row < rows.Count; row++)
        {
            cancellation.ThrowIfCancellationRequested();
            for (var column = 0; column < rows[row].Length; column++)
            {
                if (types[column] == CleaningColumnType.Skip) continue;
                var text = rows[row][column]; WorkbookLimits.ValidateText(text); XmlConvert.VerifyXmlChars(text);
                try
                {
                    var converted = CleaningText.Convert(text, header && row == 0 ? CleaningColumnType.Text : types[column], numeric, dates, notices);
                    // Explicit blank records retain the rectangular extent for a later text export.
                    var style = new StyleDefinition { NumberFormatId = converted.FormatId, NumberFormatCode = converted.FormatCode };
                    var styleId = styles.IndexOf(style);
                    if (styleId < 0) { styleId = styles.Count; styles.Add(style); }
                    cells.Add(new(row, column), new(CellContent.FromLiteral(converted.Value, styleId)));
                }
                catch (Exception ex) when (ex is ArgumentException or OverflowException)
                { throw new ArgumentException(MessageText.Format($"第 {row + 1} 行，第 {column + 1} 列：{ex.Message}"), ex); }
            }
        }
        return new(new(Guid.NewGuid(), 0, dates, [new(SheetId.New(), "Sheet1", cells)], styles: styles), notices.ToArray(), rows.Count, columns);
    }
}
