using System.Globalization;
using System.Text;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

public sealed record TextExportOptions(char Separator = ',', bool Bom = true, bool DisplayText = false);

/// <summary>Bounded, cancellable text snapshots. Export never changes the document or its saved identity.</summary>
public static class WorkbookTextExchange
{
    public const int MaximumFields = 1_000_000;
    public static CellRectangle UsedRange(WorksheetSnapshot sheet) => sheet.Cells.Count == 0 ? new(0, 0, 0, 0) :
        new(0, 0, sheet.Cells.Keys.Max(p => p.Row), sheet.Cells.Keys.Max(p => p.Column));

    public static byte[] Export(WorkbookSnapshot book, SheetId sheetId, CellRectangle range,
        TextExportOptions options, CancellationToken cancellation = default)
    {
        if (options.Separator is not (',' or '\t')) throw new ArgumentException(nameof(options));
        if ((long)range.Rows * range.Columns > MaximumFields)
            throw new InvalidOperationException(MessageText.Get("文本交换范围超过 1,000,000 格，请缩小范围。"));
        var sheet = book.GetSheet(sheetId);
        var text = new StringBuilder();
        for (var row = range.Top; row <= range.Bottom; row++)
        {
            cancellation.ThrowIfCancellationRequested();
            for (var column = range.Left; column <= range.Right; column++)
            {
                if (column != range.Left) text.Append(options.Separator);
                var cell = sheet.GetCell(row, column);
                if ((cell?.Content.Kind == CellContentKind.Formula || cell?.SpillOwner is not null) &&
                    (cell.Calculation?.Value is null || cell.Calculation.InputRevision != book.Revision ||
                     cell.Calculation.State is not (CalculationState.Fresh or CalculationState.Error)))
                    throw new InvalidOperationException(MessageText.Get("导出区域存在未计算或过期的公式结果，请先重新计算。"));
                var value = cell?.Value ?? CellValue.Blank;
                var field = options.DisplayText ? CellDisplayFormatter.Format(value, book.Styles[sheet.GetStyleId(row, column)], book.DateSystem) :
                    value.Kind == CellValueKind.Number ? value.NumberValue.ToString("R", CultureInfo.InvariantCulture) : CellInput.DisplayText(value);
                if (field.IndexOfAny([options.Separator, '"', '\r', '\n']) >= 0)
                    text.Append('"').Append(field.Replace("\"", "\"\"")).Append('"');
                else text.Append(field);
                if (text.Length > 32 * 1024 * 1024)
                    throw new InvalidOperationException(MessageText.Get("文本交换超过 32 MiB 字符预算，请缩小范围。"));
            }
            text.Append("\r\n");
        }
        var encoding = new UTF8Encoding(options.Bom, true);
        return [.. encoding.GetPreamble(), .. encoding.GetBytes(text.ToString())];
    }

    public static void Save(string path, byte[] bytes, FileVersion expected, CancellationToken cancellation = default) =>
        PackageIO.AtomicWrite(path, bytes, _ => { }, cancellation, expected);
}
