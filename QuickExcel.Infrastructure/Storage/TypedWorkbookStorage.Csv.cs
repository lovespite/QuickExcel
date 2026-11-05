using System.Xml;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Storage;

public sealed partial class TypedWorkbookStorage
{
    private static WorkbookSnapshot ImportCsv(byte[] bytes, CancellationToken cancellation)
    {
        var offset = bytes.AsSpan().StartsWith(System.Text.Encoding.UTF8.Preamble) ? 3 : 0;
        var text = PackageIO.Utf8.GetString(bytes, offset, bytes.Length - offset);
        var cells = new Dictionary<CellPosition, CellSnapshot>();
        var row = 0;
        foreach (var fields in DelimitedText.ReadRows(text, ',', WorkbookLimits.MaxCells))
        {
            cancellation.ThrowIfCancellationRequested();
            if (row >= WorkbookLimits.MaxRows || fields.Length > WorkbookLimits.MaxColumns)
                throw new InvalidDataException("CSV 超出工作表行列上限。");
            for (var col = 0; col < fields.Length; col++)
            {
                var field = fields[col];
                if (field.Length == 0) continue;
                XmlConvert.VerifyXmlChars(field);
                cells.Add(new(row, col), new(CellContent.FromLiteral(CellValue.FromText(field))));
            }
            row++;
        }
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(SheetId.New(), "Sheet1", cells)]);
    }
}
