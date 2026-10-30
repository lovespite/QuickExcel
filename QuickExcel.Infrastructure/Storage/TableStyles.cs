using QuickExcel.Domain;

namespace QuickExcel.Storage;

public static class TableStyles
{
    internal static byte[] DefaultTheme
    {
        get { using var stream = typeof(TableStyles).Assembly.GetManifestResourceStream("QuickExcel.Infrastructure.Storage.DefaultTableTheme.xml")!; using var output = new MemoryStream(); stream.CopyTo(output); return output.ToArray(); }
    }
    public static TableStyleInfo Resolve(WorkbookSnapshot snapshot, string name, bool rows = true, bool columns = false, bool first = false, bool last = false)
    {
        var bytes = snapshot.FileState?.OriginalWorkbook.ToArray() ?? TypedWorkbookXml.Write(snapshot, true, CancellationToken.None);
        var package = new XlsxPackage(bytes, new(CancellationToken.None)); var workbook = package.RelatedByType("", "officeDocument", true)!;
        return new(name, first, last, rows, columns) { Elements = XlsxTableStyles.Resolve(package, workbook, name) };
    }
}
