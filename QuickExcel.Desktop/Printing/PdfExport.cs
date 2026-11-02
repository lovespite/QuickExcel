namespace QuickExcel.Desktop;

/// <summary>Direct vector/text PDF; no XPS spool, printer queue or driver fallback.</summary>
internal static class PdfExport
{
    internal static Task WriteAsync(WorkbookPaginator paginator, string destination, CancellationToken cancellationToken = default) =>
        paginator.WritePdfAsync(destination, cancellationToken);
}
