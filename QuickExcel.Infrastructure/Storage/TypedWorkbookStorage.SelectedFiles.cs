using QuickExcel.Domain;

namespace QuickExcel.Storage;

public sealed partial class TypedWorkbookStorage
{
    // The file broker owns the stream and publication. The path supplies provenance only.
    internal static WorkbookSnapshot ReadSelectedWorkbook(Stream input, string path, FileReadBudget budget) =>
        OpenBytes(PackageIO.ReadFile(input, budget), path, IsXlsq(path), budget);

    internal static byte[] WriteSelectedWorkbook(WorkbookSnapshot snapshot, string path, CancellationToken cancellation, string? copyType = null)
    {
        var packageFormat = copyType switch
        {
            null => IsXlsq(path), "xlsx" => false, "xlsq" => true,
            _ => throw new ArgumentException("Expected xlsx or xlsq.", nameof(copyType))
        };
        snapshot = PrepareReadOnlyFormulaSave(snapshot, false);
        GuardOriginal(path, snapshot, false);
        if (packageFormat)
        {
            var package = SerializePackage(snapshot, false, cancellation);
            VerifyPackage(path, snapshot, package, cancellation);
            return package;
        }
        var bytes = BuildXlsx(snapshot, false, cancellation, out var preserve, workbookPartOnly: copyType == "xlsx");
        VerifyXlsx(path, snapshot, bytes, preserve, cancellation);
        return bytes;
    }

    internal static bool IsXlsq(string path) => Path.GetExtension(path).ToLowerInvariant() switch
    {
        ".xlsq" => true,
        ".xlsx" => false,
        _ => throw new NotSupportedException("Expected an XLSX or XLSQ file.")
    };
}
