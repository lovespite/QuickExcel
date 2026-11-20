using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class DataTransferStorageTests
{
    [Fact]
    public void ExtractedApi4ValuesExportAsScriptFreeXlsxAndNormalizeHostErrors()
    {
        var directory = Path.Combine(Path.GetTempPath(), "quickexcel-transfer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
                [new(SheetId.New(), "Data", [KeyValuePair.Create(new CellPosition(0, 0),
                    new CellSnapshot(CellContent.FromLiteral(CellValue.FromError(CellErrorCode.Script))))])]);
            var values = WorkbookDataTransfer.ExtractValues(book, [new(book.ActiveSheetId)]).Workbook;
            var path = Path.Combine(directory, "values.xlsx");
            var storage = new TypedWorkbookStorage();
            storage.ExportDataValuesChecked(path, values, FileVersions.Read(path));
            var reopened = storage.Open(path);
            Assert.Equal(CellErrorCode.Value, reopened.Sheets[0].GetCell(0, 0)!.Value.ErrorCode);
            Assert.Equal(4, values.ScriptModules.ScriptApiVersion);
            Assert.False(TypedWorkbookStorage.RequiresScriptContainer(reopened));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void CancelledCheckedXlsxExportPreservesExistingFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "quickexcel-transfer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "existing.xlsx");
            var storage = new TypedWorkbookStorage();
            var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
                [new(SheetId.New(), "Data")]);
            storage.Save(path, book);
            var before = File.ReadAllBytes(path);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => storage.SaveChecked(path, book, false,
                FileVersions.Read(path), cancellation.Token));
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void CheckedXlsxExportRejectsChangedDestination()
    {
        var directory = Path.Combine(Path.GetTempPath(), "quickexcel-transfer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "existing.xlsx");
            var storage = new TypedWorkbookStorage();
            var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
                [new(SheetId.New(), "Data")]);
            storage.Save(path, book);
            var expected = FileVersions.Read(path);
            File.WriteAllText(path, "newer external file");
            Assert.ThrowsAny<IOException>(() => storage.SaveChecked(path, book, false, expected, default));
            Assert.Equal("newer external file", File.ReadAllText(path));
        }
        finally { Directory.Delete(directory, true); }
    }
}
