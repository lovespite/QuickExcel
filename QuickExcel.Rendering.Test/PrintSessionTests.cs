using QuickExcel.Domain;
using QuickExcel.Rendering.Skia;
using UglyToad.PdfPig;

namespace QuickExcel.Rendering.Test;

public sealed class PrintSessionTests
{
    private static WorkbookSnapshot Book()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Sheet", new Dictionary<CellPosition, CellSnapshot> { [new(0, 0)] = new(CellContent.FromLiteral(CellValue.FromText("Frozen report"))) });
        return new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
    }
    private static string DirectoryPath() { var path = Path.Combine(AppContext.BaseDirectory, "test-data", "print-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(path); return path; }
    [Fact]
    public async Task CompletedPdfAndExportUseTheSameImmutableJobAndLeasesDelayDeletion()
    {
        var directory = DirectoryPath(); var book = Book();
        var session = await PrintSession.CreateAsync(book, book.ActiveSheetId, "original.xlsq", directory);
        var preview = session.PdfPath; Assert.Same(book, session.Job.Book);
        using (var parsed = PdfDocument.Open(preview)) Assert.Contains("Frozen", parsed.GetPage(1).Text);
        var destination = Path.Combine(directory, "export.pdf"); await session.ExportAsync(destination);
        using (var bitmap = await session.RenderPageAsync(0, 96)) Assert.True(bitmap.Width > 100 && bitmap.Height > 100);
        using (var original = PdfDocument.Open(preview)) using (var exported = PdfDocument.Open(destination)) Assert.Equal(original.GetPage(1).Text, exported.GetPage(1).Text);
        var lease = session.RetainPdf(); var closing = session.DisposeAsync().AsTask(); Assert.False(closing.IsCompleted); Assert.True(File.Exists(preview));
        Assert.Throws<ObjectDisposedException>(() => session.RetainPdf()); await lease.DisposeAsync(); await closing;
        Assert.False(File.Exists(preview)); Assert.True(File.Exists(destination)); await session.DisposeAsync();
        File.Delete(destination); Directory.Delete(directory);
    }
    [Fact]
    public async Task CancelledExportPreservesExistingTargetAndLeavesNoTemporaryFiles()
    {
        var directory = DirectoryPath(); var book = Book();
        await using (var session = await PrintSession.CreateAsync(book, book.ActiveSheetId, "book.xlsq", directory))
        {
            var destination = Path.Combine(directory, "existing.pdf"); await File.WriteAllTextAsync(destination, "untouched");
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => session.ExportAsync(destination, cancelled.Token));
            Assert.Equal("untouched", await File.ReadAllTextAsync(destination)); Assert.Empty(Directory.GetFiles(directory, "*.tmp", SearchOption.AllDirectories));
            Assert.Throws<ArgumentOutOfRangeException>(() => { _ = session.RenderPageAsync(0, 301); }); File.Delete(destination);
        }
        Assert.Empty(Directory.GetFileSystemEntries(directory)); Directory.Delete(directory);
    }
    [Fact]
    public async Task CancellationDuringRenderingRemovesTheIncompletePreview()
    {
        var directory = DirectoryPath(); var book = Book(); using var cancel = new CancellationTokenSource(); var visited = 0;
        StyleDefinition Resolve(WorkbookSnapshot snapshot, WorksheetSnapshot sheet, int row, int column)
        { visited++; cancel.Cancel(); return new(); }
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => PrintSession.CreateAsync(book, book.ActiveSheetId, "book", directory, Resolve, cancellationToken: cancel.Token));
        Assert.True(visited > 0); Assert.Empty(Directory.GetFileSystemEntries(directory)); Directory.Delete(directory);
    }
    [Fact]
    public async Task CancelledCreationCleansOnlyItsOwnDirectory()
    {
        var directory = DirectoryPath(); var unrelated = Path.Combine(directory, "keep.txt"); await File.WriteAllTextAsync(unrelated, "keep");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel(); var book = Book();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => PrintSession.CreateAsync(book, book.ActiveSheetId, "book", directory, cancellationToken: cancelled.Token));
        Assert.Equal([unrelated], Directory.GetFiles(directory)); Assert.Empty(Directory.GetDirectories(directory)); File.Delete(unrelated); Directory.Delete(directory);
    }
}
