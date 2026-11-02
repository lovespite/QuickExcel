using FastWpfGrid;
using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class PresentationContractTests
{
    private static WorkbookServices Services => new(new TypedWorkbookStorage(), new WorkbookCalculator(new NoScripts()));

    [Fact]
    public void SharedControllerCommitsPasteAndClearAsSingleHistoryEntries()
    {
        using var document = WorkbookDocument.CreateBlank(Services);
        var editor = new EditorController(document);
        editor.SetCells(0, GridClipboard.Prepare("1\t=2+3", new(0, 0, 0, 0), allowFormulas: false));
        Assert.Equal(CellValue.FromText("=2+3"), document.Snapshot.Sheets[0].GetCell(0, 1)!.Value);
        editor.Undo();
        Assert.Empty(document.Snapshot.Sheets[0].Cells);
        editor.Redo();
        editor.ClearCells(0, [(0, 0), (0, 1)]);
        Assert.Empty(document.Snapshot.Sheets[0].Cells);
        editor.Undo();
        Assert.Equal(2, document.Snapshot.Sheets[0].Cells.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ValidationConfirmationCannotApplyToANewerDocumentRevision(bool confirm)
    {
        var rule = new WorksheetValidation([new(0, 0, 0, 0)], ValidationKind.Whole, "=1", "=5", alert: ValidationAlert.Warning);
        var sheet = Sheet("Data").WithData(new() { Validations = [rule] });
        using var document = new WorkbookDocument(Book(sheet), Services);
        var confirmations = 0;
        var editor = new EditorController(document, _ =>
        {
            confirmations++;
            // A dialog may dispatch another edit while waiting for a response.
            document.Session.Execute(new SetCellsCommand([new(new(sheet.Id, 1, 0), CellInput.Parse("42"))]));
            return confirm;
        });
        if (confirm) Assert.Throws<InvalidOperationException>(() => editor.SetCell(0, 0, 0, "9"));
        else Assert.Throws<DataValidationException>(() => editor.SetCell(0, 0, 0, "9"));
        Assert.Equal(1, confirmations);
        Assert.Null(document.Snapshot.Sheets[0].GetCell(0, 0));
        Assert.Equal(CellValue.FromNumber(42), document.Snapshot.Sheets[0].GetCell(1, 0)!.Value);
        editor.Undo();
        Assert.Empty(document.Snapshot.Sheets[0].Cells);
    }

    [Fact]
    public void LocalFileSelectionRejectsProviderUrisAndRelativePaths()
    {
        Assert.Throws<NotSupportedException>(() => LocalFileSelection.FromUri(new("https://example.invalid/book.xlsx")));
        Assert.Throws<NotSupportedException>(() => LocalFileSelection.FromUri(new("content://provider/book.xlsx")));
        Assert.Throws<ArgumentException>(() => new LocalFileSelection("book.xlsx"));
        var path = Path.Combine(Path.GetTempPath(), "中文 book.xlsx");
        Assert.Equal(Path.GetFullPath(path), LocalFileSelection.FromUri(new Uri(path)).Path);
    }

    [Fact]
    public void SharedAssembliesHaveNoUiInfrastructureOrNativeAssemblyDependencies()
    {
        var allowed = new[] { "QuickExcel.Core", "QuickExcel.Grid.Shared" };
        foreach (var assembly in new[] { typeof(EditorController).Assembly, typeof(SeriesSizes).Assembly })
        {
            Assert.NotEqual(typeof(PresentationContractTests).Assembly, assembly);
            Assert.All(assembly.GetReferencedAssemblies(), reference =>
                Assert.True(reference.Name!.StartsWith("System.", StringComparison.Ordinal) ||
                    reference.Name == "System" || allowed.Contains(reference.Name), reference.FullName));
        }
    }
}
