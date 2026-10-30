using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public class SheetManagementTests
{
    private static WorkbookSnapshot NamedBook()
    {
        var source = Sheet("Source", (0, 0, N(10)), (1, 0, N(20)), (2, 0, N(30)), (0, 1, F("=A1+Source!$A$2")),
            (1, 1, F("=GlobalSource")), (2, 1, F("=Rate")), (3, 1, F("=Local")), (4, 1, F("=Source!Local")), (5, 1, F("=GlobalOther")),
            (0, 2, F("=Chain")), (1, 2, F("=LocalChain")), (2, 2, F("=Other!Local")), (3, 2, F("=\"Source!A1\"")), (4, 2, F("=GlobalRateAlias")));
        var other = Sheet("Other", (0, 0, N(100)), (0, 1, F("=GlobalSource+Rate")), (1, 1, F("=Source!Rate")));
        return Book(source, other).WithDefinedNames([new("GlobalSource", "=Source!$A$1"), new("GlobalOther", "=Other!$A$1"), new("Rate", "=Source!$A$1"),
            new("Rate", "=Source!$A$2", source.Id), new("Local", "=Source!$A$3", source.Id), new("Chain", "=GlobalSource*2"),
            new("LocalChain", "=Source!Local+Source!Rate", source.Id), new("Local", "=Other!$A$1", other.Id), new("GlobalRateAlias", "=Rate"),
            new("UnusedSource", "=Source!$A$3"), new("UnusedOther", "=Other!$A$1")]);
    }
    private static void Calculate(DocumentSession session)
    {
        var result = new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation());
        Assert.True(session.TryCommitCalculation(result));
    }
    [Fact]
    public void CopyLocalizesGlobalReferenceClosurePreservesScopesAndIsAtomic()
    {
        var original = NamedBook(); using var session = new DocumentSession(original);
        var command = new CopySheetCommand(original.Sheets[0].Id, "Copy", 1); session.Execute(command); Calculate(session);
        var copy = session.Snapshot.GetSheet(command.Id);
        Assert.Equal(new[] { "Source", "Copy", "Other" }, session.Snapshot.Sheets.Select(s => s.Name));
        Assert.Equal(command.Id, session.Snapshot.ActiveSheetId);
        Assert.Equal(new[] { "Chain", "GlobalRateAlias", "GlobalSource", "Local", "LocalChain", "Rate", "UnusedSource" },
            session.Snapshot.DefinedNames.Where(n => n.Scope == command.Id).Select(n => n.Name).Order());
        Assert.Contains(FormulaSyntax.Walk(FormulaSyntax.Parse(copy.GetCell(0, 1)!.Content.FormulaText)).OfType<ReferenceNode>(), r => r.Sheet == "Copy");
        Assert.Equal("=\"Source!A1\"", copy.GetCell(3, 2)!.Content.FormulaText);
        Assert.Equal(10, copy.GetCell(4, 2)!.Value.NumberValue); // Bare Rate in a definition stays global even when a local Rate exists.
        Assert.Equal(20, copy.GetCell(2, 1)!.Value.NumberValue);
        Assert.Equal(100, copy.GetCell(2, 2)!.Value.NumberValue);
        session.Execute(new SetCellsCommand([new(new(command.Id, 0, 0), N(77).Content)])); Calculate(session);
        Assert.Equal(77, session.Snapshot.GetSheet(command.Id).GetCell(1, 1)!.Value.NumberValue);
        Assert.Equal(154, session.Snapshot.GetSheet(command.Id).GetCell(0, 2)!.Value.NumberValue);
        Assert.Equal(10, session.Snapshot.GetSheet(command.Id).GetCell(4, 2)!.Value.NumberValue);
        Assert.Equal(10, session.Snapshot.Sheets[0].GetCell(1, 1)!.Value.NumberValue);
        session.Undo(); session.Undo(); Assert.Equal(2, session.Snapshot.Sheets.Length); Assert.Equal(original.DefinedNames, session.Snapshot.DefinedNames);
        Assert.False(session.CanUndo); Assert.True(session.Redo()); Calculate(session); Assert.Equal(10, session.Snapshot.GetSheet(command.Id).GetCell(1, 1)!.Value.NumberValue);
    }
    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void CopiedStylesImagesMergesDimensionsPrintingAndViewsSurviveReorderAndReopen(string extension)
    {
        var pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aA1sAAAAASUVORK5CYII=");
        var book = NamedBook(); var original = book.Sheets[0];
        var rich = CellContent.FromRichText(new([new("rich", new(Bold: true, Color: "FFFF0000")), new(" text")]));
        var cells = original.Cells.SetItem(new(0, 0), new(N(10).Content.WithStyle(1))).Add(new(1, 3), new(rich));
        var source = new WorksheetSnapshot(original.Id, original.Name, cells, [new(0, new(35))], [new(0, new(17))], ["D1:E1"], [new(pixel, "png", 3, 3, 2, 2)],
            view: new() { Zoom = 85 }, printing: new() { Areas = [new(PrintRangeKind.Cells, 0, 0, 5, 5)] });
        book = new(book.DocumentId, book.Revision, book.DateSystem, [source, book.Sheets[1]], styles: [new(), new() { Bold = true, NumberFormatId = 2 }], definedNames: book.DefinedNames);
        using var session = new DocumentSession(book); var command = new CopySheetCommand(source.Id, "Copy", 1); session.Execute(command);
        session.Execute(new ReorderSheetCommand(command.Id, 0)); Calculate(session);
        var copy = session.Snapshot.Sheets[0]; Assert.Equal(command.Id, copy.Id); Assert.Equal(85, copy.View.Zoom);
        Assert.Equal(source.Images[0].Data, copy.Images[0].Data); Assert.Equal(source.MergedCells, copy.MergedCells); Assert.True(source.Printing.ContentEquals(copy.Printing));
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "." + extension);
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var reopened = storage.Open(path);
            Assert.Equal(new[] { "Copy", "Source", "Other" }, reopened.Sheets.Select(s => s.Name));
            var restored = reopened.Sheets[0]; Assert.Equal(85, restored.View.Zoom); Assert.Equal(35, restored.Rows[0].Height); Assert.Equal(17, restored.Columns[0].Width);
            Assert.True(source.Images[0].Data.SequenceEqual(restored.Images[0].Data)); Assert.True(source.MergedCells.SequenceEqual(restored.MergedCells)); Assert.True(source.Printing.Areas.SequenceEqual(restored.Printing.Areas));
            Assert.True(reopened.Styles[restored.GetStyleId(0, 0)].Bold); Assert.Equal(2, reopened.Styles[restored.GetStyleId(0, 0)].NumberFormatId);
            Assert.Equal(rich.RichText, restored.GetCell(1, 3)!.Content.RichText);
            Assert.Equal(7, reopened.DefinedNames.Count(n => n.Scope == restored.Id));
            using var check = new DocumentSession(reopened); Calculate(check); Assert.Equal(10, check.Snapshot.Sheets[0].GetCell(4, 2)!.Value.NumberValue);
        }
        finally { File.Delete(path); }
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RenameUsesOriginalBindingAndRewritesNameDefinitions(bool local)
    {
        var book = NamedBook(); var source = book.Sheets[0]; using var session = new DocumentSession(book);
        session.Execute(new RenameDefinedNameCommand("Rate", local ? source.Id : null, "RenamedRate")); Calculate(session);
        Assert.Equal(local ? "=RenamedRate" : "=Rate", session.Snapshot.Sheets[0].GetCell(2, 1)!.Content.FormulaText);
        var alias = session.Snapshot.DefinedNames.Single(n => n.Name == "GlobalRateAlias"); Assert.Equal(local ? "=Rate" : "=RenamedRate", alias.Formula);
        Assert.Equal(20, session.Snapshot.Sheets[0].GetCell(2, 1)!.Value.NumberValue); Assert.Equal(10, session.Snapshot.Sheets[0].GetCell(4, 2)!.Value.NumberValue);
        Assert.True(session.Undo()); Assert.Equal(book.DefinedNames, session.Snapshot.DefinedNames); Assert.True(session.Redo()); Calculate(session);
        Assert.Equal(10, session.Snapshot.Sheets[0].GetCell(4, 2)!.Value.NumberValue);
    }
    [Fact]
    public void DefinitionEditsAndShadowCreationRecalculateButKeepNameIdentities()
    {
        var book = NamedBook(); using var session = new DocumentSession(book);
        session.Execute(new DefineNameCommand(new("GlobalSource", "=Other!$A$1"))); Calculate(session);
        Assert.Equal(100, session.Snapshot.Sheets[0].GetCell(1, 1)!.Value.NumberValue); Assert.Equal(200, session.Snapshot.Sheets[0].GetCell(0, 2)!.Value.NumberValue);
        session.Execute(new DefineNameCommand(new("GlobalOther", "=Source!$A$3", book.Sheets[0].Id))); Calculate(session);
        Assert.Equal(30, session.Snapshot.Sheets[0].GetCell(5, 1)!.Value.NumberValue);
        Assert.False(session.Execute(new DefineNameCommand(new("GlobalOther", "=Source!$A$3", book.Sheets[0].Id))).Changed);
    }
    [Fact]
    public void GlobalDeletionRetainsFormulaTextAsNameErrorAndUndoRestores()
    {
        var book = NamedBook(); using var session = new DocumentSession(book);
        session.Execute(new DeleteDefinedNameCommand("GlobalOther", null)); Calculate(session);
        Assert.Equal("=GlobalOther", session.Snapshot.Sheets[0].GetCell(5, 1)!.Content.FormulaText);
        Assert.Equal(CellErrorCode.Name, session.Snapshot.Sheets[0].GetCell(5, 1)!.Value.ErrorCode);
        session.Undo(); Calculate(session); Assert.Equal(100, session.Snapshot.Sheets[0].GetCell(5, 1)!.Value.NumberValue);
    }
    [Theory]
    [InlineData("R")]
    [InlineData("C")]
    [InlineData("R1C1")]
    [InlineData("A1")]
    [InlineData("bad name")]
    public void InvalidNamesNeverEnterHistory(string name)
    {
        using var session = new DocumentSession(NamedBook()); var before = session.Snapshot;
        Assert.Throws<ArgumentException>(() => session.Execute(new DefineNameCommand(new(name, "=1")))); Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }
    [Fact]
    public void AmbiguousRenameAndQualifiedLocalDeletionRejectAtomically()
    {
        using var session = new DocumentSession(NamedBook()); var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new RenameDefinedNameCommand("GlobalSource", null, "Local")));
        Assert.Throws<NotSupportedException>(() => session.Execute(new DeleteDefinedNameCommand("Rate", before.Sheets[0].Id)));
        Assert.Throws<ArgumentException>(() => session.Execute(new CopySheetCommand(before.Sheets[0].Id, "Other", 1)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }
    [Fact]
    public void RenameAndDefinitionEditAreOneTransactionAndInvalidReplacementDoesNotPartiallyRename()
    {
        var book = NamedBook(); using var session = new DocumentSession(book);
        Assert.Throws<ArgumentException>(() => session.Execute(new EditDefinedNameCommand("GlobalSource", null, "Amount", "=A1")));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
        session.Execute(new EditDefinedNameCommand("GlobalSource", null, "Amount", "=Other!$A$1")); Calculate(session);
        Assert.Equal(100, session.Snapshot.Sheets[0].GetCell(1, 1)!.Value.NumberValue);
        Assert.Equal("=Amount", session.Snapshot.Sheets[0].GetCell(1, 1)!.Content.FormulaText);
        session.Undo(); Assert.Equal(book.DefinedNames, session.Snapshot.DefinedNames); Assert.False(session.CanUndo);
    }
    [Fact]
    public void XlsxGroupsCopiedLocalNamesBeforeGlobalsForExcelBinding()
    {
        using var session = new DocumentSession(NamedBook()); session.Execute(new CopySheetCommand(session.Snapshot.Sheets[0].Id, "Copy", 1));
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            new TypedWorkbookStorage().Save(path, session.Snapshot);
            using var zip = System.IO.Compression.ZipFile.OpenRead(path); using var stream = zip.GetEntry("xl/workbook.xml")!.Open();
            var xml = System.Xml.Linq.XDocument.Load(stream); System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var names = xml.Root!.Element(ns + "definedNames")!.Elements().ToArray();
            var rate = names.Where(n => (string?)n.Attribute("name") == "Rate").ToArray();
            Assert.Equal(new int?[] { 1, 0, null }, rate.Select(n => (int?)n.Attribute("localSheetId")));
            Assert.Equal(2, Array.IndexOf(names, rate[^1]) - Array.IndexOf(names, rate[0]));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void UnusedImportedUnknownNamePreventsCopyAndRenameWithoutPartialHistory()
    {
        var book = NamedBook();
        book = book.WithDefinedNames(book.DefinedNames.Add(new("UnusedUnknown", "=FUTUREFUNC(Source!$A$1)")));
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, book);
            using var session = new DocumentSession(storage.Open(path)); var before = session.Snapshot;
            Assert.Throws<NotSupportedException>(() => session.Execute(new CopySheetCommand(before.Sheets[0].Id, "Copy", 1)));
            Assert.Throws<NotSupportedException>(() => session.Execute(new RenameDefinedNameCommand("Rate", null, "NewRate")));
            Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
        }
        finally { File.Delete(path); }
    }
}
