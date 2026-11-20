using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class NamedRangeManagementTests
{
    [Fact]
    public void EditCommentAndRenameAreOneUndoableChangeAndKeepScriptIdentity()
    {
        var definition = new DefinedName("Amount", "=Data!$A$1", Comment: "Original");
        var book = Book(Sheet("Data", (0, 0, N(42)), (0, 1, F("=Amount")))).WithDefinedNames([definition]);
        using var session = new DocumentSession(book);
        session.Execute(new EditDefinedNameCommand("Amount", null, "Total", "=Data!$A$1", "Updated"));
        var edited = Assert.Single(session.Snapshot.DefinedNames);
        Assert.Equal("Total", edited.Name); Assert.Equal("Updated", edited.Comment); Assert.Equal(definition.ScriptId, edited.ScriptId);
        Assert.Equal("=Total", session.Snapshot.Sheets[0].GetCell(0, 1)!.Content.FormulaText);
        Assert.True(session.Undo()); Assert.Equal(definition, Assert.Single(session.Snapshot.DefinedNames)); Assert.False(session.CanUndo);
        Assert.True(session.Redo());
        session.Execute(new EditDefinedNameCommand("Total", null, "Total", "=Data!$A$2"));
        Assert.Equal("Updated", Assert.Single(session.Snapshot.DefinedNames).Comment);
        session.Execute(new EditDefinedNameCommand("Total", null, "Total", "=Data!$A$2", ""));
        Assert.Equal("", Assert.Single(session.Snapshot.DefinedNames).Comment);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void CommentsRoundTripAndCommentOnlyEditsAreSaved(string extension)
    {
        var sheet = Sheet("配额表", (0, 0, N(42)));
        var comment = "配额说明\r\n第二行\t_x0041_ & <remarks>";
        var book = Book(sheet).WithDefinedNames([new("配额", "=配额表!$A$1", Comment: comment), new("LocalAmount", "=配额表!$A$1", sheet.Id, "Local")]);
        var path = Path.Combine(Path.GetTempPath(), "QuickExcel-Names-" + Guid.NewGuid() + "." + extension);
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, book);
            var restored = storage.Open(path);
            Assert.False(restored.FileState!.FormulaReadOnly);
            Assert.Empty(restored.FileState.CompatibilityDiagnostics);
            Assert.Equal(comment, restored.DefinedNames.Single(n => n.Scope is null).Comment);
            Assert.Equal("Local", restored.DefinedNames.Single(n => n.Scope is not null).Comment);
            using var session = new DocumentSession(restored);
            session.Execute(new EditDefinedNameCommand("配额", null, "配额", "=配额表!$A$1", "Revised"));
            storage.Save(path, session.Snapshot);
            Assert.Equal("Revised", storage.Open(path).DefinedNames.Single(n => n.Scope is null).Comment);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void CopiedAndRecreatedNamesHaveIndependentIdentities()
    {
        var sheet = Sheet("Data"); var name = new DefinedName("Amount", "=Data!$A$1", sheet.Id, "Description");
        using var session = new DocumentSession(Book(sheet).WithDefinedNames([name]));
        var copy = new CopySheetCommand(sheet.Id, "Copy", 1); session.Execute(copy);
        var copied = session.Snapshot.DefinedNames.Single(n => n.Scope == copy.Id);
        Assert.Equal("Description", copied.Comment); Assert.NotEqual(name.ScriptId, copied.ScriptId);
        using var replay = new DocumentSession(Book(sheet).WithDefinedNames([name]));
        replay.Execute(copy);
        Assert.Equal(copied.ScriptId, replay.Snapshot.DefinedNames.Single(n => n.Scope == copy.Id).ScriptId);
        session.Execute(new DeleteDefinedNameCommand("Amount", sheet.Id));
        session.Execute(new DefineNameCommand(new("Amount", "=Data!$A$1", sheet.Id, "Description")));
        Assert.NotEqual(name.ScriptId, session.Snapshot.DefinedNames.Single(n => n.Scope == sheet.Id).ScriptId);
    }

    [Fact]
    public void OversizedCommentDoesNotPartiallyRenameOrEnterHistory()
    {
        var book = Book(Sheet("Data")).WithDefinedNames([new("Amount", "=Data!$A$1")]);
        using var session = new DocumentSession(book);
        Assert.Throws<ArgumentException>(() => session.Execute(new EditDefinedNameCommand("Amount", null, "Total", "=Data!$A$1", new string('x', 32768))));
        Assert.Same(book, session.Snapshot); Assert.False(session.CanUndo);
    }
}
