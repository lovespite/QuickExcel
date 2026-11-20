using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Calculation;
using QuickExcel.Core.Test.Fixtures;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class SupportedCopySaveTests : IDisposable
{
    private readonly string _directory = Path.Combine(AppContext.BaseDirectory, "test-data", "supported-copy-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    public SupportedCopySaveTests() => Directory.CreateDirectory(_directory);
    public void Dispose() => Directory.Delete(_directory, true);
    private string PathFor(string name) => Path.Combine(_directory, name);
    private WorkbookServices Services => new(_storage, new WorkbookCalculator(new NoScripts()));
    private WorkbookWorkspace Workspace() => new(Services, new Files(_storage));
    private sealed class Files(TypedWorkbookStorage storage) : IWorkbookFileAccess
    {
        public FileVersion ReadVersion(string path) => FileVersions.Read(path);
        public OpenedWorkbook Open(string path) => new(storage.Open(path), FileVersions.Read(path));
    }

    private string LossySource()
    {
        var path = PathFor("source.xlsx");
        File.WriteAllBytes(path, WorkbookPackageFixture.Xlsx("", new Dictionary<string, string>
        {
            ["xl/worksheets/sheet1.xml"] = """
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData><row r="1"><c r="A1"><v>10</v></c></row></sheetData>
                  <customSheetViews/>
                </worksheet>
                """
        }));
        return path;
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public async Task SupportedCopyCanSaveEditUndoAndRedoWithoutReopening(string extension)
    {
        var source = LossySource(); var original = File.ReadAllBytes(source);
        using var workspace = Workspace(); await workspace.OpenAsync(new(source));
        var document = workspace.Document; var editor = workspace.Editor;
        var session = document.Session.SessionId; var sheetId = document.Snapshot.ActiveSheetId;
        document.SetCell(0, 0, 0, "20"); document.SetCell(0, 0, 0, "30");
        Assert.True(document.Undo());
        Assert.NotEmpty(TypedWorkbookStorage.SaveIssues(document.Snapshot));
        var target = PathFor("supported." + extension);
        await workspace.SaveAsync(new(target), discardUnsupported: true);

        Assert.Same(document, workspace.Document); Assert.Same(editor, workspace.Editor);
        Assert.Equal(session, document.Session.SessionId); Assert.Equal(sheetId, document.Snapshot.ActiveSheetId);
        Assert.True(document.CanUndo); Assert.True(document.CanRedo);
        Assert.False(document.IsDirty); Assert.Equal(target, workspace.FilePath);
        Assert.Equal(target, document.Snapshot.FileState!.SourcePath);
        Assert.Empty(document.Snapshot.FileState.CompatibilityIssues);
        Assert.Null(document.Snapshot.FileState.OriginalModel.FileState);
        Assert.Equal(sheetId, document.Snapshot.FileState.OriginalModel.ActiveSheetId);
        await SaveNormally("20");
        Assert.True(document.Undo()); await SaveNormally("10");
        Assert.True(document.Redo()); await SaveNormally("20");
        Assert.True(document.Redo()); await SaveNormally("30");
        document.SetCell(0, 0, 0, "40"); await SaveNormally("40");
        Assert.Equal(original, File.ReadAllBytes(source));

        async Task SaveNormally(string value)
        {
            Assert.Equal(value, document.GetCellInput(0, 0, 0));
            Assert.Empty(TypedWorkbookStorage.SaveIssues(document.Snapshot));
            await workspace.SaveAsync(new(target));
            Assert.False(document.IsDirty);
            Assert.Equal(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture), _storage.Open(target).Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        }
    }

    [Fact]
    public async Task FailedSupportedCopyKeepsSourceMetadataDirtyStateAndHistory()
    {
        var source = LossySource(); var original = File.ReadAllBytes(source);
        using var workspace = Workspace(); await workspace.OpenAsync(new(source));
        workspace.Document.SetCell(0, 0, 0, "20");
        var snapshot = workspace.Document.Snapshot; var version = workspace.Version;
        await Assert.ThrowsAsync<DirectoryNotFoundException>(() => workspace.SaveAsync(new(PathFor("missing/copy.xlsq")), discardUnsupported: true));
        Assert.Same(snapshot, workspace.Document.Snapshot);
        Assert.Equal(source, workspace.FilePath); Assert.Equal(version, workspace.Version);
        Assert.True(workspace.Document.IsDirty); Assert.True(workspace.Document.CanUndo);
        Assert.NotEmpty(TypedWorkbookStorage.SaveIssues(workspace.Document.Snapshot));
        Assert.Equal(original, File.ReadAllBytes(source));
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public async Task CopyThatPreservesUnsupportedContentStillRequiresConsentAfterAnEdit(string extension)
    {
        var source = LossySource(); using var workspace = Workspace();
        await workspace.OpenAsync(new(source));
        var target = PathFor("preserved." + extension);
        await workspace.SaveAsync(new(target));
        Assert.Equal(target, workspace.Document.Snapshot.FileState!.SourcePath);
        Assert.NotEmpty(workspace.Document.Snapshot.FileState.CompatibilityIssues);
        Assert.Empty(TypedWorkbookStorage.SaveIssues(workspace.Document.Snapshot));
        var saved = File.ReadAllBytes(target);
        workspace.Document.SetCell(0, 0, 0, "20");
        Assert.NotEmpty(TypedWorkbookStorage.SaveIssues(workspace.Document.Snapshot));
        await Assert.ThrowsAsync<WorkbookCompatibilityException>(() => workspace.SaveAsync(new(target)));
        Assert.True(workspace.Document.IsDirty); Assert.Equal(saved, File.ReadAllBytes(target));
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void ExplicitOverwriteRefreshesTheFacadeSaveBaseline(string extension)
    {
        var source = LossySource(); var target = PathFor("overwrite." + extension);
        _storage.Save(target, _storage.Open(source));
        using var document = WorkbookDocument.Open(target, Services);
        document.SetCell(0, 0, 0, "20");
        document.SaveChecked(target, FileVersions.Read(target), discardUnsupported: true, overwriteOriginal: true);
        Assert.Empty(TypedWorkbookStorage.SaveIssues(document.Snapshot));
        document.Save(target);
        Assert.True(document.Undo()); Assert.Empty(TypedWorkbookStorage.SaveIssues(document.Snapshot));
        document.Save(target);
        Assert.Equal(10, _storage.Open(target).Sheets[0].GetCell(0, 0)!.Value.NumberValue);
    }

    [Fact]
    public void MigratedCopyNoLongerCarriesTheV1SaveRestriction()
    {
        var source = PathFor("legacy.xlsq");
        WorkbookPackageFixture.WriteV1(source, WorkbookPackageFixture.Xlsx("<row r=\"1\"><c r=\"A1\"><v>10</v></c></row>"));
        // An explicit script rewrite is required before a v1 document can be saved.
        using var document = new WorkbookDocument(_storage.Open(source).WithScriptModules(ScriptModuleSet.Single("")), Services);
        Assert.True(document.Snapshot.FileState!.RequiresMigration);
        var target = PathFor("migrated.xlsq"); document.Save(target);
        Assert.False(document.Snapshot.FileState!.RequiresMigration);
        Assert.Equal(2, document.Snapshot.FileState.SourceFormatVersion);
        Assert.Equal(target, document.Snapshot.FileState.SourcePath);
        document.SetCell(0, 0, 0, "20"); document.Save(target);
        Assert.False(document.IsDirty);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void FormulaReadOnlyCopiesRetainSourceNavigationAndStableIdentities(string extension)
    {
        var source = PathFor("formula.xlsx");
        File.WriteAllBytes(source, WorkbookPackageFixture.Xlsx("""
            <row r="1"><c r="A1"><f t="array" ref="A1:A2">B1:B2*2</f><v>2</v></c><c r="B1"><v>1</v></c></row>
            <row r="2"><c r="A2"><v>4</v></c><c r="B2"><v>2</v></c></row>
            """));
        using var document = WorkbookDocument.Open(source, Services);
        Assert.True(document.Snapshot.FileState!.FormulaReadOnly);
        var id = document.Snapshot.ActiveSheetId; var sourceView = document.Snapshot.Sheets[0].View;
        document.Session.UpdateNavigation(id, new(20, 4), new(15, 2));
        var target = PathFor("formula-copy." + extension); document.Save(target);
        Assert.Equal(id, document.Snapshot.FileState!.OriginalModel.ActiveSheetId);
        Assert.Equal(sourceView, document.Snapshot.FileState.OriginalModel.Sheets[0].View);
        Assert.True(document.Snapshot.FileState.FormulaReadOnly);
        Assert.Empty(TypedWorkbookStorage.SaveIssues(document.Snapshot));
        document.Save(target);
        Assert.Equal(sourceView, _storage.Open(target).Sheets[0].View);
        Assert.Throws<InvalidOperationException>(() => document.SetCell(0, 0, 0, "99"));
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void RebindingSupportedCopyKeepsOpaqueThreadAttachmentsUsable(string extension)
    {
        var source = LossySource();
        AddThreadParts(source);
        using var document = WorkbookDocument.Open(source, Services);
        var id = document.Snapshot.ActiveSheetId;
        Assert.Single(document.Snapshot.FileState!.PreservedAttachments);
        var threadBytes = Part(document.Snapshot.FileState.OriginalWorkbook.ToArray(), "xl/threadedComments/threadedComment1.xml");
        document.SetCell(0, 0, 0, "20");
        var target = PathFor("thread-copy." + extension); document.Save(target, discardUnsupported: true);
        Assert.Empty(TypedWorkbookStorage.SaveIssues(document.Snapshot));
        Assert.Single(document.Snapshot.FileState!.PreservedAttachments);
        Assert.Equal(id, document.Snapshot.FileState.OriginalModel.ActiveSheetId);
        document.SetCell(0, 0, 0, "30"); document.Save(target);
        Assert.True(document.Undo()); document.Save(target);
        Assert.Equal(threadBytes, Part(_storage.Open(target).FileState!.OriginalWorkbook.ToArray(), "xl/threadedComments/threadedComment1.xml"));
        Assert.Throws<NotSupportedException>(() => document.Session.Execute(new ShiftAxisCommand(id, SheetAxis.Row, 0, 1)));
    }

    private static byte[] Part(byte[] bytes, string path)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes)); using var source = zip.GetEntry(path)!.Open();
        using var result = new MemoryStream(); source.CopyTo(result); return result.ToArray();
    }

    private static void AddThreadParts(string path)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        void Write(string name, string text)
        {
            zip.GetEntry(name)?.Delete(); using var writer = new StreamWriter(zip.CreateEntry(name).Open()); writer.Write(text);
        }
        var typesEntry = zip.GetEntry("[Content_Types].xml")!;
        XDocument types; using (var stream = typesEntry.Open()) types = XDocument.Load(stream);
        var ns = types.Root!.Name.Namespace;
        types.Root.Add(new XElement(ns + "Override", new XAttribute("PartName", "/xl/threadedComments/threadedComment1.xml"), new XAttribute("ContentType", "application/vnd.ms-excel.threadedcomments+xml")),
            new XElement(ns + "Override", new XAttribute("PartName", "/xl/persons/person.xml"), new XAttribute("ContentType", "application/vnd.ms-excel.person+xml")));
        Write("[Content_Types].xml", types.ToString());
        Write("xl/_rels/workbook.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
              <Relationship Id="people" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>
            </Relationships>
            """);
        Write("xl/worksheets/_rels/sheet1.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="threads" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
            </Relationships>
            """);
        Write("xl/persons/person.xml", """
            <personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
              <person id="{11111111-1111-4111-8111-111111111111}" displayName="Fixture Author"/>
            </personList>
            """);
        Write("xl/threadedComments/threadedComment1.xml", """
            <ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
              <threadedComment ref="A1" dT="2026-11-03T00:00:00Z" personId="{11111111-1111-4111-8111-111111111111}" id="{22222222-2222-4222-8222-222222222222}"><text>Retained discussion</text></threadedComment>
            </ThreadedComments>
            """);
    }
}
