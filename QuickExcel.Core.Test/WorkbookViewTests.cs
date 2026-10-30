using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class WorkbookViewTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "QuickExcel-views-" + Guid.NewGuid().ToString("N"));
    private readonly TypedWorkbookStorage _storage = new();
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public WorkbookViewTests() => Directory.CreateDirectory(_folder);
    public void Dispose() => Directory.Delete(_folder, true);
    private string PathFor(string name) => Path.Combine(_folder, name);
    private static WorkbookSnapshot Model() => new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
        [new(SheetId.New(), "Hidden", visibility: SheetVisibility.Hidden),
         new(SheetId.New(), "Visible", cells: [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(7))))],
             view: new() { FrozenRows = 2, FrozenColumns = 1, FrozenSplit = true, Zoom = 125, NormalZoom = 100,
                 PaneTopLeftCell = new(5, 3), ActivePane = SheetPane.BottomRight,
                 TopRight = new(default, new(0, 0, 0, 0), true, true), BottomLeft = SheetSelection.At(new(5, 0)),
                 BottomRight = new(new(7, 4), new(7, 4, 9, 5)) }),
         new(SheetId.New(), "Secret", visibility: SheetVisibility.VeryHidden),
         new(SheetId.New(), "Other")], window: new(20, -10, 25000, 12000));
    private static byte[] Part(string path, string part)
    { using var zip = ZipFile.OpenRead(path); using var stream = zip.GetEntry(part)!.Open(); using var memory = new MemoryStream(); stream.CopyTo(memory); return memory.ToArray(); }
    private static void WritePart(string path, string part, byte[] bytes)
    { using var zip = ZipFile.Open(path, ZipArchiveMode.Update); zip.GetEntry(part)?.Delete(); using var stream = zip.CreateEntry(part).Open(); stream.Write(bytes); }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void ViewsAndVisibilityRoundTripAfterValueAndStyleEdits(string extension)
    {
        var path = PathFor("source" + extension); var source = Model(); _storage.Save(path, source);
        var opened = _storage.Open(path); Assert.Empty(opened.FileState!.CompatibilityIssues);
        Assert.Equal("Visible", opened.GetSheet(opened.ActiveSheetId).Name);
        Assert.Equal(source.Window, opened.Window); Assert.Equal(source.Sheets[1].View, opened.Sheets[1].View);
        Assert.Equal(SheetVisibility.VeryHidden, opened.Sheets[2].Visibility);
        using var session = new DocumentSession(opened);
        session.Execute(new SetCellsCommand([new(new(opened.Sheets[1].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(9)))]));
        session.Execute(new SetCellStyleCommand([new(opened.Sheets[1].Id, 0, 0)], new(Bold: true)));
        _storage.Save(path, session.Snapshot);
        var saved = _storage.Open(path);
        Assert.Equal(source.Sheets.Select(s => s.View), saved.Sheets.Select(s => s.View));
        Assert.Equal(source.Sheets.Select(s => s.Visibility), saved.Sheets.Select(s => s.Visibility));
        Assert.Equal(9, saved.Sheets[1].GetCell(0, 0)!.Value.NumberValue);
        if (extension == ".xlsq")
        {
            Assert.Equal(source.Sheets.Select(s => s.Id), saved.Sheets.Select(s => s.Id));
            Assert.Contains("workbook-views", Encoding.UTF8.GetString(Part(path, "manifest.json")));
        }
    }

    [Fact]
    public void NavigationIsNotDirtyWhileFreezeZoomAndVisibilityAreUndoable()
    {
        var source = Model(); using var session = new DocumentSession(source);
        session.UpdateNavigation(source.Sheets[3].Id, new(9, 7));
        Assert.False(session.IsDirty); Assert.False(session.CanUndo); Assert.Equal(0, session.Revision);
        Assert.Equal(source.Sheets[3].Id, session.Snapshot.ActiveSheetId);
        Assert.Equal(new(9, 7), session.Snapshot.Sheets[3].View.ActiveCell);
        session.UpdateNavigation(source.Sheets[3].Id, new(9, 7), new(5, 3));
        Assert.Equal(new(5, 3), session.Snapshot.Sheets[3].View.TopLeftCell);
        Assert.False(session.IsDirty); Assert.False(session.CanUndo);
        var oldView = session.Snapshot.Sheets[1].View;
        var result = session.Execute(new SetSheetViewCommand(source.Sheets[1].Id, oldView with { Zoom = 150 }));
        Assert.True(result.CalculationUnchanged); Assert.True(session.IsDirty);
        Assert.True(session.Undo()); Assert.Equal(oldView, session.Snapshot.Sheets[1].View);
        Assert.True(session.Redo()); Assert.Equal(150, session.Snapshot.Sheets[1].View.Zoom);
        session.Execute(new SetSheetVisibilityCommand(source.Sheets[3].Id, SheetVisibility.Hidden));
        Assert.Equal(source.Sheets[1].Id, session.Snapshot.ActiveSheetId);
        Assert.Throws<ArgumentException>(() => session.Execute(new SetSheetVisibilityCommand(source.Sheets[1].Id, SheetVisibility.Hidden)));
        Assert.Throws<InvalidOperationException>(() => session.UpdateNavigation(source.Sheets[2].Id, default));
        Assert.True(session.Undo()); Assert.Equal(SheetVisibility.Visible, session.Snapshot.Sheets[3].Visibility);
        var path = PathFor("navigation.xlsx"); _storage.Save(path, session.Snapshot);
        var reopened = _storage.Open(path);
        Assert.Equal("Other", reopened.GetSheet(reopened.ActiveSheetId).Name);
    }

    [Theory]
    [InlineData("split")]
    [InlineData("multiple")]
    [InlineData("out-of-range")]
    public void UnsupportedViewsRemainProtectedAndUnchangedSaveKeepsTheirXml(string fault)
    {
        var path = PathFor("unsupported.xlsx"); _storage.Save(path, Model());
        var xml = XDocument.Parse(Encoding.UTF8.GetString(Part(path, "xl/worksheets/sheet2.xml")));
        var view = xml.Root!.Element(Main + "sheetViews")!.Element(Main + "sheetView")!;
        if (fault == "split") view.Element(Main + "pane")!.SetAttributeValue("state", "split");
        else if (fault == "multiple") view.Elements(Main + "selection").Last().SetAttributeValue("sqref", "E8:F10 H12");
        else view.SetAttributeValue("topLeftCell", "XFD1048576");
        WritePart(path, "xl/worksheets/sheet2.xml", Encoding.UTF8.GetBytes(xml.ToString()));
        var original = File.ReadAllBytes(path); var opened = _storage.Open(path);
        Assert.Contains(opened.FileState!.CompatibilityIssues, i => i.Contains("视图仅保留"));
        var copy = PathFor("unchanged.xlsx"); _storage.Save(copy, opened); Assert.Equal(original, File.ReadAllBytes(copy));
        using var session = new DocumentSession(opened);
        session.Execute(new SetCellsCommand([new(new(opened.Sheets[1].Id, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(9)))]));
        Assert.Throws<WorkbookCompatibilityException>(() => _storage.Save(path, session.Snapshot));
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Fact]
    public void SheetViewWithoutExplicitWorkbookWindowCanBePreservedInXlsq()
    {
        var source = PathFor("implicit-window.xlsx");
        _storage.Save(source, new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(SheetId.New(), "Sheet", view: new WorksheetView { Zoom = 120 }.WithActiveCell(new(3, 2)))]));
        var xml = XDocument.Parse(Encoding.UTF8.GetString(Part(source, "xl/workbook.xml")));
        xml.Root!.Element(Main + "bookViews")!.Remove();
        WritePart(source, "xl/workbook.xml", Encoding.UTF8.GetBytes(xml.ToString()));
        var original = File.ReadAllBytes(source); var book = _storage.Open(source);
        var output = PathFor("implicit-window.xlsq"); _storage.Save(output, book);
        Assert.Equal(original, Part(output, "workbook.xlsx"));
        Assert.Equal(book.Sheets[0].View, _storage.Open(output).Sheets[0].View);
    }

    [Fact]
    public void InsertingBeforeFrozenBoundaryMatchesExcelWithoutMovingActiveCell()
    {
        var source = Model(); var sheet = source.Sheets[1];
        using var session = new DocumentSession(source);
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 0, 1));
        var view = session.Snapshot.GetSheet(sheet.Id).View;
        Assert.Equal(3, view.FrozenRows); Assert.Equal(3, view.PaneTopLeftCell.Row);
        Assert.Equal(sheet.View.ActiveCell, view.ActiveCell); Assert.Equal(sheet.View.Selection, view.Selection);
        Assert.True(session.Undo()); Assert.Equal(sheet.View, session.Snapshot.GetSheet(sheet.Id).View);
        session.Execute(new ShiftAxisCommand(sheet.Id, SheetAxis.Row, 0, 2, delete: true));
        view = session.Snapshot.GetSheet(sheet.Id).View;
        Assert.Equal(0, view.FrozenRows); Assert.Equal(1, view.FrozenColumns);
        Assert.Equal(SheetPane.TopRight, view.ActivePane); Assert.Equal(sheet.View.Selection, view.Selection);
        var path = PathFor("removed-axis.xlsx"); _storage.Save(path, session.Snapshot);
        Assert.Equal(view, _storage.Open(path).Sheets[1].View);
    }

    [Fact]
    public void OldV2OptionalViewMigratesFromRatioAndStableIdWithoutLosingSelections()
    {
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(SheetId.New(), "First"), new(SheetId.New(), "Second")]);
        var path = PathFor("old.xlsq"); _storage.Save(path, book);
        var workbookBytes = Part(path, "workbook.xlsx");
        using (var memory = new MemoryStream())
        {
            memory.Write(workbookBytes);
            using (var zip = new ZipArchive(memory, ZipArchiveMode.Update, true))
            {
                var entry = zip.GetEntry("xl/workbook.xml")!; XDocument xml; using (var input = entry.Open()) xml = XDocument.Load(input);
                xml.Root!.Element(Main + "bookViews")!.Remove(); entry.Delete(); using var output = zip.CreateEntry("xl/workbook.xml").Open(); xml.Save(output);
            }
            workbookBytes = memory.ToArray();
        }
        var state = JsonNode.Parse(Part(path, "state/document.json"))!;
        state["view"] = new JsonObject { ["activeSheetId"] = book.Sheets[1].Id.ToString(), ["selection"] = "C4:D5", ["zoom"] = 1.5 };
        var stateBytes = Encoding.UTF8.GetBytes(state.ToJsonString());
        var manifest = JsonNode.Parse(Part(path, "manifest.json"))!;
        manifest["workbook"]!["sha256"] = Convert.ToHexStringLower(SHA256.HashData(workbookBytes));
        manifest["state"]!["sha256"] = Convert.ToHexStringLower(SHA256.HashData(stateBytes));
        WritePart(path, "workbook.xlsx", workbookBytes); WritePart(path, "state/document.json", stateBytes);
        WritePart(path, "manifest.json", Encoding.UTF8.GetBytes(manifest.ToJsonString()));
        var loaded = _storage.Open(path); Assert.Equal(book.Sheets[1].Id, loaded.ActiveSheetId);
        Assert.Equal(150, loaded.Sheets[1].View.Zoom); Assert.Equal(new(3, 2, 4, 3), loaded.Sheets[1].View.Selection!.Range);
        _storage.Save(path, loaded);
        var saved = _storage.Open(path); Assert.Equal(loaded.Sheets[1].View, saved.Sheets[1].View);
        Assert.Contains("workbook-views", Encoding.UTF8.GetString(Part(path, "manifest.json")));
    }
}
