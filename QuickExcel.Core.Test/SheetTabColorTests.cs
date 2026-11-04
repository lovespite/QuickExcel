using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Domain;
using QuickExcel.Calculation;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class SheetTabColorTests
{
    [Fact]
    public void ColorAndResetAreUndoableWithoutInvalidatingCalculationOrOtherSheets()
    {
        using var session = new DocumentSession(Book(Sheet("First", (0, 0, F("=1+2"))), Sheet("Other")));
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        var before = session.Snapshot; var id = before.Sheets[0].Id;
        var result = session.Execute(new SetSheetTabColorCommand(id, " #0f76ae "));
        Assert.True(result.Changed); Assert.True(result.CalculationUnchanged); Assert.True(session.IsDirty);
        Assert.Equal("FF0F76AE", session.Snapshot.Sheets[0].TabColor);
        Assert.Equal(3, session.Snapshot.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal(CalculationState.Fresh, session.Snapshot.Sheets[0].GetCell(0, 0)!.Calculation!.State);
        Assert.Null(session.Snapshot.Sheets[1].TabColor);
        var revision = session.Revision;
        Assert.False(session.Execute(new SetSheetTabColorCommand(id, "ff0f76ae")).Changed);
        Assert.Equal(revision, session.Revision);
        session.Execute(new SetSheetTabColorCommand(id, null)); Assert.Null(session.Snapshot.Sheets[0].TabColor);
        Assert.False(session.Execute(new SetSheetTabColorCommand(id, null)).Changed);
        Assert.True(session.Undo()); Assert.Equal("FF0F76AE", session.Snapshot.Sheets[0].TabColor);
        Assert.True(session.Undo()); Assert.Null(session.Snapshot.Sheets[0].TabColor); Assert.False(session.CanUndo);
        Assert.True(session.Redo()); Assert.Equal("FF0F76AE", session.Snapshot.Sheets[0].TabColor);
        Assert.True(session.Redo()); Assert.Null(session.Snapshot.Sheets[0].TabColor);
    }

    [Theory]
    [InlineData("")]
    [InlineData("red")]
    [InlineData("12345")]
    [InlineData("##123456")]
    [InlineData("FF12345G")]
    [InlineData("FF1234567")]
    public void InvalidColorCannotChangeDocumentOrHistory(string color)
    {
        using var session = new DocumentSession(Book(Sheet("Sheet"))); var before = session.Snapshot;
        Assert.Throws<ArgumentException>(() => session.Execute(new SetSheetTabColorCommand(before.ActiveSheetId, color)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo); Assert.False(session.IsDirty);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsq")]
    public void ColorCopyAndResetRoundTripWithoutLeavingTabColorXml(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "." + extension);
        try
        {
            var storage = new TypedWorkbookStorage();
            using var session = new DocumentSession(Book(Sheet("First"), Sheet("Other")));
            var id = session.Snapshot.Sheets[0].Id;
            session.Execute(new SetSheetTabColorCommand(id, "#123abc"));
            var copy = new CopySheetCommand(id, "Copy", 1); session.Execute(copy);
            storage.Save(path, session.Snapshot);
            var reopened = storage.Open(path);
            Assert.Equal(new string?[] { "FF123ABC", "FF123ABC", null }, reopened.Sheets.Select(s => s.TabColor));
            using var restored = new DocumentSession(reopened);
            restored.Execute(new SetSheetTabColorCommand(reopened.Sheets[0].Id, null));
            storage.Save(path, restored.Snapshot);
            Assert.Equal(new string?[] { null, "FF123ABC", null }, storage.Open(path).Sheets.Select(s => s.TabColor));
            using var package = ZipFile.OpenRead(path);
            using var workbook = new MemoryStream();
            if (extension == "xlsq") { using var stream = package.GetEntry("workbook.xlsx")!.Open(); stream.CopyTo(workbook); }
            using var inner = extension == "xlsq" ? new ZipArchive(workbook, ZipArchiveMode.Read) : null;
            using var sheet = (inner ?? package).GetEntry("xl/worksheets/sheet1.xml")!.Open();
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            Assert.Empty(XDocument.Load(sheet).Descendants(ns + "tabColor"));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void ResetOfImportedColorPreservesMetadataAndProtectedSheetContents()
    {
        var sheet = new WorksheetSnapshot(SheetId.New(), "Imported", tabColor: "FFABCDEF", showGridLines: false,
            defaultRowHeight: 24, view: new() { Zoom = 125 }, data: new() { Protection = new() { Sheet = true } });
        using var session = new DocumentSession(Book(sheet, Sheet("Other")));
        session.Execute(new SetSheetTabColorCommand(sheet.Id, null));
        var after = session.Snapshot.GetSheet(sheet.Id);
        Assert.Null(after.TabColor); Assert.Equal(sheet.View, after.View); Assert.Same(sheet.Data, after.Data);
        Assert.False(after.ShowGridLines); Assert.Equal(24, after.DefaultRowHeight);
        Assert.True(session.Undo()); Assert.Equal("FFABCDEF", session.Snapshot.GetSheet(sheet.Id).TabColor);
    }
}
