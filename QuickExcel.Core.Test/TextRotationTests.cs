using System.IO.Compression;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Core.Test;

public sealed class TextRotationTests
{
    [Theory]
    [InlineData(-90, 180)] [InlineData(-45, 135)] [InlineData(-17, 107)]
    [InlineData(0, 0)] [InlineData(17, 17)] [InlineData(45, 45)] [InlineData(90, 90)]
    public void SignedAnglesUsePortableXlsxValues(int degrees, int stored)
    {
        Assert.Equal(stored, TextRotationOptions.FromDegrees(degrees));
        Assert.Equal(degrees, TextRotationOptions.ToDegrees(stored));
        Assert.Throws<ArgumentOutOfRangeException>(() => TextRotationOptions.FromDegrees(-91));
        Assert.Throws<ArgumentOutOfRangeException>(() => TextRotationOptions.FromDegrees(91));
    }

    [Fact]
    public void RangeRotationPreservesRichTextFormulaAndMixedStylesAsOneHistoryEntry()
    {
        var rich = new RichCellText([new("中", new(Color: "FFFF0000")), new("文")]);
        var sheet = new WorksheetSnapshot(SheetId.New(), "Rotation", cells: [
            new(new(0, 0), new(CellContent.FromRichText(rich, 1))),
            new(new(1, 0), new(CellContent.FromFormula("=1+2", 2)))],
            rows: [new(2, new(StyleId: 1))], mergedCells: ["A3:B3"]);
        using var session = new DocumentSession(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet],
            styles: [new(), new() { Bold = true, WrapText = true }, new() { FillColor = "FFFFEECC", Horizontal = 3 }]));
        var before = session.Snapshot;
        session.Execute(new SetCellStyleCommand(Enumerable.Range(0, 3).Select(row => new CellAddress(sheet.Id, row, 0)), new(TextRotation: 255)));
        var changed = session.Snapshot;
        for (var row = 0; row < 3; row++)
        {
            var original = before.Styles[sheet.GetStyleId(row, 0)];
            Assert.Equal(original with { TextRotation = 255 }, changed.Styles[changed.Sheets[0].GetStyleId(row, 0)]);
        }
        Assert.Equal(rich, changed.Sheets[0].GetCell(0, 0)!.Content.RichText);
        Assert.Equal("=1+2", changed.Sheets[0].GetCell(1, 0)!.Content.FormulaText);
        Assert.Equal(sheet.MergedCells, changed.Sheets[0].MergedCells);
        Assert.True(session.Undo()); Assert.Equal(before.Styles.ToArray(), session.Snapshot.Styles.ToArray()); Assert.False(session.Undo());
        Assert.True(session.Redo()); Assert.Equal(changed.Styles.ToArray(), session.Snapshot.Styles.ToArray());
    }

    [Theory]
    [InlineData("xlsx")] [InlineData("xlsq")]
    public void EveryPresetAndCustomAngleRoundtripsWithoutChangingCellContent(string extension)
    {
        int[] rotations = [0, 45, 135, 255, 90, 180, 17, 107];
        var sheet = new WorksheetSnapshot(SheetId.New(), "Rotation", cells: rotations.Select((rotation, row) =>
            new KeyValuePair<CellPosition, CellSnapshot>(new(row, 0), new(CellContent.FromLiteral(CellValue.FromText("中文 Ab"), row + 1)))));
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet],
            styles: new[] { new StyleDefinition() }.Concat(rotations.Select(r => new StyleDefinition { TextRotation = r })));
        var directory = Path.Combine(Path.GetTempPath(), "QuickExcel-rotation-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "rotation." + extension); var storage = new TypedWorkbookStorage();
            if (extension == "xlsx") storage.Export(path, book); else storage.Save(path, book);
            var reopened = storage.Open(path);
            Assert.Empty(reopened.FileState!.CompatibilityIssues);
            for (var row = 0; row < rotations.Length; row++)
            {
                Assert.Equal(rotations[row], reopened.Styles[reopened.Sheets[0].GetStyleId(row, 0)].TextRotation);
                Assert.Equal("中文 Ab", reopened.Sheets[0].GetCell(row, 0)!.Value.TextValue);
            }
            if (extension == "xlsx")
            {
                using var zip = ZipFile.OpenRead(path); using var stream = zip.GetEntry("xl/styles.xml")!.Open();
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var xfs = XDocument.Load(stream).Root!.Element(ns + "cellXfs")!.Elements().ToArray();
                for (var row = 0; row < rotations.Length; row++)
                    Assert.Equal(rotations[row], (int?)xfs[sheet.GetStyleId(row, 0)].Element(ns + "alignment")?.Attribute("textRotation") ?? 0);
            }
        }
        finally { Directory.Delete(directory, true); }
    }
}
