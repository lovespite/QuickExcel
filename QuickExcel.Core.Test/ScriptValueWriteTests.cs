using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class ScriptValueWriteTests
{
    private static readonly byte[] Pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aA1sAAAAASUVORK5CYII=");

    private static WorksheetSnapshot SheetWithAttachments(bool inCellImage = false) => new(SheetId.New(), "Data",
        [new(new(0, 0), N(42))], mergedCells: ["A1:B2"],
        images: [new SheetImage(Pixel, "png", 3, 3, placeInCell: inCellImage)],
        annotations: new(links: [new(new(0, 0), new("https://example.com"))],
            notes: [new(new(0, 0), new("Author", new([new("Keep this note")])))]));

    private static WorkbookDocument Document(string source, WorksheetSnapshot sheet) => new(
        Book(sheet).WithScriptModules(M4ActionTests.Actions(source)),
        new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));

    private static void AssertAttachments(WorksheetSnapshot expected, WorksheetSnapshot actual)
    {
        Assert.Equal(expected.MergedCells, actual.MergedCells);
        Assert.Equal(expected.Images, actual.Images);
        Assert.Same(expected.Annotations, actual.Annotations);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SheetActivateWritesOrdinaryAndMergedAnchorCellsWithBothApisAndOneUndo(bool inCellImage)
    {
        var sheet = SheetWithAttachments(inCellImage);
        using var doc = Document("""
            $.events.on('sheet.activate', ({patch,event}) => {
              const s=$.getWorksheet(event.sheetId);
              s.getCell('A1').setValue(7);
              if(s.getCell('A1').getValue()!==7) throw Error('anchor read');
              s.getRange('C3:D3').setValues([[8,9]]);
              patch.setFormula(s.getId(),'A1','=C3+D3');
              s.getRange('A3:B3').setFormulas([['=C3*2','=D3*2']]);
              patch.setValue(s.getId(),'C4',10);
              if(s.getCell('A1').getValue()!==17) throw Error('formula read');
              s.getCell('A1').getFormat().getFill().setColor('#ff0000');
            });
            """, sheet);
        doc.EnableScriptsAndCalculate();
        var revision = doc.Snapshot.Revision;
        doc.DispatchScriptEvents();
        Assert.Equal(revision + 1, doc.Snapshot.Revision);
        Assert.Equal("=C3+D3", doc.Snapshot.Sheets[0].GetCell(0, 0)!.Content.FormulaText);
        Assert.Equal("8", doc.GetCellValue(0, 2, 2));
        Assert.Equal("10", doc.GetCellValue(0, 3, 2));
        Assert.Equal("=D3*2", doc.Snapshot.Sheets[0].GetCell(2, 1)!.Content.FormulaText);
        Assert.Null(doc.Snapshot.Sheets[0].GetCell(0, 1));
        AssertAttachments(sheet, doc.Snapshot.Sheets[0]);
        Assert.False(doc.Session.HasScriptEvents);
        Assert.True(doc.Undo()); Assert.False(doc.CanUndo);
        Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.Null(doc.Snapshot.Sheets[0].GetCell(2, 2));
        AssertAttachments(sheet, doc.Snapshot.Sheets[0]);
        Assert.True(doc.Redo()); Assert.Equal("8", doc.GetCellValue(0, 2, 2));
        AssertAttachments(sheet, doc.Snapshot.Sheets[0]);
    }

    [Theory]
    [InlineData("s.getCell('B1').setValue(9)", false)]
    [InlineData("s.getCell('B2').setFormula('=1')", false)]
    [InlineData("s.getRange('A1:B1').setValues([[7,9]])", false)]
    [InlineData("s.getRange('A1:B1').setFormulas([['=7','=9']])", false)]
    [InlineData("s.getCell('D4').setValue(null)", true)]
    [InlineData("s.getRange('C4:D4').setValues([[7,9]])", true)]
    [InlineData("s.getRange('C4:D4').setFormulas([['=7','=9']])", true)]
    public void InvalidTargetsRollBackEarlierEventWritesEvenWhenCaught(string write, bool inCellImage)
    {
        var sheet = SheetWithAttachments(inCellImage);
        using var doc = Document("$.events.on('sheet.activate',()=>{const s=$.getActiveWorksheet();" +
            "s.getCell('C3').setValue(5);s.getCell('A1').getFormat().getFill().setColor('#ff0000');" +
            "try{" + write + ";}catch(e){console.log(String(e));}});", sheet);
        doc.EnableScriptsAndCalculate(); var revision = doc.Snapshot.Revision;
        var error = Assert.Throws<ScriptRuntimeFailureException>(() => doc.DispatchScriptEvents());
        Assert.Contains("invalid write", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(revision, doc.Snapshot.Revision); Assert.False(doc.CanUndo);
        Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.Null(doc.Snapshot.Sheets[0].GetCell(2, 2));
        Assert.Equal(0, doc.Snapshot.Sheets[0].GetStyleId(0, 0));
        AssertAttachments(sheet, doc.Snapshot.Sheets[0]);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyHostPatchAllowsSafeWritesAndPreservesAttachments(bool inCellImage)
    {
        var sheet = SheetWithAttachments(inCellImage);
        using var session = new DocumentSession(Book(sheet).WithScriptModules(M4ActionTests.Actions("")));
        session.SetScriptsEnabled(true); var request = session.BeginAction("A");
        session.CommitAction(request, new(request.OperationId, request.Stamp.InputRevision, "A",
            [new(new(sheet.Id, 0, 0), null, "=7"), new(new(sheet.Id, 2, 2), CellValue.FromNumber(8), null)], []));
        Assert.Equal("=7", session.Snapshot.Sheets[0].GetCell(0, 0)!.Content.FormulaText);
        Assert.Equal(CellValue.FromNumber(8), session.Snapshot.Sheets[0].GetCell(2, 2)!.Value);
        AssertAttachments(sheet, session.Snapshot.Sheets[0]);
        Assert.True(session.Undo()); Assert.False(session.CanUndo);
        Assert.Equal(CellValue.FromNumber(42), session.Snapshot.Sheets[0].GetCell(0, 0)!.Value);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LegacyHostPatchRejectsMergedChildrenAndImageCellsAtomically(bool image, bool formula)
    {
        var sheet = SheetWithAttachments(image);
        using var session = new DocumentSession(Book(sheet).WithScriptModules(M4ActionTests.Actions("")));
        session.SetScriptsEnabled(true); var request = session.BeginAction("A");
        var target = image ? new CellAddress(sheet.Id, 3, 3) : new CellAddress(sheet.Id, 0, 1);
        var before = session.Snapshot;
        var error = Assert.Throws<InvalidOperationException>(() => session.CommitAction(request,
            new(request.OperationId, request.Stamp.InputRevision, "A",
                [new(new(sheet.Id, 2, 2), CellValue.FromNumber(8), null),
                 new(target, formula ? null : CellValue.FromNumber(9), formula ? "=9" : null)], [])));
        Assert.Contains(image ? "单元格内图片" : "非主单元格", error.Message);
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
    }

    [Fact]
    public void FailedMergedWriteRollsBackCellsAndAllVirtualFileVolumes()
    {
        using var doc = Document("""
            $.events.on('sheet.activate', ({fs}) => {
              fs.writeText('//mem/test','memory'); fs.writeText('//temp/test','temporary'); fs.writeText('//xlsq/test','persistent');
              const s=$.getActiveWorksheet(); s.getCell('A1').setValue(7);
              try { s.getCell('B1').setValue(9); } catch {}
            });
            """, SheetWithAttachments());
        doc.EnableScriptsAndCalculate();
        Assert.Throws<ScriptRuntimeFailureException>(() => doc.DispatchScriptEvents());
        Assert.Equal("42", doc.GetCellValue(0, 0, 0)); Assert.False(doc.CanUndo);
        Assert.Empty(doc.Snapshot.PackageFiles.Entries); Assert.Empty(doc.Session.TransientFiles.Entries);
    }

    [Fact]
    public void ConsoleMergedWritePreservesAttachmentsThroughSaveAndReopen()
    {
        var sheet = SheetWithAttachments();
        using var doc = Document("", sheet);
        doc.RunConsole("$.getActiveWorksheet().getCell('A1').setValue('updated');");
        var folder = Path.Combine(Directory.GetCurrentDirectory(), "artifacts", "script-value-writes", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder); var path = Path.Combine(folder, "merged.xlsq");
        doc.Save(path);
        var restored = new TypedWorkbookStorage().Open(path).Sheets[0];
        Assert.Equal(CellValue.FromText("updated"), restored.GetCell(0, 0)!.Value);
        Assert.Equal(sheet.MergedCells.ToArray(), restored.MergedCells.ToArray());
        Assert.True(sheet.Annotations.ContentEquals(restored.Annotations));
        var image = Assert.Single(restored.Images);
        Assert.Equal(sheet.Images[0].Data.ToArray(), image.Data.ToArray()); Assert.Equal(sheet.Images[0].Anchor, image.Anchor);
    }
}
