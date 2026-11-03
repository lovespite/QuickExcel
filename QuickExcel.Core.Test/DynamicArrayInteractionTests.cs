using QuickExcel.Calculation;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Runtime;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class DynamicArrayInteractionTests
{
    private static void Calculate(DocumentSession session) => Assert.True(session.TryCommitCalculation(
        new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));

    [Theory]
    [InlineData(CellPasteMode.All)]
    [InlineData(CellPasteMode.Formulas)]
    public void WholeArrayCopyKeepsChildrenDerivedAndPartialCopyCopiesTheirDisplayedValues(CellPasteMode mode)
    {
        var sheet = Sheet("Data",(0,0,F("=SEQUENCE(3,2)")));
        using var session = new DocumentSession(Book(sheet)); Calculate(session);
        var clip = WorkbookClipboard.Capture(session.Snapshot,sheet.Id,new(0,0,2,1));
        session.Execute(new PasteCellsCommand(clip,sheet.Id,new(4,3,4,3),session.Revision,mode)); Calculate(session);
        Assert.Equal(new CellAddress(sheet.Id,4,3),session.Snapshot.Sheets[0].GetCell(6,4)!.SpillOwner);
        Assert.Equal(CellValue.FromNumber(6),session.Snapshot.Sheets[0].GetCell(6,4)!.Value);
        var child = WorkbookClipboard.Capture(session.Snapshot,sheet.Id,new(1,0,2,1));
        session.Execute(new PasteCellsCommand(child,sheet.Id,new(10,0,10,0),session.Revision));
        Assert.Null(session.Snapshot.Sheets[0].GetCell(11,1)!.SpillOwner);
        Assert.Equal(CellValue.FromNumber(6),session.Snapshot.Sheets[0].GetCell(11,1)!.Value);
        Assert.Throws<InvalidOperationException>(() => WorkbookClipboard.Capture(session.Snapshot,sheet.Id,new(1,0,2,1),cut:true));
    }

    [Fact]
    public void CutShiftCopySheetAndUndoRebuildSpillOwnership()
    {
        var sheet = Sheet("Data",(1,1,F("=SEQUENCE(3)")),(0,5,F("=SUM(B2#)")));
        using var session = new DocumentSession(Book(sheet)); Calculate(session);
        var clip = WorkbookClipboard.Capture(session.Snapshot,sheet.Id,new(1,1,3,1),cut:true);
        session.Execute(new PasteCellsCommand(clip,sheet.Id,new(6,2,6,2),session.Revision)); Calculate(session);
        Assert.Null(session.Snapshot.Sheets[0].GetCell(2,1));
        Assert.Equal(new CellAddress(sheet.Id,6,2),session.Snapshot.Sheets[0].GetCell(8,2)!.SpillOwner);
        Assert.Equal(CellValue.FromNumber(6),session.Snapshot.Sheets[0].GetCell(0,5)!.Value);
        Assert.True(session.Undo()); Calculate(session);
        session.Execute(new ShiftAxisCommand(sheet.Id,SheetAxis.Row,0,2)); Calculate(session);
        Assert.Equal(new CellAddress(sheet.Id,3,1),session.Snapshot.Sheets[0].GetCell(5,1)!.SpillOwner);
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new ShiftAxisCommand(sheet.Id,SheetAxis.Row,4,1,delete:true)));
        Assert.Same(before,session.Snapshot);
        var copy = new CopySheetCommand(sheet.Id,"Copy",1); session.Execute(copy); Calculate(session);
        Assert.Equal(new CellAddress(copy.Id,3,1),session.Snapshot.GetSheet(copy.Id).GetCell(5,1)!.SpillOwner);
        Assert.Equal("=SEQUENCE(3)",CellInput.EditText(session.Snapshot,session.Snapshot.GetSheet(copy.Id),5,1));
    }

    [Fact]
    public void ExplicitArrayContextSurvivesCopyAndDefinedArrayIsPromotedForStorage()
    {
        var sheet = Sheet("Data",(0,0,N(2)),(1,0,N(3)),(0,2,new CellSnapshot(CellContent.FromFormula("=A1:A2",dynamicArray:true))),
            (4,0,F("=Numbers")));
        using var session = new DocumentSession(Book(sheet).WithDefinedNames([new DefinedName("Numbers","=SEQUENCE(2)")])); Calculate(session);
        Assert.True(session.Snapshot.Sheets[0].GetCell(4,0)!.Content.DynamicArray);
        var clip = WorkbookClipboard.Capture(session.Snapshot,sheet.Id,new(0,2,1,2));
        session.Execute(new PasteCellsCommand(clip,sheet.Id,new(0,4,0,4),session.Revision)); Calculate(session);
        Assert.True(session.Snapshot.Sheets[0].GetCell(0,4)!.Content.DynamicArray);
        Assert.Equal(CellValue.FromNumber(3),session.Snapshot.Sheets[0].GetCell(1,4)!.Value);
    }

    [Fact]
    public void UnavailableArrayProducerRetainsCachesAndBlocksUnknownFootprintConsumers()
    {
        var sheet = Sheet("Data",(0,0,F("=SEQUENCE(3)")),(0,3,F("=A3+1")));
        using var session = new DocumentSession(Book(sheet)); Calculate(session);
        session.Execute(new SetCellsCommand([new(new(sheet.Id,0,0),CellContent.FromFormula("=SEQUENCE(MISSING())"))]));
        Calculate(session);
        Assert.Equal(CalculationState.Blocked,session.Snapshot.Sheets[0].GetCell(0,3)!.Calculation!.State);
        Assert.Equal(CellValue.FromNumber(4),session.Snapshot.Sheets[0].GetCell(0,3)!.Value);
        Assert.Equal(CalculationState.Blocked,session.Snapshot.Sheets[0].GetCell(2,0)!.Calculation!.State);
        Assert.Equal(CellValue.FromNumber(3),session.Snapshot.Sheets[0].GetCell(2,0)!.Value);
    }

    [Fact]
    public void ChildFormattingPreservesFreshOwnershipAndRemainsAfterSpillShrinks()
    {
        var sheet = Sheet("Data",(0,0,F("=SEQUENCE(3)")));
        using var session = new DocumentSession(Book(sheet)); Calculate(session);
        var child = new CellAddress(sheet.Id,2,0);
        session.Execute(new SetCellStyleCommand([child],new(Bold:true)));
        session.Execute(new SetCellStyleCommand([child],new(Bold:true)));
        var styled = session.Snapshot.Sheets[0].GetCell(2,0)!;
        Assert.Equal(new CellAddress(sheet.Id,0,0),styled.SpillOwner);
        Assert.Equal(CalculationState.Fresh,styled.Calculation!.State);
        Assert.Equal(session.Revision,styled.Calculation.InputRevision);
        session.Execute(new SetCellsCommand([new(new(sheet.Id,0,0),CellContent.FromFormula("=SEQUENCE(2)"))])); Calculate(session);
        var remaining = session.Snapshot.Sheets[0].GetCell(2,0)!;
        Assert.Equal(CellValue.Blank,remaining.Value); Assert.Null(remaining.SpillOwner);
        Assert.True(session.Snapshot.Styles[remaining.Content.StyleId].Bold);
    }

    [Fact]
    public void TableAndSheetEdgePreventWholeArrayPublication()
    {
        var sheet = Sheet("Data",(0,0,F("=SEQUENCE(3)")),(WorkbookLimits.MaxRows-1,2,F("=SEQUENCE(2)")));
        sheet = sheet.WithData(sheet.Data with { Tables = [new WorksheetTable(1,"Source",new(1,0,2,0),[new TableColumn(1,"Column")])] });
        using var session = new DocumentSession(Book(sheet)); Calculate(session);
        Assert.Equal(CellValue.FromError(CellErrorCode.Spill),session.Snapshot.Sheets[0].GetCell(0,0)!.Value);
        Assert.Equal(CellValue.FromError(CellErrorCode.Spill),session.Snapshot.Sheets[0].GetCell(WorkbookLimits.MaxRows-1,2)!.Value);
        Assert.DoesNotContain(session.Snapshot.Sheets[0].Cells.Values,c => c.SpillOwner is not null);
    }

    [Fact]
    public void StatefulScriptsRunOnceInFinalTopologyAfterPureSpillPlanning()
    {
        var sheet = Sheet("Data",(0,0,F("=NEXT(B4)")),(1,0,F("=NEXT()")),(2,1,F("=SEQUENCE(2)")),(5,0,F("=IF(FALSE,NEXT(),NEXT())")));
        var book = Book(sheet).WithScriptSource("let n=0;quickExcel.functions.register('NEXT',()=>++n);");
        var request = Request(book) with { ScriptsEnabled = true };
        var result = new WorkbookCalculator(new JscRuntimeFactory()).Calculate(request);
        Assert.Equal(CellValue.FromNumber(2),C(result,sheet,0).Value);
        Assert.Equal(CellValue.FromNumber(1),C(result,sheet,1).Value);
        Assert.Equal(CellValue.FromNumber(3),C(result,sheet,5).Value);
    }
}
