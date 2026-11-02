using System.Globalization;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public sealed class FormatEfficiencyTests
{
    [Fact]
    public void DecimalPlacesMatchIndependentExcelRibbonCommands()
    {
        using var spec = ScalarFormulaFixture.Load("format-precision-spec"); using var oracle = ScalarFormulaFixture.Load("excel-format-precision");
        var expected = oracle.RootElement.GetProperty("Cases").EnumerateArray().ToDictionary(c => c.GetProperty("Id").GetString()!);
        var failures = new List<string>();
        foreach (var item in spec.RootElement.GetProperty("Cases").EnumerateArray())
        {
            var code = item.GetProperty("Format").GetString()!;
            var style = code == "General" ? new StyleDefinition() : new StyleDefinition { NumberFormatCode = code };
            var value = CellValue.FromNumber(item.GetProperty("Value").GetDouble());
            foreach (var step in item.GetProperty("Steps").EnumerateArray())
            {
                if (!CellDisplayFormatter.TryAdjustDecimalPlaces(style, value, step.GetInt32(), out var adjusted)) { failures.Add(item.GetProperty("Id") + " rejected " + code); break; }
                style = adjusted;
            }
            var actual = CellDisplayFormatter.Format(value, style, ExcelDateSystem.Windows1900, CultureInfo.GetCultureInfo("zh-CN"));
            var wanted = expected[item.GetProperty("Id").GetString()!].GetProperty("Text").GetString()!;
            // Accounting '_'/'?' are layout spaces; their pixel widths depend on the host font.
            if (actual.Replace(" ", "") != wanted.Replace(" ", "")) failures.Add($"{item.GetProperty("Id")}: {code} {value.NumberValue} => {style.NumberFormatCode}; expected [{wanted}], actual [{actual}]");
        }
        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }

    [Fact]
    public void ColorCommandsCanClearFillOrFontWithoutResettingOtherStyleFields()
    {
        var sheet = Sheet("S", (0, 0, N(1)));
        using var session = new DocumentSession(Book(sheet));
        var address = new CellAddress(sheet.Id, 0, 0);
        session.Execute(new SetCellStyleCommand([address], new CellStylePatch(FillColor: "FFFF0000", FontColor: "FF0070C0", Bold: true)));
        session.Execute(new SetCellStyleCommand([address], new CellStylePatch(ClearFillColor: true)));
        var style = session.Snapshot.Styles[session.Snapshot.Sheets[0].GetStyleId(0, 0)];
        Assert.Null(style.FillColor); Assert.Equal("FF0070C0", style.FontColor); Assert.True(style.Bold);
        session.Execute(new SetCellStyleCommand([address], new CellStylePatch(ClearFontColor: true)));
        style = session.Snapshot.Styles[session.Snapshot.Sheets[0].GetStyleId(0, 0)];
        Assert.Null(style.FontColor); Assert.True(style.Bold);
        session.Undo(); Assert.Equal("FF0070C0", session.Snapshot.Styles[session.Snapshot.Sheets[0].GetStyleId(0, 0)].FontColor);
    }

    [Fact]
    public void ClearingFontColorAlsoRemovesExplicitRichTextRunColors()
    {
        var runs = new RichCellText([new TextRun("A", new TextRunFont(Color: "FFFF0000")), new TextRun("B")]);
        var sheet = new WorksheetSnapshot(SheetId.New(), "Rich", [new(new(0, 0), new(CellContent.FromRichText(runs, 1)))]);
        using var session = new DocumentSession(new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet],
            styles: [new(), new() { Bold = true, FontColor = "FF00FF00" }]));
        session.Execute(new SetCellStyleCommand([new(sheet.Id, 0, 0)], new CellStylePatch(ClearFontColor: true)));
        var cell = session.Snapshot.Sheets[0].GetCell(0, 0)!;
        Assert.Equal("AB", cell.Value.TextValue); Assert.Null(cell.Content.RichText!.Runs[0].Font!.Color);
        Assert.Null(cell.Content.RichText.Runs[0].Font!.RawRgb);
    }

    [Fact]
    public void PrecisionChangesAreAtomicUndoableAndKeepFreshFormulaValues()
    {
        var sheet = Sheet("S", (0, 0, N(1.234)), (1, 0, F("=A1*2")));
        using var session = new DocumentSession(Book(sheet));
        Assert.True(session.TryCommitCalculation(new WorkbookCalculator(new NoScripts()).Calculate(session.BeginCalculation())));
        var before = session.Snapshot;
        var result = session.Execute(new AdjustDecimalPlacesCommand([new(sheet.Id, 0, 0), new(sheet.Id, 1, 0)], 1));
        Assert.True(result.Changed); Assert.True(result.CalculationUnchanged);
        Assert.Equal(CellValue.FromNumber(2.468), session.Snapshot.Sheets[0].GetCell(1, 0)!.Value);
        Assert.Equal(CalculationState.Fresh, session.Snapshot.Sheets[0].GetCell(1, 0)!.Calculation!.State);
        Assert.Equal("0.0000", session.Snapshot.Styles[session.Snapshot.Sheets[0].GetStyleId(1, 0)].NumberFormatCode);
        Assert.True(session.Undo()); Assert.Equal(before.Sheets[0].GetStyleId(0, 0), session.Snapshot.Sheets[0].GetStyleId(0, 0));
        Assert.True(session.Redo()); Assert.False(session.CanRedo);
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            var storage = new TypedWorkbookStorage(); storage.Save(path, session.Snapshot); var reopened = storage.Open(path);
            Assert.Equal("=A1*2", reopened.Sheets[0].GetCell(1, 0)!.Content.FormulaText);
            Assert.Equal(CellValue.FromNumber(2.468), reopened.Sheets[0].GetCell(1, 0)!.Value);
            Assert.Equal("0.0000", reopened.Styles[reopened.Sheets[0].GetStyleId(1, 0)].NumberFormatCode);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void UnsupportedDateOrFractionDoesNotPartlyModifyMixedSelection()
    {
        var sheet = Sheet("S", (0, 0, N(2.3)), (1, 0, new(CellContent.FromLiteral(CellValue.FromNumber(46000), 1))));
        using var session = new DocumentSession(new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new(), new() { NumberFormatCode = "yyyy-mm-dd" }]));
        var before = session.Snapshot;
        Assert.Throws<InvalidOperationException>(() => session.Execute(new AdjustDecimalPlacesCommand([new(sheet.Id, 0, 0), new(sheet.Id, 1, 0)], 1)));
        Assert.Same(before, session.Snapshot); Assert.False(session.CanUndo);
        foreach (var code in new[] { "# ?/?", "@", "[DBNum1]0", "[h]:mm:ss", "0\"x\"0" })
            Assert.False(CellDisplayFormatter.TryAdjustDecimalPlaces(new() { NumberFormatCode = code }, CellValue.FromNumber(1), 1, out _));
    }
}
