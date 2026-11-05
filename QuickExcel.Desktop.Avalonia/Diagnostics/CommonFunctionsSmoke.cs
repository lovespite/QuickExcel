using System.Text.Json;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunCommonFunctionsSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        try
        {
            _workspace.New(); ShowBook(); _grid.Select(new(5, 3)); _grid.BeginEdit();
            _grid.CellEditor.Text = "=ROW()+COLUMN()+ROWS(A1:B3)+COLUMNS(A1:B3)";
            Check(CommitEditors(), "Native grid commits reference functions");
            await CalculateAsync(); Check(Document.GetCellValue(0, 5, 3) == "15", "Reference functions calculate through the normal desktop command");
            Check(FormulaEditing.Complete("=COL", 4)!.Names.Contains("COLUMNS"), "Formula completion includes the new functions");
            _grid.Select(new(9, 0)); _grid.BeginEdit(); _grid.CellEditor.Text = "=ROW(B2:B4)"; Check(CommitEditors(), "Native editor accepts a reference array");
            await CalculateAsync(); Check(Document.GetCellValue(0, 11, 0) == "4", "Reference result spills into the grid");
            await History(false); Check(Document.GetCellInput(0, 9, 0) == "", "One undo removes the formula and its spill");
            _grid.Select(new(15, 0)); _grid.BeginEdit(); _grid.CellEditor.Text = "=IFS(TRUE,SUMPRODUCT(SEQUENCE(3),SEQUENCE(3)),FALSE,0)+SWITCH(2,1,1,2,5,0)";
            Check(CommitEditors(), "Native grid commits logical and product formulas");
            await CalculateAsync(); Check(Document.GetCellValue(0, 15, 0) == "19", "IFS, SWITCH and SUMPRODUCT calculate in the desktop");
            _grid.Select(new(20, 0)); _grid.BeginEdit(); _grid.CellEditor.Text = "=TEXTSPLIT(\"甲,乙;丙\",\",\",\";\")";
            Check(CommitEditors(), "Native grid accepts TEXTSPLIT"); await CalculateAsync();
            Check(Document.GetCellValue(0, 20, 1) == "乙" && Document.GetCellValue(0, 21, 0) == "丙", "Text arrays spill in both directions");
            _grid.Select(new(25, 0)); _grid.BeginEdit(); _grid.CellEditor.Text = "=TAKE(DROP(SORTBY(SEQUENCE(4),SEQUENCE(4),-1),1),2)";
            Check(CommitEditors(), "Native grid accepts composed array slicing and sorting"); await CalculateAsync();
            Check(Document.GetCellValue(0, 25, 0) == "3" && Document.GetCellValue(0, 26, 0) == "2", "Composed arrays render in the grid");
        }
        catch (Exception ex) { failure = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
