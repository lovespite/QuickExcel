using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunFormulaAuditSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        try
        {
            _workspace.New(); ShowBook(); var first = Document.Snapshot.ActiveSheetId; var name = Document.Snapshot.GetSheet(first).Name;
            _workspace.Editor.SetCells(0, [(0, 0, "2"), (0, 1, "=A1+1"), (0, 2, "=B1*2")]);
            var add = new AddSheetCommand("Second"); await ExecuteAsync(add);
            await ExecuteAsync(new SetCellsCommand([new(new(add.Id, 0, 0), CellContent.FromFormula("='" + name.Replace("'", "''") + "'!B1+1"))]));
            await CalculatePendingAsync(); Synchronize(); _grid.Select(new(0, 1));
            Check(((MenuItem)_menu.Items[3]!).Items.OfType<MenuItem>().Any(m => m.Name == "FormulaAuditMenu" && m.IsEnabled), "View menu exposes formula audit");
            await ShowFormulaAuditAsync(); var result = _formulaAudit!.Result!;
            Check(result.Precedents.Count == 1 && result.Dependents.Count == 2 && result.Formula == "=A1+1", "Audit shows direct references and same/cross-sheet consumers");
            var revision = Document.Snapshot.Revision; await _formulaAudit.NavigateAsync(result.Dependents.Single(t => t.Address.SheetId == add.Id));
            Check(Document.Snapshot.ActiveSheetId == add.Id && _grid.Current == new CellPosition(0, 0) && Document.Snapshot.Revision == revision, "Cross-sheet audit navigation does not dirty the workbook");
            await _formulaAudit.RefreshAsync(false, true);
            await ExecuteAsync(new SetCellsCommand([new(new(first, 0, 1), CellContent.FromFormula("=A1+5"))])); await CalculatePendingAsync();
            await _formulaAudit.RefreshAsync(false, true);
            Check(_formulaAudit.Result!.Formula == "=A1+5", "Audit refresh reads the updated formula at its fixed address");
            var old = _formulaAudit.Result.Precedents.Single();
            await ExecuteAsync(new SetCellsCommand([new(new(first, 0, 0), CellContent.FromLiteral(CellValue.FromNumber(9)))]));
            var rejected = false; try { await _formulaAudit.NavigateAsync(old); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected && Document.Snapshot.ActiveSheetId == add.Id, "Stale audit navigation requires selecting a refreshed result");
            _formulaAudit.Close(); Check(_formulaAudit is null, "Closing audit releases its refresh lifecycle");
        }
        catch (Exception ex) { failure = ex; }
        finally { _formulaAudit?.Close(); }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
