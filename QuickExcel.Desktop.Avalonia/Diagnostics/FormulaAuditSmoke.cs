using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
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
            _ribbon.SelectedItem = _ribbon.Items.Cast<TabItem>().Single(t => Equals(t.Header, UiText.Get("审阅")));
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            var auditButton = _ribbon.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "FormulaAuditRibbon");
            Check(auditButton.IsEffectivelyEnabled && AutomationProperties.GetName(auditButton) == UiText.Get("公式审核") &&
                auditButton.GetVisualDescendants().OfType<PathIcon>().Any(i => i.Data is not null && i.Bounds.Width > 0),
                "Review ribbon exposes an accessible formula audit command with a visible vector icon");
            using (var bitmap = new RenderTargetBitmap(new PixelSize((int)Bounds.Width, (int)Bounds.Height), new Vector(96, 96)))
            { bitmap.Render(this); bitmap.Save(Path.Combine(directory, "formula-audit-ribbon.png"), PngBitmapEncoderOptions.Default); }
            auditButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            var deadline = DateTime.UtcNow.AddSeconds(5);
            while (_formulaAudit?.Result is null && DateTime.UtcNow < deadline) await Task.Delay(20);
            Check(_formulaAudit?.Result is not null, "Clicking the Review command opens and populates formula audit");
            var result = _formulaAudit!.Result!;
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            Check(_formulaAudit.GetVisualDescendants().OfType<Button>().Any(b => Equals(b.Content, UiText.Language == UiText.English ? "Refresh" : "刷新")),
                "Formula audit refresh button follows the chosen language");
            using (var bitmap = new RenderTargetBitmap(new PixelSize((int)_formulaAudit.Bounds.Width, (int)_formulaAudit.Bounds.Height), new Vector(96, 96)))
            { bitmap.Render(_formulaAudit); bitmap.Save(Path.Combine(directory, "formula-audit-window.png"), PngBitmapEncoderOptions.Default); }
            Check(result.Precedents.Count == 1 && result.Dependents.Count == 2 && result.Formula == "=A1+1", "Audit shows direct references and same/cross-sheet consumers");
            Check(_grid.FormulaAuditArrowCount == 1, "Audit connects precedent and origin with a grid arrow");
            _formulaAudit.NextStep();
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
            Check(_grid.FormulaAuditArrowCount == 0, "Closing audit clears grid arrows");
            await ExecuteAsync(new SetCellsCommand([new(new(add.Id, 3, 0), CellContent.FromFormula("=1/0"))])); await CalculatePendingAsync();
            await ShowFormulaAuditAsync();
            var errors = QuickExcel.Calculation.FormulaEvaluation.Errors(Document.Snapshot);
            await _formulaAudit!.NavigateAsync(errors.Single());
            Check(_grid.Current == new CellPosition(3, 0), "Current error list navigates to the failing cell");
        }
        catch (Exception ex) { failure = ex; }
        finally { _formulaAudit?.Close(); }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
