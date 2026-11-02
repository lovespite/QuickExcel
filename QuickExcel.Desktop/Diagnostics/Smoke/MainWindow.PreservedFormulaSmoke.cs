using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using System.Windows.Controls;
using System.Windows.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunPreservedFormulaSmokeAsync(string folder, Action<bool, string> check)
    {
        var path = Path.Combine(folder, "ui-readonly-array.xlsx");
        var storage = new TypedWorkbookStorage();
        var sheet = new WorksheetSnapshot(SheetId.New(), "来源数组", cells:
            [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(3)))),
             new(new(0, 1), new(CellContent.FromFormula("=A1*2"), new(CellValue.FromNumber(6), CalculationState.Stale, 0, 0, 0))),
             new(new(1, 1), new(CellContent.FromLiteral(CellValue.FromNumber(8))))]);
        storage.Save(path, new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]));
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; XDocument xml;
            using (var stream = entry.Open()) xml = XDocument.Load(stream);
            XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var f = xml.Descendants(main + "f").Single(); f.Value = "A1:A2*2"; f.SetAttributeValue("t", "array"); f.SetAttributeValue("ref", "B1:B2");
            entry.Delete(); using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); xml.Save(output);
        }
        var original = File.ReadAllBytes(path);
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        try
        {
            await OpenPathAsync(path); await Idle();
            check(Title.Contains("只读 · 来源缓存") && FileTitle.Text.Contains("只读"), "Unsupported array workbook has a visible read-only source-cache title");
            check(FormulaBar.IsReadOnly && ScriptEditor.IsReadOnly && !RibbonPasteButton.IsEnabled, "Formula/script editors and paste reflect the read-only policy");
            GridSurface.Select(new(0, 1)); await Idle();
            check(!GridSurface.BeginEdit() && FormulaBar.Text == "=A1:A2*2", "Grid cannot edit an array anchor but the formula bar exposes its original source");
            var rejected = false;
            try { _editor.SetCell(0, 0, 0, "16"); } catch (InvalidOperationException) { rejected = true; }
            check(rejected && !_document.IsDirty && !_document.CanUndo, "Read-only writes are rejected without dirty state or history");
            await RecalculateAsync(); await Idle();
            var cached = _document.Snapshot.Sheets[0].GetCell(0, 1)!;
            check(cached.Value.NumberValue == 6 && cached.Calculation!.State == CalculationState.Blocked && StatusText.Text.Contains("未重算"), "Recalculate retains the source value and identifies it as blocked");
            CaptureGridEvidence(Path.Combine(folder, "readonly-formula.png"));
            _smokeCompatibilityDriver = dialog =>
            {
                try
                {
                    var report = ((CompatibilityDialog)dialog).PlainText;
                    check(report.Contains("只读") && report.Contains("B1:B2") && report.Contains("来源"), "Compatibility dialog identifies read-only formula range and cache provenance");
                    File.WriteAllText(Path.Combine(folder, "readonly-compatibility.txt"), report);
                }
                finally { dialog.Close(); }
            };
            await ShowCompatibilityAsync(); _smokeCompatibilityDriver = null;
            GridSurface.Select(new(4, 4));
            var saved = await SaveAsync(false);
            check(saved && original.SequenceEqual(File.ReadAllBytes(path)), $"Saving after navigation retains the exact original XLSX package (saved={saved}, status={StatusText.Text})");
            var copy = Path.Combine(folder, "ui-readonly-copy.xlsq");
            _smokeSaveDialog = dialog => { dialog.FileName = copy; return true; };
            check(await SaveAsync(true), "Read-only workbook can be saved as an XLSQ container");
            await OpenPathAsync(copy); await Idle();
            check(_document.Snapshot.FileState!.FormulaReadOnly && FormulaBar.IsReadOnly, "XLSQ reopen retains the read-only contract");
            InstallDocument(new WorkbookDocument(new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [new(SheetId.New(), "普通表")]), _services), null, "普通表"); await Idle();
            check(!FormulaBar.IsReadOnly && !ScriptEditor.IsReadOnly && RibbonPasteButton.IsEnabled && GridSurface.BeginEdit(), "Opening an ordinary workbook restores editing controls");
            GridSurface.Cancel();
        }
        finally { _smokeSaveDialog = null; _smokeCompatibilityDriver = null; }
    }
}
