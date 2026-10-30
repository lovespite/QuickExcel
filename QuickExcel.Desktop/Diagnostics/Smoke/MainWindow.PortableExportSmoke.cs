using System.IO;
using System.Text.Json;
using System.Windows;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Formulas;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunPortableExportSmokeAsync(string folder, Action<bool, string> check)
    {
        CellSnapshot Formula(string text) => new(CellContent.FromFormula(text));
        CellSnapshot Number(double value) => new(CellContent.FromLiteral(CellValue.FromNumber(value)));
        var data = new WorksheetSnapshot(SheetId.New(), "O'Brien 数据", [
            new(new(0, 0), Number(3)), new(new(1, 0), Number(4)),
            new(new(0, 2), Formula("=JS.DOUBLE(A1)")), new(new(1, 2), Formula("=IF(FALSE,JS.FAIL(),9)")),
            new(new(2, 2), Formula("=SUM(IF(TRUE,JS.NEXT(),JS.FAIL()))"))], rows: [new(0, new(Height: 28.333333))]);
        var report = new WorksheetSnapshot(SheetId.New(), "Report", [
            new(new(0, 0), Formula("='O''Brien 数据'!C1+1")), new(new(1, 0), Formula("=SUM('O''Brien 数据'!C1:C3)")),
            new(new(2, 0), Formula("=NamedScript+1")), new(new(3, 0), Formula("=IF(TRUE,8,'O''Brien 数据'!C1)")),
            new(new(0, 1), Formula("=SUM(Payload)")), new(new(1, 1), Formula("=Rate+Local")),
            new(new(2, 1), Formula("=IFERROR(1/0,9)")), new(new(3, 1), Formula("=COUNTBLANK('O''Brien 数据'!A:A)")),
            new(new(4, 1), Formula("=CONCAT(\"😀\",42)")), new(new(5, 1), Formula("=XLOOKUP(4,Payload,Payload,,0)"))]);
        var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [data, report], textCompatibilityVersion: 2)
            .WithDefinedNames([new("Payload", "='O''Brien 数据'!$A$1:$A$2"), new("Rate", "=2"), new("Local", "=5", report.Id), new("NamedScript", "=JS.DOUBLE(5)")])
            .WithScriptSource("""
                let n=0;
                quickExcel.functions.register({name:'JS.DOUBLE',apiVersion:2,parameters:[{name:'x',type:'number'}],returns:'number'},x=>x*2);
                quickExcel.functions.register({name:'JS.NEXT',apiVersion:2,parameters:[],returns:'number'},()=>++n);
                quickExcel.functions.register({name:'JS.FAIL',apiVersion:2,parameters:[],returns:'number'},()=>{throw new Error('unselected')});
                """);
        var source = Path.Combine(folder, "portable-source.xlsq"); var output = Path.Combine(folder, "portable-export.xlsx");
        InstallDocument(new WorkbookDocument(book, _services), source, "Portable export smoke.xlsq");
        try
        {
            if (!await WorkAsync("验收可移植导出重算…", _document.EnableScriptsAndCalculate)) throw new Exception(LogText.Text);
            check(true, "Portable export calculates through the desktop worker services");
            var expected = new (string Sheet, string Address, CellValue Value)[] {
                (data.Name, "A1", CellValue.FromNumber(3)), (data.Name, "A2", CellValue.FromNumber(4)),
                (data.Name, "C1", CellValue.FromNumber(6)), (data.Name, "C2", CellValue.FromNumber(9)), (data.Name, "C3", CellValue.FromNumber(1)),
                ("Report", "A1", CellValue.FromNumber(7)), ("Report", "A2", CellValue.FromNumber(16)),
                ("Report", "A3", CellValue.FromNumber(11)), ("Report", "A4", CellValue.FromNumber(8)),
                ("Report", "B1", CellValue.FromNumber(7)), ("Report", "B2", CellValue.FromNumber(7)), ("Report", "B3", CellValue.FromNumber(9)),
                ("Report", "B4", CellValue.FromNumber(1048574)), ("Report", "B5", CellValue.FromText("😀42")), ("Report", "B6", CellValue.FromNumber(4)) };
            CellSnapshot Cell(WorkbookSnapshot snapshot, string sheetName, string address)
            { var reference = (ReferenceNode)FormulaSyntax.Parse("=" + address); return snapshot.Sheets.Single(s => s.Name == sheetName).GetCell(reference.Row, reference.Column)!; }
            check(expected.All(c => Cell(_document.Snapshot, c.Sheet, c.Address).Value == c.Value), "Worker results, lazy branches, names and independent formulas match explicit acceptance values");
            _document.Save(source); var bytes = File.ReadAllBytes(source);
            _document.SetCell(0, 5, 3, "unsaved source edit");
            check(await WorkAsync("验收导出重算…", _document.EnableScriptsAndCalculate), "Source edit recalculates before portable export");
            check(_document.InspectPortableExport().StaleResults == 0 && _viewModel.StaleCacheCount == 0, "Portable export uses fresh worker results after the source edit");
            var before = _document.Snapshot; var undo = _document.CanUndo;
            _smokeSaveDialog = dialog => { dialog.FileName = output; return true; };
            _smokePortableExportDecision = message => { check(message.Contains("保留 6") && message.Contains("7 个脚本"), "Export confirmation reports retained and frozen dependency counts"); return MessageBoxResult.No; };
            ExportPortable_Click(this, new()); await _portableExportTask;
            check(!File.Exists(output) && ReferenceEquals(before, _document.Snapshot), "Cancelling portable export writes nothing and preserves the source");
            _smokePortableExportDecision = _ => MessageBoxResult.Yes;
            ExportPortable_Click(this, new()); await _portableExportTask;
            check(File.Exists(output), "Actual portable export handler writes the chosen XLSX copy");
            check(ReferenceEquals(before, _document.Snapshot) && _document.IsDirty && _document.CanUndo == undo && _document.ScriptsEnabled && bytes.SequenceEqual(File.ReadAllBytes(source)),
                "Portable export preserves unsaved edits, history, script authorization and the source file bytes");
            var saved = new TypedWorkbookStorage().Open(output);
            check(!saved.FileState!.FormulaReadOnly && saved.DefinedNames.Length == 0 && !TypedWorkbookStorage.RequiresScriptContainer(saved), "Portable copy reopens editable without names or script modules");
            check(expected.All(c => Cell(saved, c.Sheet, c.Address).Value == c.Value), "Portable XLSX caches and frozen values match explicit acceptance values");
            check(saved.Sheets.Sum(s => s.Cells.Values.Count(c => c.Content.Kind == CellContentKind.Formula)) == 6, "Only six independent supported formulas remain in the exported copy");
            var inputs = expected.Where(c => Cell(saved, c.Sheet, c.Address).Content.Kind == CellContentKind.Literal).Select(c => new {
                c.Sheet, c.Address, Kind = c.Value.Kind.ToString(), Value = c.Value.Kind == CellValueKind.Text ? (object)c.Value.TextValue : c.Value.NumberValue });
            var cases = expected.Select((c, i) => new { Id = "portable-" + (i + 1), c.Sheet, c.Address,
                Formula = Cell(saved, c.Sheet, c.Address).Content.Kind == CellContentKind.Formula ? Cell(saved, c.Sheet, c.Address).Content.FormulaText : null });
            File.WriteAllText(Path.Combine(folder, "portable-export-spec.json"), JsonSerializer.Serialize(new {
                Sheets = new[] { data.Name, report.Name }, Date1904 = false, TextCompatibilityVersion = 2, Inputs = inputs, Cases = cases, Edits = Array.Empty<object>(), GuardedCases = Array.Empty<object>() }, new JsonSerializerOptions { WriteIndented = true }));
            RefreshSheet(); GridSurface.Select(new(0, 0));
            await Dispatcher.InvokeAsync(UpdateLayout, System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            await WaitForFastGridPaintAsync(_grid.Control);
            CaptureGridEvidence(Path.Combine(folder, "portable-export-source.png"));
        }
        finally { _smokePortableExportDecision = null; _smokeSaveDialog = null; }
    }
}
