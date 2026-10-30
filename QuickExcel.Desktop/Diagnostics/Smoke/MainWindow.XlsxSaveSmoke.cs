using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using System.Windows;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunXlsxSaveSmokeAsync(string folder, Action<bool, string> check)
    {
        var storage = new TypedWorkbookStorage();
        var path = Path.Combine(folder, "ui-source.xlsx");
        var snapshot = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(SheetId.New(), "保存对照", cells: [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(2)))),
                new(new(0, 1), new(CellContent.FromFormula("=SUM(A1,3)"), new(CellValue.FromNumber(5), CalculationState.Stale, 0, 0, 0)))])]);
        storage.Save(path, snapshot);
        try
        {
            await OpenPathAsync(path);
            check(_filePath == path && _fileVersion == FileVersions.Read(path), "Opening XLSX retains its save destination and version");
            _editor.SetCell(0, 0, 0, "7"); RefreshSheet();
            check(await SaveAsync(false) && !_document.IsDirty, "Ctrl+S handler saves an edited XLSX in place");
            check(storage.Open(path).Sheets[0].GetCell(0, 1)!.Content.FormulaText == "=SUM(A1,3)", "XLSX UI save retains formula source");
            var copy = Path.Combine(folder, "ui-copy.xlsx");
            _smokeSaveDialog = dialog =>
            {
                check(dialog.Filter.Contains("保留公式") && dialog.DefaultExt.TrimStart('.') == "xlsx", "Save As offers formula-preserving XLSX");
                dialog.FileName = copy; return true;
            };
            check(await SaveAsync(true) && _filePath == copy && File.Exists(copy), "Save As updates the active XLSX destination");
            _editor.SetCell(0, 0, 0, "9"); RefreshSheet();
            File.WriteAllText(copy, "external modification");
            check(!await SaveAsync(false) && _document.IsDirty && File.ReadAllText(copy) == "external modification", "XLSX external conflict keeps the local draft and external file");
            _smokeSaveDialog = _ => false;
            check(!await SaveAsync(true) && _document.IsDirty, "Cancel XLSX Save As keeps dirty input");
            await OpenPathAsync(path);
            var sourceBytes = File.ReadAllBytes(path);
            _document.ScriptSource = "// user script must be saved"; LoadScriptModules();
            var container = Path.Combine(folder, "ui-script-main.xlsq");
            _smokeSaveDialog = dialog =>
            {
                check(!dialog.Filter.Contains("*.xlsx") && dialog.Title.Contains("脚本主文件"), "Adding script routes Save to an explicitly labelled XLSQ main file");
                dialog.FileName = container; return true;
            };
            check(await SaveAsync(false) && _filePath == container && !_document.IsDirty, "Scripted XLSX is saved as an XLSQ main file");
            check(sourceBytes.SequenceEqual(File.ReadAllBytes(path)) && storage.Open(container).ScriptSource.Contains("user script"), "XLSQ save keeps scripts and leaves the imported XLSX untouched");
            _document.SetCell(0, 0, 0, "12");
            _document.ExportXlsx(Path.Combine(folder, "ui-results-copy.xlsx"), allowStaleCaches: true);
            check(_document.IsDirty && _filePath == container, "XLSX result export does not mark the scripted main file saved");

            var protectedPath = Path.Combine(folder, "ui-unknown.xlsx"); storage.Save(protectedPath, snapshot);
            using (var zip = ZipFile.Open(protectedPath, ZipArchiveMode.Update))
            {
                var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!; XDocument xml;
                using (var input = entry.Open()) xml = XDocument.Load(input);
                xml.Root!.Add(new XElement(xml.Root.Name.Namespace + "extLst", new XElement(xml.Root.Name.Namespace + "ext", new XAttribute("uri", "urn:unknown"))));
                entry.Delete(); using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open(); xml.Save(output);
            }
            var protectedBytes = File.ReadAllBytes(protectedPath); await OpenPathAsync(protectedPath);
            _editor.SetCell(0, 0, 0, "17"); RefreshSheet();
            _smokeCompatibilitySaveDecision = MessageBoxResult.No;
            check(!await SaveAsync(false) && _document.IsDirty && protectedBytes.SequenceEqual(File.ReadAllBytes(protectedPath)), "Declining a compatibility copy keeps input and original XLSX");
            var supported = Path.Combine(folder, "ui-supported-copy.xlsx");
            _smokeCompatibilitySaveDecision = MessageBoxResult.Yes;
            _smokeSaveDialog = dialog => { dialog.FileName = supported; return true; };
            check(await SaveAsync(false) && _filePath == supported && protectedBytes.SequenceEqual(File.ReadAllBytes(protectedPath)), "Explicit compatibility Save writes a separate XLSX and preserves the original");
            check(storage.Open(supported).FileState!.CompatibilityIssues.Length == 0 && storage.Open(supported).Sheets[0].GetCell(0, 0)!.Value.NumberValue == 17, "Compatible XLSX copy contains the edited value and omits only reported unsupported content");
        }
        finally { _smokeSaveDialog = null; _smokeCompatibilitySaveDecision = null; }
    }
}
