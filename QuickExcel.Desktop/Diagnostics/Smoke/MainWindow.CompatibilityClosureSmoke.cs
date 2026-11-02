using System.IO;
using System.IO.Compression;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Xml.Linq;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop;

public partial class MainWindow
{
    private async Task RunCompatibilityClosureSmokeAsync(string folder, Action<bool, string> check)
    {
        var path = Path.Combine(folder, "shared-style-diagnostics.xlsx"); var storage = new TypedWorkbookStorage();
        var sheet = new WorksheetSnapshot(SheetId.New(), "诊断", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("A")))),
            new(new(1, 1), new(CellContent.FromLiteral(CellValue.FromText("B"))))]);
        storage.Save(path, new(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: [new() { FontName = "Arial", FontSize = 11 }]));
        XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        void ChangePart(string file, string part, Action<XElement> edit)
        {
            using var zip = ZipFile.Open(file, ZipArchiveMode.Update); var entry = zip.GetEntry(part)!; XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            edit(xml.Root!); entry.Delete(); using var output = zip.CreateEntry(part).Open(); xml.Save(output);
        }
        ChangePart(path, "xl/styles.xml", root => root.Element(main + "fonts")!.Elements().First().Add(new XElement(main + "outline")));
        var original = File.ReadAllBytes(path);
        async Task Idle() { await Dispatcher.InvokeAsync(UpdateLayout, DispatcherPriority.ApplicationIdle); await WaitForFastGridPaintAsync(_grid.Control); }
        try
        {
            await OpenPathAsync(path); await Idle(); var snapshot = _document.Snapshot;
            _smokeCompatibilityDriver = dialog =>
            {
                try
                {
                    var text = ((CompatibilityDialog)dialog).PlainText;
                    check(text.Contains("诊断 · A1 · xl/styles.xml") && text.Contains("诊断 · B2 · xl/styles.xml") && text.Contains("font/outline · cellXfs[0]"),
                        "Actual report expands shared style warnings to every affected cell while retaining the source index");
                    File.WriteAllText(Path.Combine(folder, "shared-style-report.txt"), text);
                }
                finally { dialog.Close(); }
            };
            await ShowCompatibilityAsync(); _smokeCompatibilityDriver = null;
            check(ReferenceEquals(snapshot, _document.Snapshot) && !_document.ScriptsEnabled, "Expanded compatibility report does not mutate the document or enable scripts");
            _editor.SetCell(0, 0, 2, "changed"); RefreshSheet();
            _smokeCompatibilityDriver = dialog =>
            {
                try { check(((CompatibilityDialog)dialog).PlainText.Contains("【潜在损失】"), "Editing promotes shared-feature warnings to potential loss in the actual report"); }
                finally { dialog.Close(); }
            };
            await ShowCompatibilityAsync(); _smokeCompatibilityDriver = null;
            _smokeCompatibilitySaveDecision = _ => false;
            check(!await SaveAsync(false) && _document.IsDirty && File.ReadAllBytes(path).SequenceEqual(original), "Declining the explicit compatibility copy keeps the source and unsaved edit intact");
            await UndoAsync(false); await Idle();
            check(TypedWorkbookStorage.InspectCompatibility(_document.Snapshot).OriginalPackageUnchanged, "Undo restores shared-feature original-package preservation");
            var invalid = Path.Combine(folder, "invalid-shared-style.xlsx"); File.Copy(path, invalid);
            ChangePart(invalid, "xl/worksheets/sheet1.xml", root => root.Descendants(main + "c").Single(c => (string?)c.Attribute("r") == "B2").SetAttributeValue("s", 999));
            var beforeRejectedOpen = _document.Snapshot; await OpenPathAsync(invalid); await Idle();
            check(ReferenceEquals(beforeRejectedOpen, _document.Snapshot) && LogText.Text.Contains("【拒绝】诊断 · B2 · xl/worksheets/sheet1.xml"),
                "Actual rejected-open handler shows the exact source cell and retains the previous document");
            File.WriteAllText(Path.Combine(folder, "located-rejection-log.txt"), LogText.Text);
        }
        finally { _smokeCompatibilityDriver = null; _smokeCompatibilitySaveDecision = null; }
    }
}
