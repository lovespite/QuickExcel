using System.Diagnostics;
using System.Text.Json;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.CompatibilityProbe;

public static class E4Probe
{
    public static int Run(string source, string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new output directory."); Directory.CreateDirectory(directory);
        var storage = new TypedWorkbookStorage(); var book = storage.Open(source); var sourceAudit = OpenXmlAudit.Inspect(File.ReadAllBytes(source));
        File.WriteAllText(Path.Combine(directory, "source-audit.json"), JsonSerializer.Serialize(sourceAudit, ProbeRunner.JsonOptions));
        File.WriteAllText(Path.Combine(directory, "import.json"), JsonSerializer.Serialize(new { Diagnostics = book.FileState?.CompatibilityDiagnostics, Sheets = book.Sheets.Select(s => new { s.Name, Charts = s.Data.Charts.Length, Images = s.Images.Length }) }, ProbeRunner.JsonOptions));
        if (book.Sheets[0].Data.Charts.Length != 3) throw new InvalidDataException("Expected the three Excel chart types; inspect import.json.");
        using var session = new DocumentSession(book); var id = book.Sheets[0].Id;
        session.Execute(new SetCellsCommand([new(new(id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(123.5), book.Sheets[0].GetStyleId(1, 1)))]));
        if (book.Sheets.FirstOrDefault(s => s.Name == "Pictures") is { } pictures)
        {
            if (pictures.Images.Length != 3 || pictures.Images.Select(i => i.Anchor.Kind).Distinct().Count() != 3) throw new InvalidDataException("Expected three picture behaviors.");
            // Exercise structure in a standalone picture fixture. The combined source also contains
            // note VML, whose existing workbook-wide structure guard deliberately remains in force.
            using var pictureSession = new DocumentSession(new WorkbookSnapshot(Guid.NewGuid(), 0, book.DateSystem, [pictures], styles: book.Styles, fileState: book.FileState));
            pictureSession.Execute(new ShiftAxisCommand(pictures.Id, SheetAxis.Row, 0, 1));
            pictureSession.Execute(new ShiftAxisCommand(pictures.Id, SheetAxis.Column, 0, 1));
            // Delete the inserted dimensions and insert again: offsets and behaviors must not drift.
            pictureSession.Execute(new ShiftAxisCommand(pictures.Id, SheetAxis.Column, 0, 1, delete: true));
            pictureSession.Execute(new ShiftAxisCommand(pictures.Id, SheetAxis.Row, 0, 1, delete: true));
            if (!pictures.Images.Select(i => i.Anchor).SequenceEqual(pictureSession.Snapshot.GetSheet(pictures.Id).Images.Select(i => i.Anchor))) throw new InvalidDataException("Image anchor drift after insert/delete.");
            pictureSession.Execute(new ShiftAxisCommand(pictures.Id, SheetAxis.Row, 0, 1));
            pictureSession.Execute(new ShiftAxisCommand(pictures.Id, SheetAxis.Column, 0, 1));
            storage.Save(Path.Combine(directory, "pictures-structure.xlsx"), pictureSession.Snapshot, discardUnsupported: true);
            for (var i = 0; i < pictures.Images.Length; i++) session.Execute(new SetImageCommand(pictures.Id, i, pictureSession.Snapshot.Sheets[0].Images[i]));
        }
        if (book.Sheets.FirstOrDefault(s => s.Name == "Protected") is { } protectedSheet)
        {
            var original = protectedSheet.Data.Protection ?? throw new InvalidDataException("Protection missing.");
            session.Execute(new SetCellsCommand([new(new(protectedSheet.Id, 1, 1), CellContent.FromLiteral(CellValue.FromNumber(77), protectedSheet.GetStyleId(1, 1)))]));
            session.Execute(new SetProtectionCommand(protectedSheet.Id, null));
            session.Execute(new SetNoteCommand(protectedSheet.Id, new(0, 0), new("QuickExcel", new([new TextRun("Edited note")])), flattenRichText: true));
            session.Execute(new SetNoteCommand(protectedSheet.Id, new(2, 2), new("QuickExcel", new([new TextRun("New note")]))));
            session.Execute(new SetNoteCommand(protectedSheet.Id, new(3, 3), null));
            session.Execute(new SetProtectionCommand(protectedSheet.Id, original));
        }
        var checks = 0; var outputs = new List<object>();
        foreach (var extension in new[] { ".xlsx", ".xlsq" })
        {
            var path = Path.Combine(directory, "edited" + extension); storage.Save(path, session.Snapshot, discardUnsupported: true);
            var reopened = storage.Open(path); checks++;
            if (!session.Snapshot.Sheets[0].Data.ContentEquals(reopened.Sheets[0].Data)) throw new InvalidDataException("Chart model roundtrip mismatch.");
            var data = reopened.Sheets[0].Data.Charts.Select(c => new { c.Kind, c.Title, Anchor = c.Anchor, c.Legend, c.Minimum, c.Maximum, Series = ChartData.Read(reopened, c) }).ToArray();
            outputs.Add(new { File = path, Charts = data, Pictures = reopened.Sheets.FirstOrDefault(s => s.Name == "Pictures") is { } ps ? ps.Images.Select(i => new { Kind = i.Anchor.Kind.ToString(), Bounds = i.Anchor.Bounds(ps) }).ToArray() : null,
                Protection = reopened.Sheets.FirstOrDefault(s => s.Name == "Protected")?.Data.Protection });
            if (extension == ".xlsx")
            {
                var audit = OpenXmlAudit.Inspect(File.ReadAllBytes(path));
                File.WriteAllText(Path.Combine(directory, "audit.json"), JsonSerializer.Serialize(audit, ProbeRunner.JsonOptions));
                // Retain Excel's Office chart extensions; only exact source SDK findings may recur.
                var unexpected = audit.Issues.Where(i => !sourceAudit.Issues.Any(s => s.Location == i.Location && s.Message == i.Message)).ToArray();
                if (unexpected.Length != 0) throw new InvalidDataException("Output OOXML validation failed; inspect audit.json."); checks++;
            }
        }
        File.WriteAllText(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { Checks = checks, Source = Path.GetFullPath(source), SourceSha256 = OpenXmlAudit.Hash(File.ReadAllBytes(source)), Outputs = outputs }, ProbeRunner.JsonOptions));
        Console.WriteLine($"PASS: {checks} report checks; {directory}"); return 0;
    }

    public static int Capacity(string directory)
    {
        if (Directory.Exists(directory)) throw new IOException("Use a new output directory."); Directory.CreateDirectory(directory);
        var samples = new List<object>(); var storage = new TypedWorkbookStorage();
        foreach (var count in new[] { 25_000, 100_000, 250_000 })
        for (var sample = 0; sample < 3; sample++)
        {
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); var allocated = GC.GetTotalAllocatedBytes(); var initial = GC.GetTotalMemory(true); var watch = Stopwatch.StartNew();
            var cells = Enumerable.Range(0, count).Select(i => new KeyValuePair<CellPosition, CellSnapshot>(new(i / 10, i % 10), new(CellContent.FromLiteral(CellValue.FromNumber(i * .125))))).ToArray();
            var sheet = new WorksheetSnapshot(SheetId.New(), "Capacity", cells); var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]);
            var constructMs = watch.Elapsed.TotalMilliseconds; var path = Path.Combine(directory, $"dense-{count}-{sample}.xlsx");
            watch.Restart(); storage.Save(path, book); var saveMs = watch.Elapsed.TotalMilliseconds;
            watch.Restart(); var opened = storage.Open(path); var openMs = watch.Elapsed.TotalMilliseconds;
            using var session = new DocumentSession(opened); watch.Restart();
            for (var i = 0; i < 20; i++) session.Execute(new SetCellsCommand([new(new(opened.Sheets[0].Id, i, 0), CellContent.FromLiteral(CellValue.FromNumber(-i - 1)))]));
            var editMs = watch.Elapsed.TotalMilliseconds; var history = session.HistoryBytes; for (var i = 0; i < 20; i++) session.Undo();
            var cancellationObserved = false; using var cancellation = new CancellationTokenSource(); cancellation.CancelAfter(1); watch.Restart();
            try { storage.Open(path, cancellation.Token); } catch (OperationCanceledException) { cancellationObserved = true; }
            var cancelMs = watch.Elapsed.TotalMilliseconds;
            samples.Add(new { Cells = count, Sample = sample, ConstructMs = constructMs, SaveMs = saveMs, OpenMs = openMs, TwentyEditsMs = editMs, HistoryBytes = history, CancellationObserved = cancellationObserved, CancelMs = cancelMs,
                AllocatedBytes = GC.GetTotalAllocatedBytes() - allocated, LiveMemoryDelta = GC.GetTotalMemory(false) - initial, PeakWorkingSet = Process.GetCurrentProcess().PeakWorkingSet64, FileBytes = new FileInfo(path).Length });
            File.WriteAllText(Path.Combine(directory, "samples.json"), JsonSerializer.Serialize(samples, ProbeRunner.JsonOptions));
            Console.WriteLine($"{count}/{sample}: open {openMs:F1} ms; save {saveMs:F1} ms; history {history}; cancellation {cancellationObserved}/{cancelMs:F1} ms");
            if (!cancellationObserved || cancelMs > 2000) throw new InvalidOperationException("Cancellation gate failed.");
        }
        var sparse = new WorksheetSnapshot(SheetId.New(), "Sparse", [new(new(WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1), new(CellContent.FromLiteral(CellValue.FromNumber(42))))]);
        var sparseBook = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sparse]); var sparsePath = Path.Combine(directory, "sparse.xlsx");
        var sparseWatch = Stopwatch.StartNew(); storage.Save(sparsePath, sparseBook); var sparseOpen = storage.Open(sparsePath);
        var sparseMs = sparseWatch.Elapsed.TotalMilliseconds;
        var styleSamples = new List<object>();
        foreach (var count in new[] { 4096, 8192 })
        for (var sample = 0; sample < 3; sample++)
        {
            GC.Collect(); var allocated = GC.GetTotalAllocatedBytes(); var watch = Stopwatch.StartNew();
            var styles = Enumerable.Range(0, count).Select(i => new StyleDefinition { FontColor = "FF" + i.ToString("X6") }).ToArray();
            var sheet = new WorksheetSnapshot(SheetId.New(), "Styles", [new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromNumber(42), count - 1)))]);
            var book = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet], styles: styles); var path = Path.Combine(directory, $"styles-{count}-{sample}.xlsx");
            storage.Save(path, book); var saveMs = watch.Elapsed.TotalMilliseconds; watch.Restart(); var opened = storage.Open(path);
            if (opened.Styles.Length != count || opened.Styles[^1] != styles[^1]) throw new InvalidDataException("Style tier mismatch.");
            styleSamples.Add(new { Styles = count, Sample = sample, SaveMs = saveMs, OpenMs = watch.Elapsed.TotalMilliseconds, AllocatedBytes = GC.GetTotalAllocatedBytes() - allocated, PeakWorkingSet = Process.GetCurrentProcess().PeakWorkingSet64, FileBytes = new FileInfo(path).Length });
        }
        File.WriteAllText(Path.Combine(directory, "style-samples.json"), JsonSerializer.Serialize(styleSamples, ProbeRunner.JsonOptions));
        File.WriteAllText(Path.Combine(directory, "environment.json"), JsonSerializer.Serialize(new { Runtime = Environment.Version.ToString(), OS = Environment.OSVersion.ToString(), WorkbookLimits.MaxRows, WorkbookLimits.MaxColumns,
            WorkbookLimits.MaxCells, WorkbookLimits.MaxStyles, SparseStoredCells = sparseOpen.Sheets[0].Cells.Count, SparseRoundtripMs = sparseMs, SparseFileBytes = new FileInfo(sparsePath).Length }, ProbeRunner.JsonOptions));
        return 0;
    }
}
