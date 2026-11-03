using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private sealed class CompatibilitySmokeFiles : ILocalFileDialogs
    {
        public string? NextSave;
        public int SaveCalls;
        public FileDialogRequest? LastSave;
        public Task<LocalFileSelection?> OpenAsync(FileDialogRequest request) => Task.FromResult<LocalFileSelection?>(null);
        public Task<LocalFileSelection?> SaveAsync(FileDialogRequest request)
        {
            SaveCalls++; LastSave = request;
            return Task.FromResult(NextSave is null ? null : new LocalFileSelection(NextSave));
        }
    }

    private async Task RunCompatibilitySmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        static Task Layout() => Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background).GetTask();
        static void Click(Window window, string caption) => window.GetVisualDescendants().OfType<Button>()
            .Single(button => button.Content as string == caption).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        void AcceptReport() => _dialogDriver = async window => { await Layout(); Click(window, "另存受支持副本"); };
        try
        {
            var files = new CompatibilitySmokeFiles(); _files = files; _clipboard = new MemoryClipboard();
            var source = Path.Combine(directory, "compatibility-source.xlsx");
            var sheet = new WorksheetSnapshot(SheetId.New(), "兼容性示例", [
                new(new(0, 0), new(CellContent.FromLiteral(CellValue.FromText("原始值")))),
                new(new(1, 1), new(CellContent.FromFormula("=UNKNOWN_FN(1)"), new(CellValue.FromNumber(7), CalculationState.Stale, 0, 0, 0))),
                new(new(2, 2), new(CellContent.FromRichText(new([new TextRun("保留富文本")]))))]);
            new TypedWorkbookStorage().Save(source, new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900, [sheet]));
            using (var zip = ZipFile.Open(source, ZipArchiveMode.Update))
            {
                const string part = "xl/worksheets/sheet1.xml";
                XDocument xml; using (var input = zip.GetEntry(part)!.Open()) xml = XDocument.Load(input);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                xml.Root!.Add(new XElement(ns + "customSheetViews", new XElement(ns + "customSheetView",
                    new XAttribute("guid", "{86A0202C-6747-4F98-BA3B-B964C0800EA5}"), new XAttribute("view", "normal"))));
                zip.GetEntry(part)!.Delete(); using var output = zip.CreateEntry(part).Open(); xml.Save(output);
            }
            var original = await File.ReadAllBytesAsync(source);
            await OpenPathAsync(source); await Layout();
            Check(Document.Snapshot.FileState!.CompatibilityIssues.Length > 0, "Fixture contains unsupported worksheet views");
            Check(await SaveAsync(false) && files.SaveCalls == 0, "Unchanged source saves without a compatibility prompt or picker");
            Check(original.SequenceEqual(File.ReadAllBytes(source)), "Unchanged save preserves original package bytes");
            _grid.Select(new(0, 0)); _grid.BeginEdit(); _grid.CellEditor.Text = "修改后值";
            var beforeReport = Document.Snapshot.Revision;
            _dialogDriver = async window =>
            {
                await Layout();
                Check(window.Name == "CompatibilityReportDialog" && window.Title == "保存兼容性报告", "Save opens the detailed compatibility report");
                var groups = window.GetVisualDescendants().OfType<Expander>().ToArray();
                var loss = groups.Single(group => group.Header!.ToString()!.StartsWith("潜在损失"));
                Check(loss.IsExpanded && groups.Length >= 3, "Save report expands losses and separates calculation and preservation notices");
                var rows = ((StackPanel)loss.Content!).Children.OfType<ListBox>().Single().Items.Cast<CompatibilityIssueView>().ToArray();
                Check(rows.Any(row => row.Feature.Contains("customSheetViews") && row.Location.Contains("兼容性示例") && row.Location.Contains("xl/worksheets/sheet1.xml") && row.Impact.Contains("不会写入受支持副本")), "Loss rows include feature, worksheet, package location and save impact");
                Check(Document.GetCellInput(0, 0, 0) == "修改后值" && !_grid.IsEditing, "Save report inspects the committed cell draft");
                beforeReport = Document.Snapshot.Revision;
                Click(window, "复制报告"); await Layout();
                var copied = ((MemoryClipboard)_clipboard).Value.Text!;
                Check(copied.Contains("customSheetViews") && copied.Contains("兼容性示例") && copied.Contains("UNKNOWN_FN") && copied.Contains("富文本"), "Copy includes the full loss, calculation and preserved-content report");
                using var bitmap = new RenderTargetBitmap(new PixelSize((int)window.Bounds.Width, (int)window.Bounds.Height), new Vector(96, 96));
                bitmap.Render(window); bitmap.Save(Path.Combine(directory, "save-compatibility-report.png"), PngBitmapEncoderOptions.Default);
                Click(window, "取消");
            };
            Check(!await SaveAsync(false) && Document.IsDirty && files.SaveCalls == 0, "Cancel retains edits without opening the save picker");
            Check(Document.Snapshot.Revision == beforeReport && !Document.ScriptsEnabled, "Viewing and copying the report does not mutate content or enable scripts");
            _dialogDriver = async window => { await Layout(); window.Close(); };
            Check(!await SaveAsync(true) && Document.IsDirty && files.SaveCalls == 0, "Closing the report cancels Save As without discarding edits");
            AcceptReport();
            Check(!await SaveAsync(false) && Document.IsDirty && files.SaveCalls == 1, "Canceling the picker after report confirmation retains edits");
            Check(files.LastSave!.SuggestedName!.Contains("受支持副本"), "Save picker suggests an explicit supported-copy filename");
            files.NextSave = source; AcceptReport();
            Check(!await SaveAsync(false) && Document.IsDirty, "Confirmed lossy save still rejects overwriting the original file");
            Check(original.SequenceEqual(File.ReadAllBytes(source)), "Canceled and rejected saves leave original bytes intact");
            files.NextSave = Path.Combine(directory, "supported-copy.xlsx"); AcceptReport();
            Check(await SaveAsync(false) && !Document.IsDirty && _workspace.FilePath == files.NextSave, "Explicit report confirmation saves and binds the supported copy");
            var copy = new TypedWorkbookStorage().Open(files.NextSave);
            Check(copy.Sheets[0].GetCell(0, 0)!.Content.LiteralValue.TextValue == "修改后值" && copy.FileState!.CompatibilityIssues.IsEmpty, "Supported copy contains the edits and removes only unsupported content");
            Check(copy.Sheets[0].GetCell(2, 2)!.Content.RichText?.Text == "保留富文本", "Supported copy retains modeled rich text");
            Check(original.SequenceEqual(File.ReadAllBytes(source)), "Successful copy also preserves the original file");
            await OpenPathAsync(files.NextSave);
            var pickerCalls = files.SaveCalls;
            Check(await SaveAsync(false) && files.SaveCalls == pickerCalls, "Reopened supported copy saves without a loss prompt");
            _dialogDriver = async window =>
            {
                await Layout();
                Check(window.Name == "CompatibilityReportDialog" && window.Title == "兼容性报告" &&
                    !window.GetVisualDescendants().OfType<Button>().Any(button => button.Content as string == "另存受支持副本"), "Menu report shares the detailed view without save actions");
                Click(window, "关闭");
            };
            await CompatibilityAsync();
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
