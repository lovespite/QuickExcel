using System.Text;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Core;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private sealed class TextExchangeSmokeFiles(string path) : ILocalFileDialogs
    {
        public bool Cancel { get; set; }
        public Task<LocalFileSelection?> OpenAsync(FileDialogRequest request) => Task.FromResult<LocalFileSelection?>(null);
        public Task<LocalFileSelection?> SaveAsync(FileDialogRequest request) => Task.FromResult<LocalFileSelection?>(Cancel ? null : new(path));
    }

    private async Task RunTextExchangeSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new List<string>(); Exception? failure = null;
        void Check(bool passed, string label) { if (!passed) throw new InvalidOperationException(label); checks.Add(label); }
        try
        {
            var path = Path.Combine(directory, "export.csv"); var files = new TextExchangeSmokeFiles(path); _files = files;
            _workspace.New(); ShowBook(); _workspace.Editor.SetCell(0, 0, 0, "中文,\"甲\"\n乙"); Synchronize();
            var before = Document.Snapshot; var revision = before.Revision;
            Check(((MenuItem)_menu.Items[0]!).Items.OfType<MenuItem>().Any(i => i.Header as string == UiText.Get("导出 CSV / TSV") && i.IsEnabled), "Normal File menu exposes text export");
            _dialogDriver = window => { ((FieldsDialog)window).Close(true); return Task.CompletedTask; };
            await ExportTextAsync();
            var bytes = File.ReadAllBytes(path);
            Check(bytes.AsSpan().StartsWith(Encoding.UTF8.Preamble), "Default export writes UTF-8 BOM");
            Check(DelimitedText.ReadRows(Encoding.UTF8.GetString(bytes.AsSpan(3)), ',', 10).Single().Single() == "中文,\"甲\"\n乙", "Desktop export preserves quoted Unicode and newlines");
            Check(Document.Snapshot.Revision == revision && Document.IsDirty && _workspace.FilePath is null, "Export preserves document history, dirty state and save identity");
            _dialogDriver = window => { window.Close(false); return Task.CompletedTask; };
            await ExportTextAsync(); Check(bytes.SequenceEqual(File.ReadAllBytes(path)), "Cancel options leaves output unchanged");
            files.Cancel = true; _dialogDriver = window => { window.Close(true); return Task.CompletedTask; };
            await ExportTextAsync(); Check(bytes.SequenceEqual(File.ReadAllBytes(path)), "Cancel file picker leaves output unchanged");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
