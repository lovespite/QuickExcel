using System.Text;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private sealed class TextExchangeSmokeFiles(string path) : ILocalFileDialogs
    {
        public bool Cancel { get; set; }
        public string? OpenPath { get; set; }
        public Task<LocalFileSelection?> OpenAsync(FileDialogRequest request) => Task.FromResult<LocalFileSelection?>(OpenPath is null ? null : new(OpenPath));
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
            var exportMenu = ((MenuItem)_menu.Items[0]!).Items.OfType<MenuItem>().Single(i => Equals(i.Header, UiText.Get("导出")));
            var textExport = exportMenu.Items.OfType<MenuItem>().Single(i => Equals(i.Header, UiText.Get("导出 CSV / TSV")));
            Check(exportMenu.IsEnabled && textExport.IsEnabled && exportMenu.Items.OfType<MenuItem>().Select(i => i.Header as string)
                .SequenceEqual(new[] { "导出值快照", "导出可移植公式", "导出 CSV / TSV" }.Select(UiText.Get)), "File Export submenu groups all three export commands");
            _dialogDriver = window => { ((FieldsDialog)window).Close(true); return Task.CompletedTask; };
            textExport.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); await _operation;
            var bytes = File.ReadAllBytes(path);
            Check(bytes.AsSpan().StartsWith(Encoding.UTF8.Preamble), "Default export writes UTF-8 BOM");
            Check(DelimitedText.ReadRows(Encoding.UTF8.GetString(bytes.AsSpan(3)), ',', 10).Single().Single() == "中文,\"甲\"\n乙", "Desktop export preserves quoted Unicode and newlines");
            Check(Document.Snapshot.Revision == revision && Document.IsDirty && _workspace.FilePath is null, "Export preserves document history, dirty state and save identity");
            _dialogDriver = window => { window.Close(false); return Task.CompletedTask; };
            await ExportTextAsync(); Check(bytes.SequenceEqual(File.ReadAllBytes(path)), "Cancel options leaves output unchanged");
            files.Cancel = true; _dialogDriver = window => { window.Close(true); return Task.CompletedTask; };
            await ExportTextAsync(); Check(bytes.SequenceEqual(File.ReadAllBytes(path)), "Cancel file picker leaves output unchanged");
            var input = Path.Combine(directory, "input.csv"); await File.WriteAllTextAsync(input, "001,2.5\r\n=1+1,3\r\n", new UTF8Encoding(true));
            files.OpenPath = input;
            var previous = Document;
            _dialogDriver = window => { window.Close(false); return Task.CompletedTask; };
            await ImportTextAsync(); Check(ReferenceEquals(previous, Document), "Cancel import preserves the current document");
            _discardChoice = () => Task.FromResult(SaveDecision.Discard);
            async Task ImportDriver(Window window)
            {
                _dialogDriver = ImportDriver;
                await global::Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, global::Avalonia.Threading.DispatcherPriority.Background);
                if (window.Title == UiText.Get("文本导入预览")) window.GetVisualDescendants().OfType<ComboBox>().ElementAt(1).SelectedIndex = 1;
                window.Close(true);
            }
            _dialogDriver = ImportDriver;
            await ImportTextAsync();
            Check(Document.GetCellInput(0, 0, 0) == "001" && Document.Snapshot.Sheets[0].GetCell(1, 0)!.Content.Kind == QuickExcel.Domain.CellContentKind.Literal,
                "Import defaults identifiers and leading equals to literal text");
            Check(Document.Snapshot.Sheets[0].GetCell(0, 1)!.Value.NumberValue == 2.5, "Per-column import converts only explicitly selected numeric columns");
            Check(_workspace.FilePath is null && !Document.ScriptsEnabled && File.ReadAllText(input).StartsWith("001,2.5"), "Import creates an unbound document without modifying the source or authorizing scripts");
            previous = Document;
            Task StaleDriver(Window window)
            {
                _dialogDriver = StaleDriver;
                if (window.Title == UiText.Get("确认导入")) _workspace.Editor.SetCell(0, 5, 0, "changed");
                window.Close(true); return Task.CompletedTask;
            }
            _dialogDriver = StaleDriver;
            var staleRejected = false;
            try { await ImportTextAsync(); } catch (InvalidOperationException) { staleRejected = true; }
            Check(staleRejected && ReferenceEquals(previous, Document) && Document.GetCellInput(0, 5, 0) == "changed", "Stale preview rejects replacement and retains newer edits");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
