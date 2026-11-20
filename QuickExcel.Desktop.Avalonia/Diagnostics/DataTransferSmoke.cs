using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private sealed class DataTransferSmokeFiles : ILocalFileDialogs
    {
        public string? OpenPath { get; set; }
        public string? SavePath { get; set; }
        public Task<LocalFileSelection?> OpenAsync(FileDialogRequest request) => Task.FromResult(OpenPath is null ? null : new LocalFileSelection(OpenPath));
        public Task<LocalFileSelection?> SaveAsync(FileDialogRequest request) => Task.FromResult(SavePath is null ? null : new LocalFileSelection(SavePath));
    }

    private async Task RunDataTransferSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool passed, string label) { if (!passed) throw new InvalidOperationException(label); checks.Add(label); }
        var files = new DataTransferSmokeFiles(); _files = files;
        void Drive(DataTransferFormat format, bool export, Action<FieldsDialog>? edit = null, int mode = 0)
        {
            Task Driver(Window window)
            {
                _dialogDriver = Driver;
                if (window is DataTransferWindow transfer)
                {
                    Check(transfer.Format.Items.Count == 4, "Unified transfer view exposes XLSX, CSV/TSV, ODBC and SQLite");
                    transfer.Format.SelectedIndex = (int)format; transfer.Direction.SelectedIndex = export ? 1 : 0; transfer.Mode.SelectedIndex = mode;
                    window.Close(transfer.Request);
                }
                else { edit?.Invoke((FieldsDialog)window); window.Close(true); }
                return Task.CompletedTask;
            }
            _dialogDriver = Driver;
        }
        try
        {
            _workspace.New(); ShowBook(); _workspace.Editor.SetCell(0, 0, 0, "existing"); Synchronize();
            var menu = ((MenuItem)_menu.Items[0]!).Items.OfType<MenuItem>().Single(i => Equals(i.Header, UiText.Get("数据导入 / 导出…")));
            Check(menu.IsEnabled, "Unified transfer command is available from File");
            var original = Document; var originalRevision = Document.Snapshot.Revision;
            _dialogDriver = window => { Check(window is DataTransferWindow, "File command opens common transfer view"); window.Close(null); return Task.CompletedTask; };
            await DataTransferAsync(); Check(ReferenceEquals(original, Document) && originalRevision == Document.Snapshot.Revision, "Cancel transfer view leaves workbook unchanged");

            files.OpenPath = Path.Combine(directory, "source.csv"); await File.WriteAllTextAsync(files.OpenPath, "001,=1+2\r\n002,中文\r\n");
            Drive(DataTransferFormat.Text, false, dialog =>
            {
                if (dialog.Title == UiText.Get("数据导入预览")) ((ComboBox)dialog.Input("destination")).SelectedIndex = 1;
            });
            await DataTransferAsync();
            Check(ReferenceEquals(original, Document) && Document.Snapshot.Sheets.Length == 2 && Document.Snapshot.Sheets[1].Name == "Sheet1 (2)", "CSV import adds sheet with previewed collision rename");
            Check(Document.Snapshot.Sheets[1].GetCell(0, 0)!.Value.TextValue == "001" && Document.Snapshot.Sheets[1].GetCell(0, 1)!.Content.Kind == CellContentKind.Literal, "Unified CSV preserves literal identifiers and formula-like text");
            await History(false); Check(Document.Snapshot.Sheets.Length == 1 && Document.GetCellInput(0, 0, 0) == "existing", "Current-workbook data import is a single undo transaction");

            var fixture = new WorkbookSnapshot(Guid.NewGuid(), 0, ExcelDateSystem.Windows1900,
            [new(SheetId.New(), "Sheet1", [KeyValuePair.Create(new CellPosition(0, 0), new CellSnapshot(CellContent.FromLiteral(CellValue.FromText("xlsx"))))]),
             new(SheetId.New(), "Extra", [KeyValuePair.Create(new CellPosition(0, 0), new CellSnapshot(CellContent.FromLiteral(CellValue.FromNumber(42))))])]);
            files.OpenPath = Path.Combine(directory, "source.xlsx"); new TypedWorkbookStorage().Save(files.OpenPath, fixture);
            Drive(DataTransferFormat.Xlsx, false, dialog =>
            {
                if (dialog.Title == UiText.Get("选择导入工作表")) ((CheckBox)dialog.Input("sheet1")).IsChecked = false;
                if (dialog.Title == UiText.Get("数据导入预览")) { ((ComboBox)dialog.Input("destination")).SelectedIndex = 1; ((TextBox)dialog.Input("name0")).Text = "Selected"; }
            });
            await DataTransferAsync();
            Check(Document.Snapshot.Sheets.Length == 2 && Document.Snapshot.Sheets[1].Name == "Selected" && Document.Snapshot.Sheets[1].GetCell(0, 0)!.Value.TextValue == "xlsx", "XLSX sheet selection and custom target name reach import command");
            Drive(DataTransferFormat.Xlsx, false, dialog =>
            {
                if (dialog.Title == UiText.Get("数据导入预览")) { ((ComboBox)dialog.Input("destination")).SelectedIndex = 1; ((ComboBox)dialog.Input("conflict")).SelectedIndex = 1; }
            });
            await DataTransferAsync();
            Check(Document.Snapshot.Sheets.Length == 3 && Document.Snapshot.Sheets[^1].Name == "Extra" && Document.GetCellInput(0, 0, 0) == "existing", "XLSX skip conflict retains original sheet and imports other selected sheet");

            Task OdbcDriver(Window window)
            {
                _dialogDriver = OdbcDriver;
                if (window is DataTransferWindow transfer) { transfer.Format.SelectedIndex = (int)DataTransferFormat.Odbc; window.Close(transfer.Request); }
                else
                {
                    Check(window.Title == UiText.Get("连接 SQL 数据库") && ((TextBox)((FieldsDialog)window).Input("connection")).PasswordChar != '\0', "ODBC route masks ephemeral connection credentials before any connection");
                    window.Close(false);
                }
                return Task.CompletedTask;
            }
            _dialogDriver = OdbcDriver; await DataTransferAsync();

            _workspace.New(); ShowBook(); _workspace.Editor.SetCell(0, 0, 0, "id"); _workspace.Editor.SetCell(0, 0, 1, "note");
            _workspace.Editor.SetCell(0, 1, 0, "7"); _workspace.Editor.SetCell(0, 1, 1, "中文 ' quote"); Synchronize();
            files.SavePath = Path.Combine(directory, "roundtrip.sqlite");
            Drive(DataTransferFormat.Sqlite, true, dialog => { if (dialog.Title == UiText.Get("导出到数据库")) ((TextBox)dialog.Input("table")).Text = "Orders"; });
            await DataTransferAsync();
            var data = DatabaseTransfer.ImportTable(new(DatabaseProvider.Sqlite, files.SavePath), new("Orders"));
            Check(data.Rows == 1 && data.Workbook.Sheets[0].GetCell(1, 1)!.Value.TextValue == "中文 ' quote", "SQLite export routes parameterized data through confirmation to real file");
            files.OpenPath = files.SavePath; _discardChoice = () => Task.FromResult(SaveDecision.Discard);
            Drive(DataTransferFormat.Sqlite, false);
            await DataTransferAsync();
            Check(Document.Snapshot.Sheets[0].GetCell(1, 1)!.Value.TextValue == "中文 ' quote" && _workspace.FilePath is null && Document.IsDirty && !Document.ScriptsEnabled, "SQLite import creates an unsaved unbound workbook without script authorization");
            var importedDocument = Document; var discardPrompted = false;
            _discardChoice = () => { discardPrompted = true; return Task.FromResult(SaveDecision.Cancel); };
            Check(!await ConfirmDiscardAsync() && discardPrompted && ReferenceEquals(importedDocument, Document) && Document.IsDirty,
                "Cancel closing an imported workbook preserves its unsaved data");
            _discardChoice = () => Task.FromResult(SaveDecision.Discard);

            files.SavePath = Path.Combine(directory, "cancelled.sqlite");
            Task CancelWrite(Window window)
            {
                _dialogDriver = CancelWrite;
                window.Close(window.Title != UiText.Get("确认数据库写入")); return Task.CompletedTask;
            }
            _dialogDriver = CancelWrite; await TransferDatabaseAsync(true, true);
            Check(!File.Exists(files.SavePath), "Cancel database confirmation does not create a file");
            Task StaleWrite(Window window)
            {
                _dialogDriver = StaleWrite;
                if (window.Title == UiText.Get("确认数据库写入")) _workspace.Editor.SetCell(0, 5, 0, "newer");
                window.Close(true); return Task.CompletedTask;
            }
            _dialogDriver = StaleWrite; var rejected = false;
            try { await TransferDatabaseAsync(true, true); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected && !File.Exists(files.SavePath) && Document.GetCellInput(0, 5, 0) == "newer", "Stale database preview rejects write and retains newer workbook edit");

            files.SavePath = Path.Combine(directory, "selected.xlsx"); _grid.Select(new CellPosition(0, 0));
            Drive(DataTransferFormat.Xlsx, true, mode: 3); await DataTransferAsync();
            var exported = new TypedWorkbookStorage().Open(files.SavePath);
            Check(exported.Sheets.Length == 1 && exported.Sheets[0].Cells.Count == 1 && exported.Sheets[0].GetCell(0, 0)!.Value.TextValue == "id", "XLSX selection export writes only selected values");

            files.OpenPath = Path.Combine(directory, "source.xlsx");
            var sourceBytes = File.ReadAllBytes(files.OpenPath);
            _discardChoice = () => Task.FromResult(SaveDecision.Discard);
            Drive(DataTransferFormat.Xlsx, false, mode: 1); await DataTransferAsync();
            Check(Document.IsDirty && _workspace.FilePath is null && Document.Snapshot.FileState is not null &&
                Document.Snapshot.Sheets.Select(sheet => sheet.Name).SequenceEqual(new[] { "Sheet1", "Extra" }) && !Document.ScriptsEnabled,
                "Complete XLSX import retains source metadata as an unsaved unbound copy");
            await _workspace.SaveAsync(new(Path.Combine(directory, "full-copy.xlsx")));
            Check(!Document.IsDirty && sourceBytes.SequenceEqual(File.ReadAllBytes(files.OpenPath)) &&
                new TypedWorkbookStorage().Open(Path.Combine(directory, "full-copy.xlsx")).Sheets.Length == 2,
                "Complete XLSX copy saves and reopens without modifying the source");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; _discardChoice = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
