using System.Text.Json;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls.Presenters;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private sealed class SmokeFiles : ILocalFileDialogs
    {
        public string? NextSave;
        public string? NextOpen;
        public Task<LocalFileSelection?> OpenAsync(FileDialogRequest request) => Task.FromResult(NextOpen is null ? null : new LocalFileSelection(NextOpen));
        public Task<LocalFileSelection?> SaveAsync(FileDialogRequest request) => Task.FromResult(NextSave is null ? null : new LocalFileSelection(NextSave));
    }
    private async Task RunWorkspaceSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory);
        var checks = new List<string>(); Exception? error = null;
        void Check(bool condition, string description) { if (!condition) throw new InvalidOperationException(description); checks.Add(description); }
        try
        {
            _files = new SmokeFiles(); var files = (SmokeFiles)_files;
            ShowBook(); _workspace.New();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            Check(_formula.GetVisualDescendants().OfType<TextPresenter>().Any() && _grid.CellEditor.GetVisualDescendants().OfType<TextPresenter>().Any(), "Formula and grid editors have real native TextBox templates and IME presenters");
            Check(!Document.IsDirty && !_grid.IsEditing && !_formulaDirty, "Binding native editors does not create drafts or dirty the workbook");
            var presenter = _grid.CellEditor.GetVisualDescendants().OfType<TextPresenter>().Single();
            presenter.PreeditText = "ni";
            Check(_grid.IsComposing && !CommitEditors(), "Synthetic preedit blocks document commands");
            presenter.PreeditText = null; _grid.Cancel();
            _grid.Select(new(0, 0)); _grid.BeginEdit(); _grid.CellEditor.Text = "12";
            Check(CommitEditors(), "Native cell editor commits through shared controller");
            Check(Document.GetCellInput(0, 0, 0) == "12", "Native editor writes its owner cell");
            await History(false); Check(Document.GetCellInput(0, 0, 0) == "", "Undo removes one edit");
            await History(true); Check(Document.GetCellInput(0, 0, 0) == "12", "Redo restores edit");
            _grid.Select(new(0, 1)); _grid.BeginEdit(); _grid.CellEditor.Text = "=A1+3"; Check(CommitEditors(), "Formula edit commits");
            await CalculateAsync(); Check(Document.GetCellValue(0, 0, 1) == "15", "Shared typed calculator evaluates native grid input");
            _grid.Select(new(1, 0)); _formula.Focus(); _formula.Text = "原格草稿";
            _grid.Select(new(2, 1));
            Check(Document.GetCellInput(0, 1, 0) == "原格草稿" && Document.GetCellInput(0, 2, 1) == "", "Formula-bar draft remains owned by original cell during selection change");
            _grid.BeginEdit(); _grid.CellEditor.Text = "取消输入"; _grid.Cancel();
            Check(Document.GetCellInput(0, 2, 1) == "" && !_grid.IsEditing, "Cancel drops only the uncommitted grid draft");
            _discardChoice = () => Task.FromResult(SaveDecision.Cancel);
            Check(!await ConfirmDiscardAsync() && Document.IsDirty, "Unsaved cancel retains document");
            _discardChoice = () => Task.FromResult(SaveDecision.Save);
            Check(!await ConfirmDiscardAsync() && Document.IsDirty, "Save-dialog cancel also cancels document transition");
            files.NextSave = Path.Combine(directory, "工作簿 roundtrip.xlsq");
            Check(await SaveAsync(false), "Save-as succeeds through native host file workflow");
            Check(!Document.IsDirty && File.Exists(files.NextSave) && _workspace.Version?.Sha256 is not null, "Save marks captured revision and remembers fingerprint");
            var saved = files.NextSave;
            await OpenPathAsync(saved);
            Check(Document.GetCellInput(0, 1, 0) == "原格草稿" && Document.GetCellValue(0, 0, 1) == "15", "Reopen preserves Unicode, formula and calculated value");
            Check(!Document.ScriptsEnabled, "Open never inherits prior script authorization");
            var revision = Document.Snapshot.Revision;
            _grid.Select(new(3, 3));
            Check(Document.Snapshot.Revision == revision && !Document.IsDirty, "Navigation does not mark dirty");
            await AddSheetAsync(); Check(Document.Snapshot.Sheets.Length == 2, "Add sheet uses shared command/history");
            await History(false); Check(Document.Snapshot.Sheets.Length == 1, "Undo sheet creation restores original worksheet");
            await History(true); Check(Document.Snapshot.Sheets.Length == 2, "Redo sheet creation preserves identity");
            await ZoomAsync(125); Check(Document.Snapshot.Sheets[SheetIndex].View.Zoom == 125 && Document.IsDirty, "Zoom is a persistent undoable view command");
            await PlaceSheets(SheetNavigationPlacement.Left); Check(_settings.Value.Sheets == SheetNavigationPlacement.Left, "Sheet placement saves to isolated settings path");
            await PlaceSheets(SheetNavigationPlacement.Bottom);
            Check(await SaveAsync(false), "Save uses the opened version");
            var identity = Document.Session.SessionId;
            var invalid = Path.Combine(directory, "invalid.xlsx"); await File.WriteAllTextAsync(invalid, "invalid zip");
            try { await OpenPathAsync(invalid); throw new InvalidOperationException("Invalid workbook unexpectedly opened"); }
            catch (Exception ex) when (ex is not InvalidOperationException || ex.Message != "Invalid workbook unexpectedly opened")
            { Check(Document.Session.SessionId == identity && !_workspace.IsBusy, "Failed open retains previous document and releases busy state"); }
            _grid.BeginEdit(); _grid.CellEditor.Text = "保留冲突编辑"; Check(CommitEditors(), "Edit before external conflict commits");
            await File.AppendAllTextAsync(saved, "external change");
            Check(!await SaveAsync(false) && Document.IsDirty && Document.Session.SessionId == identity, "External modification rejects save and retains edits");
            files.NextSave = Path.Combine(directory, "conflict-copy.xlsq");
            Check(await SaveAsync(true) && !Document.IsDirty, "Save-as resolves conflict without overwriting changed original");
            _grid.BeginEdit(); _grid.CellEditor.Text = "关闭时提交";
            _discardChoice = () => Task.FromResult(SaveDecision.Cancel);
            Close();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            Check(IsVisible && !_allowClose && Document.IsDirty && !_grid.IsEditing, "Real window Closing handler commits pending native editor then honors cancel");
            _discardChoice = () => Task.FromResult(SaveDecision.Save);
            Check(await ConfirmDiscardAsync() && !Document.IsDirty, "Close save waits until checked save completes");
            Check(Directory.Exists(Path.Combine(directory, "settings")), "Smoke settings remain isolated");
            await RunMergedClipboardChecksAsync(Check);
            await RunExternalClipboardChecksAsync(Check);
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            using var screenshot = new RenderTargetBitmap(new PixelSize((int)Bounds.Width, (int)Bounds.Height), new Vector(96, 96));
            screenshot.Render(this); screenshot.Save(Path.Combine(directory, "workspace.png"), PngBitmapEncoderOptions.Default);
            Check(_grid.Bounds.Width > 500 && _grid.Bounds.Height > 200, "Visible native grid has usable layout");
        }
        catch (Exception ex) { error = ex; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new
        { passed = error is null, checks, error = error?.ToString(), scaling = RenderScaling,
            runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            os = System.Runtime.InteropServices.RuntimeInformation.OSDescription }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true;
        (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(error is null ? 0 : 1);
    }
}
