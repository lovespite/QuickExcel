using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls.Presenters;
using Avalonia.Input;
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
            await CheckFormulaBarHeightAsync(Check);
            await CheckEditorLineBreaksAsync(Check);
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
            Check(Document.GetCellInput(0, 6, 0) == "=1+\n3+2" && Document.GetCellInput(0, 6, 1) == "甲\n续乙",
                "Reopen preserves line breaks entered in both worksheet editors");
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

    private async Task CheckEditorLineBreaksAsync(Action<bool, string> check)
    {
        InputElement Focused() => (InputElement)FocusManager!.GetFocusedElement()!;
        void KeyEvent(Key key, KeyModifiers modifiers = KeyModifiers.None, bool release = false) =>
            Focused().RaiseEvent(new KeyEventArgs { RoutedEvent = release ? KeyUpEvent : KeyDownEvent, Key = key, KeyModifiers = modifiers });
        void AltEnter()
        {
            KeyEvent(Key.LeftAlt, KeyModifiers.Alt);
            KeyEvent(Key.Enter, KeyModifiers.Alt);
            KeyEvent(Key.Enter, KeyModifiers.Alt, release: true);
            KeyEvent(Key.LeftAlt, release: true);
        }
        void TypeText(string text) => Focused().RaiseEvent(new TextInputEventArgs { RoutedEvent = TextInputEvent, Text = text });

        _grid.Select(new(6, 0)); _formula.Focus(); _formula.Text = "=1+2";
        _formula.CaretIndex = _formula.SelectionStart = _formula.SelectionEnd = 3;
        var revision = Document.Snapshot.Revision;
        var presenter = _formula.GetVisualDescendants().OfType<TextPresenter>().Single();
        presenter.PreeditText = "zhong";
        AltEnter();
        check(_formula.IsComposing && _formula.Text == "=1+2" && Document.Snapshot.Revision == revision,
            "Alt+Enter leaves an active IME composition and its draft intact");
        presenter.PreeditText = null;
        AltEnter();
        check(_formula.Text == "=1+\n2" && _formula.CaretIndex == 4 && _formulaDirty && _formula.IsKeyboardFocusWithin &&
            Document.Snapshot.Revision == revision, "Alt+Enter inserts one LF at the formula caret and keeps the draft active");
        TypeText("3+");
        check(_formula.Text == "=1+\n3+2" && Document.Snapshot.Revision == revision, "Typing continues at the formula caret after releasing Alt");
        KeyEvent(Key.Enter);
        check(!_formulaDirty && Document.GetCellInput(0, 6, 0) == "=1+\n3+2", "Plain Enter commits the multiline formula");
        await CalculateAsync();
        check(Document.GetCellValue(0, 6, 0) == "6", "The committed multiline formula calculates");

        _grid.Select(new(6, 1)); _grid.BeginEdit();
        var cell = _grid.CellEditor;
        cell.Text = "甲替换乙"; cell.SelectionStart = 1; cell.SelectionEnd = 3;
        revision = Document.Snapshot.Revision;
        AltEnter();
        check(cell.Text == "甲\n乙" && cell.CaretIndex == 2 && cell.SelectionStart == cell.SelectionEnd && _grid.IsEditing &&
            _grid.Current == new CellPosition(6, 1) && Document.Snapshot.Revision == revision,
            "Alt+Enter replaces the cell editor selection with one LF and stays in the original cell");
        TypeText("续");
        check(cell.Text == "甲\n续乙" && Document.Snapshot.Revision == revision, "Typing continues at the cell caret after releasing Alt");
        KeyEvent(Key.Enter);
        check(Document.GetCellInput(0, 6, 1) == "甲\n续乙" && _grid.Current == new CellPosition(7, 1),
            "Plain Enter commits the multiline cell and moves down");

        _grid.Select(new(0, 0)); AltEnter();
        check(_grid.IsEditing && cell.Text == "12\n", "Alt+Enter from a selected cell appends a line break to its existing text");
        KeyEvent(Key.Escape);
        check(!_grid.IsEditing && Document.GetCellInput(0, 0, 0) == "12", "Escape discards the line break with the rest of the draft");
        cell.IsReadOnly = true; AltEnter();
        check(!_grid.IsEditing && cell.Text == "12", "Alt+Enter respects a read-only cell editor");
        cell.IsReadOnly = false;
    }

    private async Task CheckFormulaBarHeightAsync(Action<bool, string> check)
    {
        static async Task Layout() => await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        _grid.Select(new(5, 0));
        await Layout();
        var compact = _formula.Bounds.Height;
        var sheetHeight = _grid.Bounds.Height;
        var multiline = string.Join("\n", Enumerable.Repeat("多行内容", 16));
        _formula.Focus();
        _formula.Text = multiline;
        await Layout();
        var scroll = _formula.GetVisualDescendants().OfType<ScrollViewer>().Single();
        check(_formula.Bounds.Height == compact && _grid.Bounds.Height == sheetHeight && scroll.Extent.Height > scroll.Viewport.Height,
            "Multiline input scrolls inside the formula bar at its original height");

        var splitter = _book.GetVisualDescendants().OfType<GridSplitter>().Single();
        var start = splitter.TranslatePoint(new Point(splitter.Bounds.Width / 2, splitter.Bounds.Height / 2), this)!.Value;
        var end = start + new Vector(0, 48);
        using var mouse = new Pointer(202, PointerType.Mouse, true);
        splitter.RaiseEvent(new PointerPressedEventArgs(splitter, mouse, this, start, 1,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed), KeyModifiers.None, 1));
        splitter.RaiseEvent(new PointerEventArgs(PointerMovedEvent, splitter, mouse, this, end, 2,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other), KeyModifiers.None));
        splitter.RaiseEvent(new PointerReleasedEventArgs(splitter, mouse, this, end, 3,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased), KeyModifiers.None, MouseButton.Left));
        await Layout();
        var expanded = _formula.Bounds.Height;
        check(expanded > compact && _grid.Bounds.Height < sheetHeight && _formulaDirty && _formula.Text == multiline,
            "Dragging the formula bar divider expands the editor while preserving its draft");
        _grid.Select(new(5, 1));
        await Layout();
        check(string.IsNullOrEmpty(_formula.Text) && _formula.Bounds.Height == expanded, "Selecting an empty cell retains the chosen formula bar height");
        _grid.Select(new(5, 0));
        await Layout();
        check(_formula.Text == multiline && _formula.Bounds.Height == expanded, "Selecting a multiline cell retains the chosen formula bar height");
        splitter.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = Key.Up });
        await Layout();
        check(_formula.Bounds.Height < expanded && _formula.Text == multiline, "The formula bar divider also shrinks with the keyboard");
    }
}
