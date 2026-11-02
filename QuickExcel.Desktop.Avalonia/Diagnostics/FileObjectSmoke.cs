using Avalonia.Automation.Peers;
using Avalonia.Automation.Provider;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Presentation;
using QuickExcel.Storage;
using SkiaSharp;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunFileObjectChecksAsync(string directory, Action<bool, string> check)
    {
        var previousFiles = _files; var files = new SmokeFiles(); _files = files;
        static void Click(Window window, string caption) => window.GetVisualDescendants().OfType<Button>().Single(b => b.Content as string == caption).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        static async Task Ready() => await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
        try
        {
            var imagePath = Path.Combine(directory, "native-image.png");
            using (var bitmap = new SKBitmap(40, 20)) { bitmap.Erase(SKColors.Magenta); using var encoded = bitmap.Encode(SKEncodedImageFormat.Png, 100); using var output = File.Create(imagePath); encoded.SaveTo(output); }
            files.NextOpen = imagePath; _grid.Select(new(8, 5));
            _dialogDriver = async chooser =>
            {
                await Ready(); _dialogDriver = async editor =>
                { await Ready(); var fields = (FieldsDialog)editor; ((TextBox)fields.Input("width")).Text = "240"; ((TextBox)fields.Input("height")).Text = "120"; Click(editor, "确定"); };
                Click(chooser, "确定");
            };
            await ImagesAsync(); var sheet = Document.Snapshot.Sheets[SheetIndex];
            check(sheet.Images.Length == 1 && sheet.Images[0].Anchor.Bounds(sheet).Width == 240, "Native image chooser and anchor dialog preserve decoded image and dimensions");
            await History(false); check(Document.Snapshot.Sheets[SheetIndex].Images.IsEmpty, "Image insertion is one undo entry"); await History(true);
            check(_grid.CellAnnotation(new(4, 0))?.Contains("原生批注") == true, "Inert hover information exposes preserved note text");
            var peer = ControlAutomationPeer.CreatePeerForElement(_grid); var value = (IValueProvider)peer;
            check(peer.GetChildren().Count == 3 && peer.GetName()!.Contains("第 9 行"), "Grid accessibility exposes current row and column with only three children");
            _grid.Select(new(18, 12)); var revision = Document.Snapshot.Revision; value.SetValue("无障碍草稿");
            check(_grid.IsEditing && Document.Snapshot.Revision == revision, "Accessibility value setter owns a native draft without bypassing commands");
            check(await _grid.CommitAsync() && Document.GetCellInput(SheetIndex, 18, 12) == "无障碍草稿", "Accessibility draft commits through shared editor"); await History(false);
            var id = Document.Snapshot.ActiveSheetId;
            _grid.Select(new(17, 12)); _grid.BeginEdit(); _grid.CellEditor.Text = "缩放中的草稿"; var originalInput = Document.GetCellInput(SheetIndex, 17, 12);
            await ZoomAsync(150, preserveDraft: true);
            check(_grid.IsEditing && _grid.CellEditor.Text == "缩放中的草稿" && Document.GetCellInput(SheetIndex, 17, 12) == originalInput, "Wheel zoom preserves original-cell draft without submitting text");
            _grid.Cancel(); await History(false);
            _grid.Select(new(18, 12));
            await ExecuteAsync(new SetCellsCommand([new(new(id, 18, 12), CellContent.FromRichText(new([new TextRun("富文本")])))]));
            check(!_grid.BeginEdit() && _formula.IsReadOnly && value.IsReadOnly, "Rich text is read-only in cell, formula bar and accessibility until explicit flattening");
            await History(false);
            await CalculateAsync(); var original = Path.Combine(directory, "object-source.xlsq"); await _workspace.SaveAsync(new(original));
            _workspace.Editor.SetCell(SheetIndex, 15, 12, "待保存"); Synchronize(); var beforeExport = Document.Snapshot.Revision;
            files.NextSave = Path.Combine(directory, "values-export.xlsx");
            _dialogDriver = async confirm => { await Ready(); Click(confirm, "确定"); }; await ExportAsync(false);
            check(File.Exists(files.NextSave) && _workspace.FilePath == original && Document.IsDirty && Document.Snapshot.Revision == beforeExport, "Native export confirmation writes a copy without saving or rebinding the source");
            var exported = new TypedWorkbookStorage().Open(files.NextSave);
            check(exported.Sheets[0].Images.Length == 1 && exported.Sheets[0].Data.Charts.Length == 1 && exported.Sheets[0].Annotations.Notes.Count == 1, "Export copy preserves image, chart and note metadata");
            _dialogDriver = async report => { await Ready(); Click(report, "复制报告"); Click(report, "关闭"); }; await CompatibilityAsync();
            check(!string.IsNullOrWhiteSpace(((MemoryClipboard)_clipboard).Value.Text), "Native compatibility report copies full grouped diagnostics");
            files.NextOpen = original; _discardChoice = () => Task.FromResult(SaveDecision.Discard); await RecoveryAsync();
            check(_workspace.FilePath is null && _workspace.Version is null && !Document.ScriptsEnabled && Document.Snapshot.Sheets[0].Images.Length == 1, "Native recovery opens an unbound unauthorized copy with attachments intact");
            _grid.SelectRange(new(0, 0, WorkbookLimits.MaxRows - 1, WorkbookLimits.MaxColumns - 1), new(0, 0));
            check(peer.GetChildren().Count == 3, "Select-all never creates unbounded accessibility children");
        }
        finally { _files = previousFiles; _discardChoice = null; _dialogDriver = null; }
    }
}
