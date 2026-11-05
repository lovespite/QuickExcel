using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunCellContentSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? failure = null;
        void Check(bool value, string label) { if (!value) throw new InvalidOperationException(label); checks.Add(label); }
        void Dialog(Action<FieldsDialog> edit, bool accept = true) => _dialogDriver = w => { edit((FieldsDialog)w); w.Close(accept); return Task.CompletedTask; };
        static void Text(FieldsDialog d, string key, string value) => ((TextBox)d.Input(key)).Text = value;
        static void Choose(FieldsDialog d, string key, int value) => ((ComboBox)d.Input(key)).SelectedIndex = value;
        try
        {
            _workspace.New(); ShowBook(); var sheet = Document.Snapshot.ActiveSheetId;
            _workspace.Editor.SetCell(0, 0, 0, "Text"); Synchronize();
            var revision = Document.Snapshot.Revision;
            Dialog(d => Text(d, "target", "https://example.com"), false); await HyperlinkAsync();
            Check(Document.Snapshot.Revision == revision, "Cancel hyperlink preserves document");
            Dialog(d => Text(d, "target", "https://example.com")); await HyperlinkAsync();
            Check(Document.Snapshot.GetSheet(sheet).Annotations.Links[new(0, 0)].Target == "https://example.com" && Document.GetCellValue(0, 0, 0) == "Text", "Hyperlink editor preserves cell text and stores inert target");
            await History(false); Check(Document.Snapshot.GetSheet(sheet).Annotations.Links.IsEmpty, "Hyperlink creation is one undoable edit"); await History(true);
            Dialog(d => { Text(d, "target", "https://changed.example.com"); _workspace.Editor.SetCell(0, 8, 0, "newer"); });
            var stale = false; try { await HyperlinkAsync(); } catch (InvalidOperationException) { stale = true; }
            Check(stale && Document.Snapshot.GetSheet(sheet).Annotations.Links[new(0, 0)].Target == "https://example.com", "Stale hyperlink dialog preserves newer revision");
            Dialog(d => { Text(d, "author", "Test Author"); Text(d, "text", "Root 中文"); }); await ThreadCommentsAsync();
            var root = Document.Snapshot.GetSheet(sheet).Data.ThreadComments.Single();
            Dialog(d => Text(d, "text", "Reply")); await ThreadCommentsAsync();
            var reply = Document.Snapshot.GetSheet(sheet).Data.ThreadComments.Single(c => c.ParentId is not null);
            Check(reply.ParentId == root.Id && reply.Id != root.Id, "Thread reply has independent UUID and points to root");
            Dialog(d => Choose(d, "operation", 1)); await ThreadCommentsAsync();
            Check(Document.Snapshot.GetSheet(sheet).Data.ThreadComments.Single(c => c.Id == root.Id).Resolved, "Desktop resolves discussion without changing identity");
            Dialog(d => { Choose(d, "operation", 3); Choose(d, "selected", 1); }); await ThreadCommentsAsync();
            Check(Document.Snapshot.GetSheet(sheet).Data.ThreadComments.Length == 1, "Desktop deletes selected reply only");
            await History(false); Check(Document.Snapshot.GetSheet(sheet).Data.ThreadComments.Contains(reply), "Undo restores original reply UUID and metadata");
            foreach (var extension in new[] { "xlsx", "xlsq" })
            {
                var file = Path.Combine(directory, "content." + extension); var storage = new TypedWorkbookStorage();
                if (extension == "xlsx") storage.ExportPortable(file, Document.Snapshot); else storage.Save(file, Document.Snapshot);
                var reopened = storage.Open(file).Sheets[0];
                Check(reopened.Data.ThreadComments.Any(c => c.Id == root.Id && c.Resolved) && reopened.Data.ThreadComments.Any(c => c.Id == reply.Id && c.ParentId == root.Id), "Comments and resolution survive " + extension);
            }
            Dialog(d => { Choose(d, "operation", 3); Choose(d, "selected", 0); }); await ThreadCommentsAsync();
            Check(Document.Snapshot.GetSheet(sheet).Data.ThreadComments.IsEmpty, "Deleting root deletes its replies atomically");
        }
        catch (Exception ex) { failure = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new { passed = failure is null, checks, error = failure?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true; (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null ? 0 : 1);
    }
}
