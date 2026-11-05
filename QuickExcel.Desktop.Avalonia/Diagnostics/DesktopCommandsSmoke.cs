using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using QuickExcel.Presentation;
using QuickExcel.Security;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunDesktopCommandsSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks = new List<string>(); Exception? error = null;
        void Check(bool condition, string label) { if (!condition) throw new InvalidOperationException(label); checks.Add(label); }
        async Task Press(Key key, KeyModifiers modifiers = KeyModifiers.None)
        { _grid.RaiseEvent(new KeyEventArgs { RoutedEvent = KeyDownEvent, Key = key, KeyModifiers = modifiers, Source = _grid }); await _operation; }
        try
        {
            var files = new SmokeFiles(); _files = files; _workspace.New(); ShowBook(); _grid.Focus();
            foreach (var key in new[] { Key.B, Key.I, Key.U, Key.D5 }) await Press(key, KeyModifiers.Control);
            Check(CurrentStyle.Bold && CurrentStyle.Italic && CurrentStyle.Underline == "single" && CurrentStyle.Strike,
                "Ctrl+B/I/U/5 applies Excel font shortcuts to the cell selection");
            foreach (var key in new[] { Key.B, Key.I, Key.U, Key.D5 }) await Press(key, KeyModifiers.Control);
            Check(!CurrentStyle.Bold && !CurrentStyle.Italic && CurrentStyle.Underline is null or "none" && !CurrentStyle.Strike,
                "Font shortcuts toggle the formats off");
            _grid.BeginEdit(); _grid.CellEditor.Text = "draft";
            await Press(Key.B, KeyModifiers.Control);
            Check(_grid.IsEditing && !CurrentStyle.Bold, "Formatting shortcuts leave an active text draft alone"); _grid.Cancel();
            files.NextSave = Path.Combine(directory, "shortcut-saved.xlsq");
            await Press(Key.F12);
            Check(File.Exists(files.NextSave) && _workspace.FilePath == files.NextSave && !Document.IsDirty, "F12 opens Save As and saves the workbook");
            _workspace.Editor.SetCell(0, 0, 0, "saved by Shift+F12"); Synchronize();
            await Press(Key.F12, KeyModifiers.Shift); Check(!Document.IsDirty, "Shift+F12 saves the current workbook");
            await Press(Key.F11, KeyModifiers.Shift); Check(Document.Snapshot.Sheets.Length == 2, "Shift+F11 inserts a worksheet");
            _discardChoice = () => Task.FromResult(SaveDecision.Discard);
            await Press(Key.F4, KeyModifiers.Control); Check(_start.IsVisible && !_book.IsVisible, "Ctrl+F4 closes the workbook");
            files.NextOpen = files.NextSave;
            await Press(Key.F12, KeyModifiers.Control); Check(_book.IsVisible && _workspace.FilePath == files.NextSave, "Ctrl+F12 opens a workbook from Home");
            _dialogDriver = window => { Check(window is FieldsDialog, "F5 opens the Go To dialog"); window.Close(false); return Task.CompletedTask; };
            await Press(Key.F5); Check(_dialogDriver is null, "F5 reaches the worksheet Go To command");

            var certificate = Path.Combine(directory, "signing-test.pfx"); files.NextSave = certificate;
            const string secret = "local-signing-test-password";
            Task Driver(Window window)
            {
                _dialogDriver = Driver; var dialog = (FieldsDialog)window;
                if (window.Title == UiText.Get("签署全部脚本")) ((ComboBox)dialog.Input("kind")).SelectedIndex = 1;
                else
                {
                    ((TextBox)dialog.Input("password")).Text = secret;
                    ((CheckBox)dialog.Input("rememberFile")).IsChecked = true;
                    ((CheckBox)dialog.Input("rememberPassword")).IsChecked = true;
                }
                window.Close(true); return Task.CompletedTask;
            }
            _dialogDriver = Driver; await SignScriptsAsync();
            Check(Document.Snapshot.ScriptModules.Signature is not null && File.Exists(certificate), "Signing creates a certificate and signs the workbook");
            var remembered = SigningPreferences.Read();
            Check(remembered.KeyFilePath == certificate && remembered.PasswordForFile(certificate) == secret, "Both signing preferences survive reopening the local store");
            Task Reuse(Window window)
            {
                _dialogDriver = Reuse; var dialog = (FieldsDialog)window;
                if (window.Title == UiText.Get("签署全部脚本"))
                    Check(dialog.Index("kind") == 0 && ((ComboBox)dialog.Input("kind")).Items.Count == 3, "Remembered key file is offered without reopening the picker");
                else
                {
                    Check(dialog.String("password") == secret && ((TextBox)dialog.Input("password")).PasswordChar == '●', "Remembered password fills the masked field");
                    ((CheckBox)dialog.Input("rememberFile")).IsChecked = false;
                    ((CheckBox)dialog.Input("rememberPassword")).IsChecked = false;
                }
                window.Close(true); return Task.CompletedTask;
            }
            files.NextOpen = null; _dialogDriver = Reuse; await SignScriptsAsync();
            Check(!SigningPreferences.Read().RememberKeyFile && !SigningPreferences.Read().RememberPassword, "Unchecking both choices clears local signing memory");
            _dialogDriver = null;
            files.NextSave = Path.Combine(directory, "signed-workbook.xlsq"); await SaveAsync(true);
            _workspace.New();
            using var package = System.IO.Compression.ZipFile.OpenRead(files.NextSave);
            Check(package.Entries.All(entry => !entry.FullName.Contains("signing-preferences") && !entry.FullName.EndsWith(".pfx")), "Workbook package contains no private-key or signing-preferences part");
            foreach (var entry in package.Entries)
            {
                using var reader = new StreamReader(entry.Open()); var text = await reader.ReadToEndAsync();
                Check(!text.Contains(secret) && !text.Contains(certificate), "Workbook part excludes remembered password and local key path: " + entry.FullName);
            }
        }
        catch (Exception ex) { error = ex; }
        finally { _dialogDriver = null; }
        await File.WriteAllTextAsync(Path.Combine(directory, "result.json"), JsonSerializer.Serialize(new
        { passed = error is null, checks, error = error?.ToString() }, new JsonSerializerOptions { WriteIndented = true }));
        _allowClose = true;
        (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(error is null ? 0 : 1);
    }
}
