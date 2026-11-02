using System.Windows;
using Microsoft.Win32;

namespace QuickExcel.Desktop;

internal sealed class WpfFileDialogs(Window owner, Func<SaveFileDialog, bool?>? save = null) : ILocalFileDialogs
{
    private static string Filter(FileDialogRequest request) => string.Join("|",
        request.Types.Select(type => type.Name + "|" + string.Join(";", type.Patterns)));

    public Task<LocalFileSelection?> OpenAsync(FileDialogRequest request)
    {
        var dialog = new OpenFileDialog { Title = request.Title, Filter = Filter(request) };
        return Task.FromResult(dialog.ShowDialog(owner) == true ? new LocalFileSelection(dialog.FileName) : null);
    }

    public Task<LocalFileSelection?> SaveAsync(FileDialogRequest request)
    {
        var dialog = new SaveFileDialog { Title = request.Title, Filter = Filter(request),
            FilterIndex = request.SelectedType + 1, FileName = request.SuggestedName ?? "",
            DefaultExt = request.DefaultExtension ?? "", AddExtension = true };
        var accepted = save is null ? dialog.ShowDialog(owner) : save(dialog);
        return Task.FromResult(accepted == true ? new LocalFileSelection(dialog.FileName) : null);
    }
}
