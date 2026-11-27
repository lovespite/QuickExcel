using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using QuickExcel.Core;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private Task<ScriptFileSelection> SelectScriptFilesAsync(DocumentSession session, ScriptFilePickerRequest request,
        Func<IDisposable> beginInteraction, CancellationToken cancellation)
    {
        async Task<ScriptFileSelection> ShowAsync()
        {
            cancellation.ThrowIfCancellationRequested();
            if (!ReferenceEquals(Document.Session, session)) throw new OperationCanceledException(cancellation);
            var owner = _scriptWindow is { IsVisible: true } window ? window : this;
            var provider = owner.StorageProvider;
            if (request.Directory ? !provider.CanPickFolder : request.Save ? !provider.CanSave : !provider.CanOpen)
                throw new NotSupportedException();
            var access = request.Access switch { "r" => UiText.Get("允许读取"), "w" => UiText.Get("允许写入和覆盖"), _ => UiText.Get("允许读取、写入和覆盖") };
            var scope = request.Directory ? UiText.Get("文件夹及其全部子项") : UiText.Get("所选文件");
            var title = $"QuickExcel · {scope} · {access}";
            if (request.AllowDelete) title += " · " + UiText.Get("允许删除子项");
            if (request.Mode == "create") title += " · " + UiText.Get("替换文件内容");
            var workbook = Path.GetFileName(_workspace.FilePath) ?? UiText.Get("未命名.xlsq");
            title += $" · {DisplayName(workbook)} · {DisplayName(request.Action)}";
            if (request.Title is { Length: > 0 }) title += " — " + request.Title;
            var types = request.Filters.Select(f => new FilePickerFileType(f.Name)
                { Patterns = f.Extensions.Select(e => "*." + e).ToArray() }).ToArray();

            IReadOnlyList<IStorageItem> items;
            using (beginInteraction())
            {
                // Win32StorageProvider uses the native Shell FileOpenDialog with FOS_PICKFOLDERS.
                if (request.Directory)
                    items = await provider.OpenFolderPickerAsync(new() { Title = title, AllowMultiple = request.Multiple });
                else if (request.Save)
                {
                    var file = await provider.SaveFilePickerAsync(new() { Title = title, SuggestedFileName = request.SuggestedName,
                        FileTypeChoices = types, ShowOverwritePrompt = true });
                    items = file is null ? [] : [file];
                }
                else items = await provider.OpenFilePickerAsync(new() { Title = title, AllowMultiple = request.Multiple, FileTypeFilter = types });
            }
            try
            {
                cancellation.ThrowIfCancellationRequested();
                if (!ReferenceEquals(Document.Session, session)) throw new OperationCanceledException(cancellation);
                return new(items.Select(i => i.TryGetLocalPath() ?? throw new NotSupportedException()).ToArray(), _workspace.FilePath);
            }
            finally { foreach (var item in items) item.Dispose(); }
        }
        return Dispatcher.UIThread.InvokeAsync(ShowAsync).WaitAsync(cancellation);
    }

    private static string DisplayName(string name) => string.Concat(name.Where(c =>
        !char.IsControl(c) && char.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.Format).Take(160));
}

