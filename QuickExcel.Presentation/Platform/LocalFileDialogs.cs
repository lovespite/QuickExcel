namespace QuickExcel.Presentation;

/// <summary>A selected local file, not an arbitrary storage-provider URI.</summary>
public sealed record LocalFileSelection
{
    public string Path { get; }

    public LocalFileSelection(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !System.IO.Path.IsPathFullyQualified(path))
            throw new ArgumentException(UiText.Get("当前存储器需要完整的本地文件路径。"), nameof(path));
        Path = System.IO.Path.GetFullPath(path);
    }

    public static LocalFileSelection FromUri(Uri uri)
    {
        ArgumentNullException.ThrowIfNull(uri);
        if (!uri.IsAbsoluteUri || !uri.IsFile)
            throw new NotSupportedException(UiText.Get("此位置不提供本地文件路径，请选择本地工作簿。"));
        return new(uri.LocalPath);
    }
}

public sealed record FileDialogType(string Name, IReadOnlyList<string> Patterns);
public sealed record FileDialogRequest(string Title, IReadOnlyList<FileDialogType> Types,
    string? SuggestedName = null, string? DefaultExtension = null, int SelectedType = 0);

/// <summary>Cancellation returns null. Providers without local paths must report an explicit error.</summary>
public interface ILocalFileDialogs
{
    Task<LocalFileSelection?> OpenAsync(FileDialogRequest request);
    Task<LocalFileSelection?> SaveAsync(FileDialogRequest request);
}
