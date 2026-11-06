using Avalonia.Controls;
using Avalonia.Platform.Storage;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Presentation;
using QuickExcel.Runtime;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

internal static class DesktopServices
{
    // Linux verifies the actual kernel/cgroup protection on every launch; macOS
    // verifies the Seatbelt profile and its polling limits on every launch.
    internal static bool SupportsScripts =>
        (OperatingSystem.IsWindows() || OperatingSystem.IsLinux()) &&
        System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture == System.Runtime.InteropServices.Architecture.X64 ||
        OperatingSystem.IsMacOS() &&
        System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture == System.Runtime.InteropServices.Architecture.Arm64;
    public static WorkbookServices Create() => new(new TypedWorkbookStorage(),
        new WorkbookCalculator(SupportsScripts ? new ScriptWorkerRuntimeFactory() : new UnavailableScripts()),
        SupportsScripts ? new ScriptWorkerActionExecutor() : null);
    private sealed class UnavailableScripts : IScriptRuntimeFactory
    {
        public IScriptRuntime Create(string source) => throw new PlatformNotSupportedException(UiText.Get("此平台尚未完成独立脚本沙箱，脚本执行不可用。"));
    }
}

internal sealed class WorkbookFileAccess(IWorkbookStorage storage) : IWorkbookFileAccess
{
    public IDisposable? AcquireExclusive(string path, bool allowMissing = false)
    {
        if (!allowMissing && !Path.GetExtension(path).Equals(".xlsq", StringComparison.OrdinalIgnoreCase)) TypedWorkbookStorage.CheckWriteAccess(path);
        return (storage as TypedWorkbookStorage)?.AcquireExclusiveFile(path, allowMissing);
    }
    public IDisposable? AcquireReadOnly(string path) => (storage as TypedWorkbookStorage)?.AcquireReadOnlyFile(path);
    public OpenedWorkbook Open(string path)
    {
        if (Path.GetExtension(path).Equals(".xlsq", StringComparison.OrdinalIgnoreCase))
            return new(storage.Open(path), ReadVersion(path));
        using var lease = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return new(storage.Open(path), FileVersions.Read(lease));
    }
    public FileVersion ReadVersion(string path) => storage is TypedWorkbookStorage typed ? typed.ReadVersion(path) : FileVersions.Read(path);
}

internal sealed class AvaloniaFileDialogs(TopLevel owner) : ILocalFileDialogs
{
    private static FilePickerFileType Type(FileDialogType type) => new(type.Name) { Patterns = type.Patterns };
    public async Task<LocalFileSelection?> OpenAsync(FileDialogRequest request)
    {
        var items = await owner.StorageProvider.OpenFilePickerAsync(new() { Title = request.Title, AllowMultiple = false, FileTypeFilter = request.Types.Select(Type).ToArray() });
        if (items.Count == 0) return null;
        using var file = items[0];
        return LocalFileSelection.FromUri(file.Path);
    }
    public async Task<LocalFileSelection?> SaveAsync(FileDialogRequest request)
    {
        var types = request.Types.Select(Type).ToArray();
        using var file = await owner.StorageProvider.SaveFilePickerAsync(new() { Title = request.Title,
            SuggestedFileName = request.SuggestedName, DefaultExtension = request.DefaultExtension?.TrimStart('.'),
            FileTypeChoices = types, SuggestedFileType = types[Math.Clamp(request.SelectedType, 0, types.Length - 1)], ShowOverwritePrompt = true });
        return file is null ? null : LocalFileSelection.FromUri(file.Path);
    }
}
