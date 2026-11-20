using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task OpenSampleWorkbooksAsync()
    {
        if(!await CommitEditorsAsync())return;
        var catalog=await Task.Run(()=>new WorkbookSampleCatalog(Path.Combine(AppContext.BaseDirectory,"samples")));
        var selected=await ShowOwnedAsync<WorkbookSample?>(new SampleWorkbookDialog(catalog.Samples));
        if(selected is null)return;
        // Read and validate before replacing the current document. Examples have
        // no file binding or remembered execution authorization in the new copy.
        var snapshot=await Task.Run(()=>catalog.Read(selected.Id));
        if(!await ConfirmDiscardAsync())return;
        _workspace.Import(snapshot);
        ShowBook();
        _status.Text=UiText.Get("已打开示例副本：")+(UiText.Language==UiText.English?selected.TitleEn:selected.Title);
    }
}
