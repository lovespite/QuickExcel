using System.Text.Json;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private async Task RunSamplesSmokeAsync(string directory)
    {
        Directory.CreateDirectory(directory); var checks=new List<string>(); Exception? failure=null;
        void Check(bool passed,string text) { if(!passed)throw new InvalidOperationException(text); checks.Add(text); }
        static Task Layout()=>Dispatcher.UIThread.InvokeAsync(()=>{},DispatcherPriority.Background).GetTask();
        static void Click(Control control,string name)=>control.GetVisualDescendants().OfType<Button>().Single(b=>b.Name==name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        async Task Capture(Control control,string name)
        {
            await Layout();
            using var bitmap=new RenderTargetBitmap(new PixelSize((int)control.Bounds.Width,(int)control.Bounds.Height),new Vector(96,96));
            bitmap.Render(control);bitmap.Save(Path.Combine(directory,name+".png"),PngBitmapEncoderOptions.Default);
        }
        try
        {
            _workspace.New(); ShowHome(); await Layout();
            var button=_start.GetVisualDescendants().OfType<Button>().Single(b=>b.Name=="OpenSampleWorkbooks");
            Check(button.IsVisible&&button.IsEnabled&&AutomationProperties.GetName(button)==UiText.Get("打开示例工作簿"),"Home exposes Open Sample Workbook");
            await Capture(this,"samples-home");
            var original=Document;
            var pickerExplored=false;
            _dialogDriver=async window=>
            {
                Check(window is SampleWorkbookDialog,"Home button opens the sample picker");
                var list=window.GetVisualDescendants().OfType<ListBox>().Single(l=>l.Name=="SampleWorkbooks");
                var search=window.GetVisualDescendants().OfType<TextBox>().Single(t=>t.Name=="SampleSearch");
                Check(list.Items.Count==9&&list.Items.Cast<WorkbookSample>().Select(s=>s.Id).Distinct().Count()==9,"The picker displays all nine distributed libraries");
                await Capture(window,"samples-list");
                search.Text="LODASH";await Layout();Check(list.Items.Count==1&&((WorkbookSample)list.Items[0]!).Id=="lodash","Sample search ignores letter case");
                search.Text="no-such-example";await Layout();
                Check(list.Items.Count==0&&!window.GetVisualDescendants().OfType<Button>().Single(b=>b.Name=="SampleOpen").IsEnabled,"Empty search disables opening");
                search.Text="";window.Width=520;window.Height=360;await Layout();
                list.ScrollIntoView(list.Items[8]!);await Layout();
                var open=window.GetVisualDescendants().OfType<Button>().Single(b=>b.Name=="SampleOpen");
                Check(open.TranslatePoint(default,window)!.Value.Y+open.Bounds.Height<=window.Bounds.Height,"Open remains visible at the minimum dialog size");
                await Capture(window,"samples-minimum");pickerExplored=true;Click(window,"SampleCancel");
            };
            button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));await _operation;
            Check(pickerExplored,"Search and minimum-size dialog checks completed: "+_status.Text);
            Check(ReferenceEquals(Document,original)&&!Document.IsDirty&&_start.IsVisible,"Cancel leaves the current workbook untouched");
            _workspace.Editor.SetCell(0,0,0,"pending edit");
            _discardChoice=()=>Task.FromResult(SaveDecision.Cancel);
            _dialogDriver=window=>{Click(window,"SampleOpen");return Task.CompletedTask;};
            await OpenSampleWorkbooksAsync();
            Check(ReferenceEquals(Document,original)&&Document.GetCellInput(0,0,0)=="pending edit","Canceling unsaved changes prevents the selected sample from replacing the workbook");
            var catalog=new WorkbookSampleCatalog(Path.Combine(AppContext.BaseDirectory,"samples"));
            var template=Path.Combine(AppContext.BaseDirectory,"samples",catalog.Samples[0].File);
            var templateBytes=File.ReadAllBytes(template);
            _discardChoice=()=>Task.FromResult(SaveDecision.Discard);
            _dialogDriver=window=>{Click(window,"SampleOpen");return Task.CompletedTask;};
            await OpenSampleWorkbooksAsync();await Layout();
            Check(_book.IsVisible&&!_start.IsVisible&&_workspace.FilePath is null&&_workspace.Version is null&&Document.IsDirty,"A selected sample opens as an editable unsaved copy");
            Check(!Document.ScriptsEnabled&&!Document.IsReadOnly&&Document.Snapshot.Sheets.Length==4,"Opening examples does not enable their scripts");
            Check(Document.GetCellInput(0,0,0).StartsWith("Lodash",StringComparison.Ordinal)&&Document.Snapshot.ScriptModules.Modules.ContainsKey("example"),"The sample workbook and embedded source reach the normal workspace");
            await Capture(this,"sample-workbook");
            _files=new SmokeFiles { NextSave=Path.Combine(directory,"my-example.xlsq") };
            Check(await SaveAsync(false)&&_workspace.FilePath==Path.Combine(directory,"my-example.xlsq"),"Saving a sample asks for and binds a new file");
            Check(File.ReadAllBytes(template).SequenceEqual(templateBytes)&&_recentFiles.Paths.All(p=>p!=template),"Opening and saving copies preserves the bundled template and does not add it to recent files");
            _workspace.New();ShowHome();
            var fileMenu=(MenuItem)_menu.Items[0]!;
            Check(fileMenu.Items.OfType<MenuItem>().Single(i=>Equals(i.Header,UiText.Get("打开示例工作簿"))).IsEnabled,"File menu also exposes samples from Home");
        }
        catch(Exception error) { failure=error; }
        finally { _dialogDriver=null; }
        await File.WriteAllTextAsync(Path.Combine(directory,"result.json"),JsonSerializer.Serialize(new { passed=failure is null,checks,error=failure?.ToString() },new JsonSerializerOptions { WriteIndented=true }));
        _allowClose=true;
        (Application.Current!.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)!.Shutdown(failure is null?0:1);
    }
}
