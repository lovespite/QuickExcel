using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

internal sealed class SampleWorkbookDialog : Window
{
    public SampleWorkbookDialog(IReadOnlyList<WorkbookSample> samples)
    {
        Title=UiText.Get("打开示例工作簿");
        Width=730; Height=620; MinWidth=520; MinHeight=360;
        WindowStartupLocation=WindowStartupLocation.CenterOwner;
        ShowInTaskbar=false;
        var root=new global::Avalonia.Controls.Grid { RowDefinitions=new("Auto,Auto,*,Auto,Auto"), Margin=new(22) };
        root.Children.Add(new TextBlock { Text=UiText.Get("选择一个示例，查看输入、脚本与预期结果。"), FontSize=18, FontWeight=FontWeight.SemiBold, TextWrapping=TextWrapping.Wrap, Margin=new(0,0,0,14) });
        var search=new TextBox { Name="SampleSearch", PlaceholderText=UiText.Get("搜索库名或示例内容"), Margin=new(0,0,0,12) };
        AutomationProperties.SetName(search,UiText.Get("搜索示例工作簿"));
        global::Avalonia.Controls.Grid.SetRow(search,1); root.Children.Add(search);
        var list=new ListBox { Name="SampleWorkbooks" };
        list.Styles.Add(new Style(s=>s.OfType<ListBoxItem>()) { Setters={new Setter(ListBoxItem.HorizontalContentAlignmentProperty,HorizontalAlignment.Stretch)} });
        AutomationProperties.SetName(list,UiText.Get("示例工作簿"));
        ScrollViewer.SetHorizontalScrollBarVisibility(list,global::Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled);
        ScrollViewer.SetVerticalScrollBarVisibility(list,global::Avalonia.Controls.Primitives.ScrollBarVisibility.Auto);
        list.ItemTemplate=new FuncDataTemplate<WorkbookSample>((sample,_)=>
        {
            if(sample is null)return null;
            var english=UiText.Language==UiText.English;
            var panel=new StackPanel { Spacing=5, Margin=new(8,9), HorizontalAlignment=HorizontalAlignment.Stretch };
            panel.Children.Add(new TextBlock { Text=english?sample.Id+" · "+sample.TitleEn:sample.Title, FontSize=16, FontWeight=FontWeight.SemiBold, TextWrapping=TextWrapping.Wrap });
            panel.Children.Add(new TextBlock { Text=english?sample.DescriptionEn:sample.Description, Foreground=Brushes.Gray, TextWrapping=TextWrapping.Wrap });
            return panel;
        });
        global::Avalonia.Controls.Grid.SetRow(list,2); root.Children.Add(list);
        var empty=new TextBlock { Name="NoSampleMatches", Text=UiText.Get("没有匹配的示例。"), Margin=new(16), IsHitTestVisible=false };
        global::Avalonia.Controls.Grid.SetRow(empty,2); root.Children.Add(empty);
        var hint=new TextBlock { Text=UiText.Get("示例以可编辑副本打开，脚本需手动应用和执行。"), Foreground=Brushes.Gray, TextWrapping=TextWrapping.Wrap, Margin=new(0,12) };
        global::Avalonia.Controls.Grid.SetRow(hint,3); root.Children.Add(hint);
        var buttons=new StackPanel { Orientation=Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Right, Spacing=10 };
        var open=new Button { Name="SampleOpen", Content=UiText.Get("打开示例工作簿"), IsDefault=true };
        var cancel=new Button { Name="SampleCancel", Content=UiText.Get("取消"), IsCancel=true };
        void OpenSelected() { if(list.SelectedItem is WorkbookSample selected)Close(selected); }
        open.Click+=(_,_)=>OpenSelected(); cancel.Click+=(_,_)=>Close((WorkbookSample?)null);
        list.DoubleTapped+=(_,_)=>OpenSelected();
        list.SelectionChanged+=(_,_)=>open.IsEnabled=list.SelectedItem is WorkbookSample;
        void Filter()
        {
            var query=search.Text?.Trim()??"";
            var matches=samples.Where(s=>string.Join(' ',s.Id,s.Title,s.TitleEn,s.Description,s.DescriptionEn).Contains(query,StringComparison.OrdinalIgnoreCase)).ToArray();
            list.ItemsSource=matches; list.SelectedIndex=matches.Length>0?0:-1;
            open.IsEnabled=matches.Length>0; empty.IsVisible=matches.Length==0;
        }
        search.TextChanged+=(_,_)=>Filter(); Filter();
        buttons.Children.Add(open); buttons.Children.Add(cancel);
        global::Avalonia.Controls.Grid.SetRow(buttons,4); root.Children.Add(buttons);
        Content=root;
    }
}
