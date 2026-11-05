using Avalonia.Controls;
using Avalonia.Media;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private readonly TextBlock _pivotStatus = new() { Margin=new(12,5),Foreground=Brushes.DarkOrange };
    private WorksheetPivot? CurrentPivot => Document.Snapshot.Sheets[SheetIndex].Data.Pivots.FirstOrDefault(p=>p.Range.Contains(_grid.Current.Row,_grid.Current.Column));
    private void UpdatePivotStatus()
    {
        var count=Document.Snapshot.Sheets.Sum(s=>s.Data.Pivots.Count(p=>p.Stale));
        var unknown=PivotReport.HasUnmodeledDefinitions(Document.Snapshot);
        _pivotStatus.Text=unknown ? UiText.Get("透视来源结果可能过期，部分定义尚不能刷新") : count>0 ? UiText.Format($"{count} 个透视表待刷新") : "";
        _pivotStatus.IsVisible=count>0 || unknown;
    }

    private async Task CreatePivotAsync()
    {
        if (!await FlushEditsAsync()) return;
        var snapshot=Document.Snapshot; var tables=snapshot.Sheets.SelectMany(s=>s.Data.Tables.Where(t=>t.HeaderRow).Select(t=>(Sheet:s,Table:t))).ToArray();
        if (tables.Length==0) throw new InvalidOperationException(UiText.Get("请先将含标题的数据区域创建为表格，再创建透视表。"));
        var targets=snapshot.Sheets.Where(s=>s.Visibility==SheetVisibility.Visible).ToArray(); var suffix=1;
        while (snapshot.Sheets.SelectMany(s=>s.Data.Pivots).Any(p=>p.Name.Equals("Pivot"+suffix,StringComparison.OrdinalIgnoreCase))) suffix++;
        var dialog=new FieldsDialog(UiText.Get("创建透视表")).Note(UiText.Get("选择本地表格、一个行字段、可选列字段及汇总方式。结果需放在空白区域；源数据修改后可显式刷新。"))
            .Text("name",UiText.Get("透视表名称"),"Pivot"+suffix).Choice("source",UiText.Get("来源表格"),tables.Select(t=>$"{t.Sheet.Name} · {t.Table.Name}").ToArray())
            .Choice("row",UiText.Get("行字段"),tables[0].Table.Columns.Select(c=>c.Name).ToArray())
            .Choice("column",UiText.Get("列字段"),new[]{UiText.Get("无")}.Concat(tables[0].Table.Columns.Select(c=>c.Name)).ToArray())
            .Choice("value",UiText.Get("汇总字段"),tables[0].Table.Columns.Select(c=>c.Name).ToArray(),tables[0].Table.Columns.Length-1)
            .Choice("function",UiText.Get("汇总方式"),[UiText.Get("求和"),UiText.Get("计数"),UiText.Get("平均值"),UiText.Get("最小值"),UiText.Get("最大值")]).Text("caption",UiText.Get("汇总标题"),UiText.Get("合计"))
            .Choice("target",UiText.Get("结果工作表"),targets.Select(s=>s.Name).ToArray(),Array.FindIndex(targets,s=>s.Id==snapshot.ActiveSheetId))
            .Text("anchor",UiText.Get("结果左上角"),QuickExcel.Formulas.FormulaSyntax.ColumnName(_grid.Current.Column)+(_grid.Current.Row+1));
        ((ComboBox)dialog.Input("source")).SelectionChanged+=(_,_)=>
        {
            var names=tables[dialog.Index("source")].Table.Columns.Select(c=>c.Name).ToArray();
            foreach (var key in new[]{"row","value","column"})
            {
                var input=(ComboBox)dialog.Input(key); input.ItemsSource=key=="column" ? new[]{UiText.Get("无")}.Concat(names).ToArray() : names;
                input.SelectedIndex=key=="value" ? names.Length-1 : 0;
            }
        };
        CreatePivotCommand Command()
        {
            var table=tables[dialog.Index("source")]; var cell=CellRectangle.Parse(dialog.String("anchor"));
            if (cell.Count!=1) throw new ArgumentException(UiText.Get("结果位置须为单个单元格，例如 A1。"));
            var column=dialog.Index("column"); if (column-1==dialog.Index("row")) throw new ArgumentException(UiText.Get("行字段与列字段不能相同。"));
            return new(targets[dialog.Index("target")].Id,new(cell.Top,cell.Left),new(table.Sheet.Name,table.Table.Range,table.Table.Name),
                dialog.String("name"),dialog.Index("row"),column==0 ? null : column-1,dialog.Index("value"),(PivotAggregation)dialog.Index("function"),dialog.String("caption"));
        }
        dialog.Validate=()=>{ _=Command(); };
        if (!await dialog.ShowAsync(this)) return;
        await ExecuteAsync(Command(),snapshot.Revision);
        var destination=targets[dialog.Index("target")].Id; var anchor=CellRectangle.Parse(dialog.String("anchor"));
        Document.Session.UpdateNavigation(destination,new(anchor.Top,anchor.Left)); Synchronize();
        _status.Text=UiText.Get("已创建透视表；源数据修改后使用“刷新透视表”。");
    }

    private async Task RefreshPivotAsync()
    {
        if (!await FlushEditsAsync()) return;
        var pivot=CurrentPivot ?? throw new InvalidOperationException(UiText.Get("请先选中要刷新的透视表。"));
        await ExecuteAsync(new RefreshPivotCommand(Document.Snapshot.ActiveSheetId,pivot.Name));
        _status.Text=UiText.Get("已刷新透视表 ")+pivot.Name;
    }

    private async Task ConvertPivotToValuesAsync()
    {
        if (!await FlushEditsAsync()) return;
        var pivot=CurrentPivot ?? throw new InvalidOperationException(UiText.Get("请先选中透视表。")); var revision=Document.Snapshot.Revision; var sheet=Document.Snapshot.ActiveSheetId;
        if (await Dialogs.Confirm(this,UiText.Get("透视表转换为值"),UiText.Get("保留当前显示结果并移除透视定义，此操作可以撤销。")))
            await ExecuteAsync(new RemovePivotCommand(sheet,pivot.Name),revision);
    }
}
