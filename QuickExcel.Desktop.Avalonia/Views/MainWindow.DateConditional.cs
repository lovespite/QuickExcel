using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Domain;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private static FieldsDialog DateRuleFields(ConditionalFormatRule old, out Func<ConditionalFormatRule> read)
    {
        static int Tri(bool? value) => value is null ? 0 : value.Value ? 1 : 2;
        var dialog = new FieldsDialog(UiText.Get("发生日期"))
            .Text("ranges", UiText.Get("应用区域（多个区域用空格分隔）"), string.Join(' ', old.Ranges))
            .Choice("period", UiText.Get("日期范围"), DatePeriodLabels, (int)old.DatePeriod!.Value)
            .Text("fill", UiText.Get("填充色（ARGB，留空不改变）"), old.Style.FillColor ?? "")
            .Text("font", UiText.Get("字体色（ARGB，留空不改变）"), old.Style.FontColor ?? "")
            .Choice("bold", UiText.Get("粗体"), [UiText.Get("不改变"), UiText.Get("是"), UiText.Get("否")], Tri(old.Style.Bold))
            .Choice("italic", UiText.Get("斜体"), [UiText.Get("不改变"), UiText.Get("是"), UiText.Get("否")], Tri(old.Style.Italic))
            .Check("stop", UiText.Get("为真时停止后续规则"), old.StopIfTrue);
        read = () =>
        {
            string? Color(string key) => string.IsNullOrWhiteSpace(dialog.String(key)) ? null : dialog.String(key).Trim().TrimStart('#').ToUpperInvariant();
            bool? Flag(string key) => dialog.Index(key) == 0 ? null : dialog.Index(key) == 1;
            var ranges = dialog.String("ranges").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(CellRectangle.Parse);
            return new(ranges, ConditionalFormatKind.TimePeriod, old.Priority, old.Style with { FillColor = Color("fill"), FontColor = Color("font"), Bold = Flag("bold"), Italic = Flag("italic") },
                stopIfTrue: dialog.Boolean("stop"), datePeriod: (ConditionalDatePeriod)dialog.Index("period")) { ScriptId = old.ScriptId };
        };
        return dialog;
    }
    private static readonly string[] DatePeriodLabels = [UiText.Get("今天"), UiText.Get("昨天"), UiText.Get("最近七天"), UiText.Get("本周"), UiText.Get("上周"),
        UiText.Get("上月"), UiText.Get("明天"), UiText.Get("下周"), UiText.Get("下月"), UiText.Get("本月")];
    private void InstallCalendarRefresh()
    {
        var day = DateTime.Today;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(1) };
        timer.Tick += (_, _) =>
        {
            var current = DateTime.Today; if (current == day) return; day = current;
            if (_book.IsVisible) _grid.InvalidatePresentation();
        };
        timer.Start(); Closed += (_, _) => timer.Stop();
    }
    private async Task DateConditionalFormatAsync()
    {
        if (!await FlushEditsAsync()) return; Document.Session.RequireWritable(); RequireConditionalFormatEditing();
        var snapshot = Document.Snapshot; var session = Document.Session.SessionId; var selection = SelectedRange();
        var old = new ConditionalFormatRule([Rectangle(selection)], ConditionalFormatKind.TimePeriod, 1, ConditionalStyles[0].Style, datePeriod: ConditionalDatePeriod.Today);
        var dialog = RuleFields(Rectangle(selection), old, out var read);
        dialog.Title = UiText.Get("发生日期");
        dialog.Note(UiText.Get("按本机当前日期判断，最近七天含今天；每周从星期日开始。数字和明确的年-月-日文本可匹配，模糊的地区日期文本不转换。午夜后一分钟内刷新，不修改工作簿内容。"));
        dialog.Validate = () => read().ValidateBinding(snapshot, snapshot.ActiveSheetId);
        if (!await dialog.ShowAsync(this)) return;
        if (session != Document.Session.SessionId || snapshot.ActiveSheetId != Document.Snapshot.ActiveSheetId || snapshot.Revision != Document.Snapshot.Revision || selection != SelectedRange())
            throw new InvalidOperationException(UiText.Get("文档已改变，请重新预览。"));
        await AddConditionalFormatAsync(snapshot, read());
    }
}
