using QuickExcel.Domain;

namespace QuickExcel.Presentation.Web;

public sealed record ScriptEventTemplate(string Name, string DescriptionZh, string DescriptionEn, string TimingZh, string TimingEn)
{
    public string Handler => "on" + string.Concat(Name.Split('.').Select(p => char.ToUpperInvariant(p[0]) + p[1..]));
    public string Description => UiText.Language == UiText.English ? DescriptionEn : DescriptionZh;
    public string Timing => UiText.Language == UiText.English ? TimingEn : TimingZh;
    public bool Matches(string query) => new[] { Name, Handler, DescriptionZh, DescriptionEn, TimingZh, TimingEn }
        .Any(text => text.Contains(query.Trim(), StringComparison.OrdinalIgnoreCase));
    public string Code(ScriptLanguage language) => "\n" + (language == ScriptLanguage.JavaScript ? "/** @param {QuickExcel.EventContext} context */\n" : "") +
        "export function " + Handler + "(context" + (language == ScriptLanguage.TypeScript ? ": QuickExcel.EventContext): void" : ")") +
        " {\n    const { workbook, host, event } = context;\n    // " + Description + "\n    \n}\n";
}

public static class ScriptEventCatalog
{
    public static IReadOnlyList<ScriptEventTemplate> All { get; } = Array.AsReadOnly(new ScriptEventTemplate[] {
        new("workbook.open", "处理工作簿首次启用脚本。", "Handle the first activation of workbook scripts.", "本次打开工作簿后首次授权脚本时；主入口初始化完成后派发。", "Dispatched after initialization when scripts are first authorized in this open workbook."),
        new("workbook.beforeClose", "在关闭前处理工作簿。", "Handle a workbook before closing.", "用户发起关闭、替换工作簿之前；处理器不能通过返回值取消关闭。", "Before a requested close or workbook replacement; a return value cannot cancel closing."),
        new("workbook.close", "处理工作簿关闭通知。", "Handle the workbook closing notification.", "关闭流程确认后、会话释放之前；只能读取，不能写入工作簿或虚拟文件。", "After closing is confirmed, before session release; workbook and virtual-file writes are prohibited."),
        new("workbook.activate", "处理工作簿窗口激活。", "Handle workbook window activation.", "工作簿窗口获得激活时。", "When the workbook window becomes active."),
        new("workbook.deactivate", "处理工作簿窗口失去激活。", "Handle workbook window deactivation.", "工作簿窗口失去激活时。", "When the workbook window becomes inactive."),
        new("workbook.beforeSave", "在保存前更新数据。", "Update workbook data before saving.", "保存请求通过前置检查后、写入文件之前。", "After save preflight succeeds and before the file is written."),
        new("workbook.saved", "处理保存完成通知。", "Handle successful saves.", "工作簿文件成功写入后；不能修改工作簿或持久文件。", "After a successful file write; workbook and persistent-file edits are prohibited."),
        new("workbook.calculated", "处理正式重算完成通知。", "Handle completed workbook calculations.", "正式计算结果被接受后；事务内临时计算不会触发，且不能修改工作簿或持久文件。", "After a calculation result is accepted; private calculations do not trigger it, and workbook or persistent-file edits are prohibited."),
        new("sheet.added", "处理新增工作表。", "Handle newly added worksheets.", "新增工作表的命令成功提交后，包括撤销或重做。", "After a worksheet addition is committed, including undo or redo."),
        new("sheet.deleted", "处理删除工作表。", "Handle deleted worksheets.", "删除命令成功提交后；事件中保留已删除工作表的 ID 和名称。", "After deletion is committed; the event retains the deleted worksheet ID and name."),
        new("sheet.renamed", "处理工作表重命名。", "Handle worksheet renaming.", "重命名命令成功提交后。", "After a rename command is committed."),
        new("sheet.activate", "处理切换到工作表。", "Handle worksheet activation.", "活动工作表切换后，或首次授权脚本时。", "After the active worksheet changes, or when scripts are first authorized."),
        new("sheet.deactivate", "处理离开工作表。", "Handle worksheet deactivation.", "从原活动工作表切换到另一张表时。", "When switching away from the previously active worksheet."),
        new("selection.changed", "响应选区变化。", "Respond to selection changes.", "会话接受新的活动单元格或选区后。", "After the session accepts a new active cell or selection."),
        new("cell.click", "响应单元格点击。", "Respond to cell clicks.", "网格接收到单元格点击后。", "After the grid receives a cell click."),
        new("cell.edit", "响应进入单元格编辑。", "Respond to entering cell edit mode.", "编辑器开始编辑单元格时；不是每次按键。", "When the editor begins editing a cell, rather than on every keystroke."),
        new("cell.commit", "响应单元格编辑提交。", "Respond to committed cell edits.", "单元格编辑成功提交后。", "After a cell edit is successfully committed."),
        new("cells.changed", "响应单元格内容变化。", "Respond to cell content changes.", "变更命令提交后，包括粘贴、清除、撤销和重做；脚本事件中的写入不递归触发自身。", "After edits are committed, including paste, clear, undo and redo; event-handler writes do not recursively trigger themselves.")
    });
}
