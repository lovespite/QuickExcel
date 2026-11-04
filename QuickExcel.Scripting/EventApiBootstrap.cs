namespace QuickExcel.Scripting;

public sealed partial class JscSession
{
    private const string EventApiBootstrap = """
        const eventNames = [
            'workbook.open', 'workbook.beforeClose', 'workbook.close', 'workbook.activate', 'workbook.deactivate',
            'workbook.beforeSave', 'workbook.saved', 'workbook.calculated',
            'sheet.added', 'sheet.deleted', 'sheet.renamed', 'sheet.activate', 'sheet.deactivate',
            'selection.changed', 'cell.click', 'cell.edit', 'cell.commit', 'cells.changed'
        ];
        function registerEvent(name, handler) {
            if (!actionMode || sharedMode && !(registrationRole() & 2)) return;
            if (loaded) throw new ScriptError('InvalidArgument', '事件处理器只能在脚本加载期间注册。');
            if (!eventNames.includes(name) || typeof handler !== 'function')
                throw new ScriptError('InvalidArgument', '无效的事件注册。');
            const list = eventHandlers[name] ?? (eventHandlers[name] = []);
            if (list.length >= 32 || ownKeys(eventHandlers).length > 256)
                throw new ScriptError('InvalidArgument', '事件处理器数量超过限制。');
            list.push(handler);
        }
        const events = freeze({ on(name, handler) { registerEvent(name, handler); } });
        """;
}
