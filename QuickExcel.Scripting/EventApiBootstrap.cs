namespace QuickExcel.Scripting;

public sealed partial class JscSession
{
    private const string EventApiBootstrap = """
        const eventDefinitions = [
            ['ON_OPEN','workbook.open',['onOpen','onWorkbookOpen']],
            ['ON_BEFORE_CLOSE','workbook.beforeClose',['onBeforeClose','onWorkbookBeforeClose']],
            ['ON_CLOSE','workbook.close',['onClose','onWorkbookClose']],
            ['ON_ACTIVATE','workbook.activate',['onActivate','onWorkbookActivate']],
            ['ON_DEACTIVATE','workbook.deactivate',['onDeactivate','onWorkbookDeactivate']],
            ['ON_BEFORE_SAVE','workbook.beforeSave',['onBeforeSave','onWorkbookBeforeSave']],
            ['ON_SAVED','workbook.saved',['onSaved','onWorkbookSaved']],
            ['ON_CALCULATED','workbook.calculated',['onCalculated','onWorkbookCalculated']],
            ['ON_WORKSHEET_ADDED','sheet.added',['onWorksheetAdded']],
            ['ON_WORKSHEET_DELETED','sheet.deleted',['onWorksheetDeleted']],
            ['ON_WORKSHEET_RENAMED','sheet.renamed',['onWorksheetRenamed']],
            ['ON_WORKSHEET_ACTIVATE','sheet.activate',['onWorksheetActivate']],
            ['ON_WORKSHEET_DEACTIVATE','sheet.deactivate',['onWorksheetDeactivate']],
            ['ON_SELECTION_CHANGED','selection.changed',['onSelectionChanged']],
            ['ON_CELL_CLICK','cell.click',['onCellClick']],
            ['ON_CELL_EDIT','cell.edit',['onCellEdit']],
            ['ON_CELL_COMMIT','cell.commit',['onCellCommit']],
            ['ON_CELLS_CHANGED','cells.changed',['onCellsChanged']]
        ];
        const eventNames = eventDefinitions.map(definition => definition[1]);
        const eventConstants = freeze(Object.fromEntries(eventDefinitions.map(([key,name]) => [key,name])));
        Object.defineProperty(globalThis, 'QxEventNames', { value: eventConstants });
        function registerEvent(name, handler) {
            if (!actionMode || sharedMode && !(registrationRole() & 2)) return;
            if (loaded || !eventNames.includes(name) || typeof handler !== 'function') throw new TypeErrorCtor('Invalid event registration.');
            const key = '@event/' + name;
            if (!eventHandlers[name]) {
                if (ownKeys(actions).length >= 256) throw new TypeErrorCtor('Too many registrations.');
                eventHandlers[name] = [];
                actions[key] = context => { for (const fn of eventHandlers[name]) {
                    const result = fn(context);
                    if (result !== undefined && result !== null) throw new TypeErrorCtor('Event handlers must complete synchronously.');
                }};
                actionInfo[key] = freeze({name:key, description:'Event: ' + name});
            }
            if (eventHandlers[name].length >= 32) throw new TypeErrorCtor('Too many handlers.');
            eventHandlers[name].push(handler);
        }
        function createEvents() {
            const events = { on: registerEvent };
            for (const [, name, properties] of eventDefinitions) {
                let handler = null;
                const listener = context => {
                    if (name.startsWith('sheet.') && (typeof context.event.sheetId !== 'string' || typeof context.event.sheetName !== 'string'))
                        throw new TypeErrorCtor('Worksheet event requires an id and name.');
                    const argument = name.startsWith('sheet.')
                        ? freeze({id:context.event.sheetId, name:context.event.sheetName}) : context.event;
                    return handler(argument, context);
                };
                for (const property of properties) Object.defineProperty(events, property, {
                    enumerable: true,
                    get() { return handler; },
                    set(value) {
                        if (!actionMode || sharedMode && !(registrationRole() & 2)) return;
                        if (loaded || value !== null && typeof value !== 'function') throw new TypeErrorCtor('Event handlers must be functions or null and set during initialization.');
                        if (value === null && handler !== null) {
                            const list = eventHandlers[name]; list.splice(list.indexOf(listener), 1);
                            if (!list.length) { delete eventHandlers[name]; delete actions['@event/' + name]; delete actionInfo['@event/' + name]; }
                        } else if (value !== null && handler === null) registerEvent(name, listener);
                        handler = value;
                    }
                });
            }
            return freeze(events);
        }
        """;
}
