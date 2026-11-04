/** QuickExcel native bridge contract. No DOM, Node, OS filesystem, network or asynchronous host API. Virtual files are session scoped. */
type QxErrorCode = 'Value' | 'Name' | 'Reference' | 'DivisionByZero' | 'Number' | 'NotAvailable' | 'Null' | 'Script' | 'Cycle' | 'Calculation';
type QxScalar = number | string | boolean | null | { readonly kind: 'error'; readonly code: QxErrorCode };
interface QxFunctionDescriptor {
  name: string;
  apiVersion: 2;
  parameters?: readonly { name: string; type?: 'any' | 'number' | 'text' | 'boolean' | 'blank' | 'range' }[];
  returns?: 'any' | 'number' | 'text' | 'boolean' | 'blank';
  profile?: 'conservative';
}
type QxEventName = 'workbook.open' | 'workbook.beforeClose' | 'workbook.close' | 'workbook.activate' | 'workbook.deactivate'
  | 'workbook.beforeSave' | 'workbook.saved' | 'workbook.calculated' | 'sheet.added' | 'sheet.deleted' | 'sheet.renamed'
  | 'sheet.activate' | 'sheet.deactivate' | 'selection.changed' | 'cell.click' | 'cell.edit' | 'cell.commit' | 'cells.changed';
/** Supported event names. ON_BEFORE_SAVE fires before saving; ON_SAVED fires after a successful save. */
declare enum QxEventNames {
  ON_OPEN = 'workbook.open', ON_BEFORE_CLOSE = 'workbook.beforeClose', ON_CLOSE = 'workbook.close',
  ON_ACTIVATE = 'workbook.activate', ON_DEACTIVATE = 'workbook.deactivate',
  ON_BEFORE_SAVE = 'workbook.beforeSave', ON_SAVED = 'workbook.saved',
  ON_CALCULATED = 'workbook.calculated', ON_WORKSHEET_ADDED = 'sheet.added', ON_WORKSHEET_DELETED = 'sheet.deleted',
  ON_WORKSHEET_RENAMED = 'sheet.renamed', ON_WORKSHEET_ACTIVATE = 'sheet.activate', ON_WORKSHEET_DEACTIVATE = 'sheet.deactivate',
  ON_SELECTION_CHANGED = 'selection.changed', ON_CELL_CLICK = 'cell.click', ON_CELL_EDIT = 'cell.edit',
  ON_CELL_COMMIT = 'cell.commit', ON_CELLS_CHANGED = 'cells.changed'
}
interface QxEvent {
  readonly name: QxEventName;
  readonly sessionId: string;
  readonly sequence: number;
  /** Revision at notification time. snapshot.baseRevision is the state used by this handler. */
  readonly revision: number;
  readonly sheetId: string | null;
  readonly sheetName: string | null;
  /** A1 rectangles. cells.changed carries a bounding rectangle per sheet. */
  readonly ranges: readonly string[];
  readonly isUndoOrRedo: boolean;
}
interface QxFileInfo { readonly path: string; readonly type: 'file' | 'directory'; readonly size: number }
/** Worksheet identity at notification time. Deletion events retain the original id and name. */
interface QxEventWorksheet { readonly id: string; readonly name: string }
type QxEventContext = QxActionContext & { readonly event: QxEvent };
type QxEventHandler = ((event: QxEvent, context: QxEventContext) => void | null) | null;
type QxWorksheetEventHandler = ((sheet: QxEventWorksheet, context: QxEventContext) => void | null) | null;
/** Register during initialization. Assigning a property replaces its handler; null removes it. Property handlers coexist with on() listeners. */
interface QxEvents {
  on(name: QxEventName, handler: (context: QxEventContext) => void | null): void;
  onOpen: QxEventHandler; onWorkbookOpen: QxEventHandler;
  onBeforeClose: QxEventHandler; onWorkbookBeforeClose: QxEventHandler;
  onClose: QxEventHandler; onWorkbookClose: QxEventHandler;
  onActivate: QxEventHandler; onWorkbookActivate: QxEventHandler;
  onDeactivate: QxEventHandler; onWorkbookDeactivate: QxEventHandler;
  /** Fires before saving. Persistent changes made by the handler are included in that save. */
  onBeforeSave: QxEventHandler; onWorkbookBeforeSave: QxEventHandler;
  onSaved: QxEventHandler; onWorkbookSaved: QxEventHandler;
  onCalculated: QxEventHandler; onWorkbookCalculated: QxEventHandler;
  onWorksheetAdded: QxWorksheetEventHandler; onWorksheetDeleted: QxWorksheetEventHandler;
  onWorksheetRenamed: QxWorksheetEventHandler; onWorksheetActivate: QxWorksheetEventHandler;
  onWorksheetDeactivate: QxWorksheetEventHandler;
  onSelectionChanged: QxEventHandler; onCellClick: QxEventHandler; onCellEdit: QxEventHandler;
  onCellCommit: QxEventHandler; onCellsChanged: QxEventHandler;
}
interface QxVirtualFiles {
  /** Only //xlsq/, //temp/, //mem/. Case sensitive, no relative segments, escapes or OS paths. */
  exists(path: string): boolean;
  stat(path: string): QxFileInfo;
  list(directory: string, options?: { recursive?: boolean }): readonly QxFileInfo[];
  /** Parent directories are implicit; mkdir also preserves empty directories. No trailing slash except volume roots. */
  mkdir(path: string): void;
  /** Directory must be empty. */
  remove(path: string): void;
  /** Reads bytes from a virtual file. Paths are restricted to //xlsq/, //temp/ and //mem/. */
  readBytes(path: string): Uint8Array;
  writeBytes(path: string, bytes: Uint8Array): void;
  /** Strict UTF-8. */
  readText(path: string): string;
  /** Writes UTF-8 text to a virtual file and commits atomically with this operation's cell changes. */
  writeText(path: string, text: string): void;
  readJson(path: string): any;
  writeJson(path: string, value: unknown): void;
  copy(source: string, destination: string): void;
  move(source: string, destination: string): void;
}
interface QxActionContext {
  /** The current workbook transaction. Reads see earlier writes in this operation; all changes commit together on success. */
  readonly workbook: ExcelScript.Workbook;
  /** A transaction-local view; all writes commit atomically with patch. No OS path mapping. */
  readonly fs: QxVirtualFiles;
  readonly event: QxEvent | null;
  readonly snapshot: {
    readonly baseRevision: number;
    readonly dateSystem: 'Windows1900' | 'Mac1904';
    readonly sheets: readonly { readonly id: string; readonly name: string }[];
    /** Reads a cell from the immutable snapshot. Use an A1 address and a sheetId from snapshot.sheets. */
    getCell(sheetId: string, cell: string): Readonly<{ value: QxScalar; formula: string | null; state: string; hasValue: boolean }>;
    /** Returns row-major values for a rectangle such as A1:C3. Blank cells are null. */
    getRange(sheetId: string, range: string): readonly (readonly QxScalar[])[];
  };
  readonly patch: {
    /** Sets a cell value using an A1 address. Changes commit together on success and can be undone as one operation. */
    setValue(sheetId: string, cell: string, value: QxScalar): void;
    setValues(sheetId: string, range: string, values: readonly (readonly QxScalar[])[]): void;
    /** Sets a formula such as =SUM(A1:A3). This does not directly execute scripts. */
    setFormula(sheetId: string, cell: string, formula: string): void;
    setFormulas(sheetId: string, range: string, formulas: readonly (readonly string[])[]): void;
  };
}
/** The ExcelScript object subset supported by QuickExcel. Office services and network access are not available. */
declare namespace ExcelScript {
  interface RangeFill {
    /** Sets the fill color in #RRGGBB format, for example #ff0000. */
    setColor(color: string): void;
    /** Returns the top-left cell's fill color, or an empty string when it has no fill. */
    getColor(): string;
    clear(): void;
  }
  interface RangeFormat { getFill(): RangeFill }
  interface Range {
    /** Returns the A1 address within the worksheet. */
    getAddress(): string;
    getWorksheet(): Worksheet;
    getFormat(): RangeFormat;
    /** Returns the top-left cell's current value, calculating the transaction draft if needed. Blocked formulas throw; formula errors return an error value with kind and code. */
    getValue(): QxScalar;
    getValues(): QxScalar[][];
    /** Sets every cell in the range to the same value. */
    setValue(value: QxScalar): void;
    setValues(values: readonly (readonly QxScalar[])[]): void;
    getFormula(): string;
    setFormula(formula: string): void;
    setFormulas(formulas: readonly (readonly string[])[]): void;
  }
  interface Worksheet {
    getId(): string;
    getName(): string;
    getCell(address: string): Range;
    /** Row and column indices are zero-based. */
    getCell(row: number, column: number): Range;
    getRange(address: string): Range;
  }
  interface Workbook {
    getActiveCell(): Range;
    getActiveWorksheet(): Worksheet;
    /** Returns the current selection. Throws for multiple areas; use getActiveRanges() for those selections. */
    getActiveRange(): Range;
    getActiveRanges(): Range[];
    getWorksheet(nameOrId: string): Worksheet | undefined;
    getWorksheets(): Worksheet[];
  }
}
declare const quickExcel: ExcelScript.Workbook & {
  readonly functions: {
    /** Registers a worksheet function during module initialization. UDFs run in a persistent context separate from actions/events, cannot access workbook objects, and should depend only on their arguments. This overload receives scalars or flattened ranges. */
    register(name: string, callback: (...args: any[]) => QxScalar): void;
    /** Descriptor API v2 preserves Blank and rectangular ranges; JS namespace required. */
    register(descriptor: QxFunctionDescriptor, callback: (...args: any[]) => QxScalar): void;
  };
  /** File scriptApiVersion 2 and an actions entry binding are required. */
  readonly actions: { register(name: string | { name: string; description?: string }, callback: (context: QxActionContext) => void | null): void };
  /** Uses the same API v2 actions entry binding and trust decision. Register only during module initialization. */
  readonly events: QxEvents;
};
/** Alias for the same quickExcel object, including functions, actions and events. */
declare const $: typeof quickExcel;
/** Writes to the script workspace debug console. Supports multiple arguments and objects; output appears after execution. */
declare const console: {
  /** Writes a log message. */
  log(...values: unknown[]): void;
  /** Writes an informational message. */
  info(...values: unknown[]): void;
  /** Writes a warning, displayed with a yellow background in the console. */
  warn(...values: unknown[]): void;
  /** Writes an error with a red console background. Does not throw or stop execution. */
  error(...values: unknown[]): void;
  /** Writes a debug message that can be shown with the console level filter. */
  debug(...values: unknown[]): void;
};
