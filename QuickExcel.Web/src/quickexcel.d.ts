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
/** 现有事件的名称。ON_BEFORE_SAVE 在保存前触发，ON_SAVED 在保存成功后触发。 */
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
/** 通知时的工作表身份；删除事件仍保留删除前的 id/name。 */
interface QxEventWorksheet { readonly id: string; readonly name: string }
type QxEventContext = QxActionContext & { readonly event: QxEvent };
type QxEventHandler = ((event: QxEvent, context: QxEventContext) => void | null) | null;
type QxWorksheetEventHandler = ((sheet: QxEventWorksheet, context: QxEventContext) => void | null) | null;
/** 初始化时注册。属性赋值替换该事件的单个处理器，设为 null 移除；与 on() 监听器共存。 */
interface QxEvents {
  on(name: QxEventName, handler: (context: QxEventContext) => void | null): void;
  onOpen: QxEventHandler; onWorkbookOpen: QxEventHandler;
  onBeforeClose: QxEventHandler; onWorkbookBeforeClose: QxEventHandler;
  onClose: QxEventHandler; onWorkbookClose: QxEventHandler;
  onActivate: QxEventHandler; onWorkbookActivate: QxEventHandler;
  onDeactivate: QxEventHandler; onWorkbookDeactivate: QxEventHandler;
  /** 保存前触发；处理器的持久修改包含在此次保存中。 */
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
  /** 读取虚拟文件的字节内容。路径限于 //xlsq/、//temp/、//mem/。 */
  readBytes(path: string): Uint8Array;
  writeBytes(path: string, bytes: Uint8Array): void;
  /** Strict UTF-8. */
  readText(path: string): string;
  /** 将文本按 UTF-8 写入虚拟文件，与本次单元格修改一起原子提交。 */
  writeText(path: string, text: string): void;
  readJson(path: string): any;
  writeJson(path: string, value: unknown): void;
  copy(source: string, destination: string): void;
  move(source: string, destination: string): void;
}
interface QxActionContext {
  /** 当前执行的工作簿对象；读取能看到本轮先前的修改，成功后整体提交。 */
  readonly workbook: ExcelScript.Workbook;
  /** A transaction-local view; all writes commit atomically with patch. No OS path mapping. */
  readonly fs: QxVirtualFiles;
  readonly event: QxEvent | null;
  readonly snapshot: {
    readonly baseRevision: number;
    readonly dateSystem: 'Windows1900' | 'Mac1904';
    readonly sheets: readonly { readonly id: string; readonly name: string }[];
    /** 读取不可变快照中的单元格；cell 使用 A1 地址。sheetId 来自 snapshot.sheets。 */
    getCell(sheetId: string, cell: string): Readonly<{ value: QxScalar; formula: string | null; state: string; hasValue: boolean }>;
    /** 按行返回矩形区域的值，例如 A1:C3；空白为 null。 */
    getRange(sheetId: string, range: string): readonly (readonly QxScalar[])[];
  };
  readonly patch: {
    /** 设置单元格值；操作成功后统一提交，可整次撤销。cell 使用 A1 地址。 */
    setValue(sheetId: string, cell: string, value: QxScalar): void;
    setValues(sheetId: string, range: string, values: readonly (readonly QxScalar[])[]): void;
    /** 设置单元格公式，例如 =SUM(A1:A3)；不直接执行脚本。 */
    setFormula(sheetId: string, cell: string, formula: string): void;
    setFormulas(sheetId: string, range: string, formulas: readonly (readonly string[])[]): void;
  };
}
/** QuickExcel 支持的 ExcelScript 对象子集；不包含 Office 服务或网络能力。 */
declare namespace ExcelScript {
  interface RangeFill {
    /** 设置 #RRGGBB 填充色，例如 #ff0000。 */
    setColor(color: string): void;
    /** 返回左上角单元格填充色；没有填充时返回空字符串。 */
    getColor(): string;
    clear(): void;
  }
  interface RangeFormat { getFill(): RangeFill }
  interface Range {
    /** 工作表内的 A1 地址。 */
    getAddress(): string;
    getWorksheet(): Worksheet;
    getFormat(): RangeFormat;
    /** 返回左上角当前值；必要时先计算事务草稿。受阻公式会抛错，公式错误返回带 kind/code 的错误值。 */
    getValue(): QxScalar;
    getValues(): QxScalar[][];
    /** 将区域中每个单元格设置为同一个值。 */
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
    /** 行列索引从 0 开始。 */
    getCell(row: number, column: number): Range;
    getRange(address: string): Range;
  }
  interface Workbook {
    getActiveCell(): Range;
    getActiveWorksheet(): Worksheet;
    /** 当前选区；多区域选择时抛错，改用 getActiveRanges()。 */
    getActiveRange(): Range;
    getActiveRanges(): Range[];
    getWorksheet(nameOrId: string): Worksheet | undefined;
    getWorksheets(): Worksheet[];
  }
}
declare const quickExcel: ExcelScript.Workbook & {
  readonly functions: {
    /** 注册工作表函数；仅限模块初始化阶段。UDF 与动作/事件使用独立的持久上下文，不能访问工作簿对象；应只依赖参数。此重载接收标量或展平的区域参数。 */
    register(name: string, callback: (...args: any[]) => QxScalar): void;
    /** Descriptor API v2 preserves Blank and rectangular ranges; JS namespace required. */
    register(descriptor: QxFunctionDescriptor, callback: (...args: any[]) => QxScalar): void;
  };
  /** File scriptApiVersion 2 and an actions entry binding are required. */
  readonly actions: { register(name: string | { name: string; description?: string }, callback: (context: QxActionContext) => void | null): void };
  /** Uses the same API v2 actions entry binding and trust decision. Register only during module initialization. */
  readonly events: QxEvents;
};
/** quickExcel 的简写；与 quickExcel 是同一对象，支持 functions、actions 和 events。 */
declare const $: typeof quickExcel;
/** 输出到脚本工作区的调试控制台；支持多个参数和对象，执行结束后显示。 */
declare const console: {
  /** 输出普通日志。 */
  log(...values: unknown[]): void;
  /** 输出信息。 */
  info(...values: unknown[]): void;
  /** 输出警告，在控制台以黄色背景显示。 */
  warn(...values: unknown[]): void;
  /** 输出错误，在控制台以红色背景显示；不会抛出异常或中止执行。 */
  error(...values: unknown[]): void;
  /** 输出调试信息，可通过级别筛选查看。 */
  debug(...values: unknown[]): void;
};
