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
interface QxVirtualFiles {
  /** Only //xlsq/, //temp/, //mem/. Case sensitive, no relative segments, escapes or OS paths. */
  exists(path: string): boolean;
  stat(path: string): QxFileInfo;
  list(directory: string, options?: { recursive?: boolean }): readonly QxFileInfo[];
  /** Parent directories are implicit; mkdir also preserves empty directories. No trailing slash except volume roots. */
  mkdir(path: string): void;
  /** Directory must be empty. */
  remove(path: string): void;
  readBytes(path: string): Uint8Array;
  writeBytes(path: string, bytes: Uint8Array): void;
  /** Strict UTF-8. */
  readText(path: string): string;
  writeText(path: string, text: string): void;
  readJson(path: string): any;
  writeJson(path: string, value: unknown): void;
  copy(source: string, destination: string): void;
  move(source: string, destination: string): void;
}
interface QxActionContext {
  /** A transaction-local view; all writes commit atomically with patch. No OS path mapping. */
  readonly fs: QxVirtualFiles;
  readonly event: QxEvent | null;
  readonly snapshot: {
    readonly baseRevision: number;
    readonly dateSystem: 'Windows1900' | 'Mac1904';
    readonly sheets: readonly { readonly id: string; readonly name: string }[];
    getCell(sheetId: string, cell: string): Readonly<{ value: QxScalar; formula: string | null; state: string; hasValue: boolean }>;
    getRange(sheetId: string, range: string): readonly (readonly QxScalar[])[];
  };
  readonly patch: {
    setValue(sheetId: string, cell: string, value: QxScalar): void;
    setValues(sheetId: string, range: string, values: readonly (readonly QxScalar[])[]): void;
    setFormula(sheetId: string, cell: string, formula: string): void;
    setFormulas(sheetId: string, range: string, formulas: readonly (readonly string[])[]): void;
  };
}
declare const quickExcel: {
  readonly functions: {
    /** Legacy scalar / flattened range arguments. Registration is allowed only during module initialization. */
    register(name: string, callback: (...args: any[]) => QxScalar): void;
    /** Descriptor API v2 preserves Blank and rectangular ranges; JS namespace required. */
    register(descriptor: QxFunctionDescriptor, callback: (...args: any[]) => QxScalar): void;
  };
  /** File scriptApiVersion 2 and an actions entry binding are required. */
  readonly actions: { register(name: string | { name: string; description?: string }, callback: (context: QxActionContext) => void | null): void };
  /** Uses the same API v2 actions entry binding and trust decision. Register only during module initialization. */
  readonly events: { on(name: QxEventName, handler: (context: QxActionContext & { readonly event: QxEvent }) => void | null): void };
};
declare const console: { log(...values: unknown[]): void; info(...values: unknown[]): void; warn(...values: unknown[]): void; error(...values: unknown[]): void };
