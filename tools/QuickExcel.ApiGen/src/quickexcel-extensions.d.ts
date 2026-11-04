// QuickExcel host extensions. Hand-written; not part of the Office Scripts API.
// Appended after the generated QuickExcel namespace declarations.
declare namespace QuickExcel {
  type ScriptValue = string | number | boolean | null;

  type EventName =
    | 'workbook.open' | 'workbook.beforeClose' | 'workbook.close'
    | 'workbook.activate' | 'workbook.deactivate'
    | 'workbook.beforeSave' | 'workbook.saved' | 'workbook.calculated'
    | 'sheet.added' | 'sheet.deleted' | 'sheet.renamed'
    | 'sheet.activate' | 'sheet.deactivate'
    | 'selection.changed' | 'cell.click' | 'cell.edit' | 'cell.commit' | 'cells.changed';

  type ScriptErrorCode =
    | 'UnsupportedApi' | 'InvalidArgument' | 'ItemNotFound' | 'InvalidReference'
    | 'BudgetExceeded' | 'Locked' | 'TransactionAborted';

  /** Every error thrown by the QuickExcel object model carries a stable code. */
  interface ScriptError extends Error {
    readonly code: ScriptErrorCode;
    readonly member?: string;
  }

  /** The entry point of an action module. One module can declare one action. */
  type Main = (workbook: Workbook, host: Host) => void;

  interface Host {
    readonly events: Events;
    readonly functions: Functions;
    readonly values: ValueUtilities;
    /** Available only inside an action, event or console transaction. */
    readonly files: VirtualFiles;
  }

  interface ValueUtilities {
    /** Returns the top-left cell's current error string, or null for a non-error value. */
    getErrorCode(range: Range): string | null;
  }

  interface Functions {
    /** Registers a worksheet function during module initialization. Value-only interoperability. */
    register(name: string, handler: (...args: ScriptValue[]) => ScriptValue): void;
    /** Descriptor form preserves blank arguments, rectangular ranges and the JS namespace. */
    register(descriptor: FunctionDescriptor, handler: (...args: unknown[]) => ScriptValue): void;
  }

  interface FunctionDescriptor {
    readonly name: string;
    readonly apiVersion: 2;
    readonly parameters?: readonly { readonly name: string; readonly type?: 'any' | 'number' | 'text' | 'boolean' | 'blank' | 'range' }[];
    readonly returns?: 'any' | 'number' | 'text' | 'boolean' | 'blank';
    readonly profile?: 'conservative';
  }

  interface Events {
    /** Registers a handler during module initialization. Handlers run synchronously. */
    on(name: EventName, handler: (context: EventContext) => void): void;
  }

  interface Event {
    readonly name: EventName;
    readonly sessionId: string;
    readonly sequence: number;
    readonly revision: number;
    readonly sheetId: string | null;
    readonly sheetName: string | null;
    readonly ranges: readonly string[];
    readonly isUndoOrRedo: boolean;
  }

  interface ActionContext {
    readonly workbook: Workbook;
    readonly host: Host;
    readonly event: Event | null;
  }

  interface EventContext extends ActionContext {
    readonly event: Event;
  }

  interface VirtualFileInfo {
    readonly path: string;
    readonly type: 'file' | 'directory';
    readonly size: number;
  }

  interface VirtualFiles {
    exists(path: string): boolean;
    stat(path: string): VirtualFileInfo;
    list(directory: string, options?: { recursive?: boolean }): readonly VirtualFileInfo[];
    mkdir(path: string): void;
    remove(path: string): void;
    readBytes(path: string): Uint8Array;
    writeBytes(path: string, bytes: Uint8Array): void;
    readText(path: string): string;
    writeText(path: string, text: string): void;
    readJson(path: string): unknown;
    writeJson(path: string, value: unknown): void;
    copy(source: string, destination: string): void;
    move(source: string, destination: string): void;
  }
}

/** Host extensions during module initialization and every action, event or console request. */
declare const host: QuickExcel.Host;
/** Writes to the script workspace debug console. Output appears after execution. */
declare const console: {
  log(...values: unknown[]): void;
  info(...values: unknown[]): void;
  warn(...values: unknown[]): void;
  error(...values: unknown[]): void;
  debug(...values: unknown[]): void;
};
