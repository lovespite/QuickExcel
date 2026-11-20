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
  type Main = (workbook: Workbook, host: Host) => void | Promise<unknown>;

  /** String storage; quotas count UTF-8 key + value bytes. Entries use ordinal key order.
   * Available inside initialization, actions, events and console transactions. */
  interface IStorage {
    get(key: string): string | undefined;
    set(key: string, value: string): void;
    clear(): void;
    remove(key: string): void;
    keys(): string[];
    values(): string[];
    entries(): Array<[string, string]>;
  }

  interface Host {
    /** Declare exact domains or IPv4/IPv6 addresses at the top level of main.js/main.ts.
     * The host asks for consent before initialization. Only Actions may call fetch.
     * Turning off "Allow network access" in Settings rejects permission requests and revokes prior grants. */
    readonly permissions: { readonly network: { request(host: string): void } };
    /** Saved in the XLSQ workbook. 1 MiB per entry, no storage total quota. */
    readonly localStorage: IStorage;
    /** Session memory only. 1 MiB per entry, 100 MiB total. */
    readonly sessionStorage: IStorage;
    readonly values: ValueUtilities;
    /** Available only inside an action, event or console transaction. */
    readonly files: VirtualFiles;
  }

  /** Legacy API v3 compatibility. API v4 has no registration methods. */
  interface LegacyRegistration {
    readonly events: Events;
    readonly functions: Functions;
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

/** Alias for the script storage contract. */
type IStorage = QuickExcel.IStorage;
declare const localStorage: IStorage;
declare const sessionStorage: IStorage;

/** Bounded host-brokered HTTP(S); available only during Actions after network consent.
 * Settings control each body limit (default 8 MiB, range 1–10 MiB) and the Action deadline
 * including network waits (default 10 seconds, range 1–300 seconds).
 * The 16 MiB cumulative IPC budget and 1-second native JavaScript turn limit still apply. */
declare function fetch(input: string | Request, init?: RequestInit): Promise<Response>;
type HeadersInit = Headers | Record<string, string> | Iterable<readonly [string, string]>;
type BodyInit = string | ArrayBuffer | ArrayBufferView | Blob | URLSearchParams | FormData;
interface RequestInit {
  method?: string;
  headers?: HeadersInit;
  body?: BodyInit | null;
  redirect?: 'follow' | 'error' | 'manual';
  credentials?: 'omit' | 'same-origin';
  signal?: AbortSignal | null;
  mode?: 'cors';
  cache?: 'default' | 'no-store' | 'no-cache' | 'reload';
}
declare class Headers implements Iterable<[string, string]> {
  constructor(init?: HeadersInit);
  append(name: string, value: string): void;
  delete(name: string): void;
  get(name: string): string | null;
  has(name: string): boolean;
  set(name: string, value: string): void;
  entries(): IterableIterator<[string, string]>;
  keys(): IterableIterator<string>;
  values(): IterableIterator<string>;
  forEach(callback: (value: string, key: string, parent: Headers) => void, thisArg?: unknown): void;
  [Symbol.iterator](): IterableIterator<[string, string]>;
}
interface Body {
  readonly bodyUsed: boolean;
  text(): Promise<string>;
  json(): Promise<any>;
  arrayBuffer(): Promise<ArrayBuffer>;
  blob(): Promise<Blob>;
  formData(): Promise<FormData>;
}
/** Buffered body reader; responses are fully received within the host-configured body limit first (default 8 MiB). */
interface FetchBodyReader {
  read(): Promise<{ done: boolean; value?: Uint8Array }>;
  cancel(): Promise<void>;
  releaseLock(): void;
}
declare class Request implements Body {
  constructor(input: string | Request, init?: RequestInit);
  readonly url: string;
  readonly method: string;
  readonly headers: Headers;
  readonly redirect: 'follow' | 'error' | 'manual';
  readonly credentials: 'omit' | 'same-origin';
  readonly signal: AbortSignal;
  readonly bodyUsed: boolean;
  clone(): Request;
  text(): Promise<string>;
  json(): Promise<any>;
  arrayBuffer(): Promise<ArrayBuffer>;
  blob(): Promise<Blob>;
  formData(): Promise<FormData>;
}
interface ResponseInit { status?: number; statusText?: string; headers?: HeadersInit; }
declare class Response implements Body {
  constructor(body?: BodyInit | null, init?: ResponseInit);
  readonly url: string;
  readonly status: number;
  readonly statusText: string;
  readonly ok: boolean;
  readonly headers: Headers;
  readonly redirected: boolean;
  readonly type: string;
  readonly bodyUsed: boolean;
  readonly body: { readonly locked: boolean; getReader(): FetchBodyReader } | null;
  clone(): Response;
  text(): Promise<string>;
  json(): Promise<any>;
  arrayBuffer(): Promise<ArrayBuffer>;
  bytes(): Promise<Uint8Array>;
  blob(): Promise<Blob>;
  formData(): Promise<FormData>;
  static error(): Response;
  static redirect(url: string, status?: number): Response;
  static json(data: unknown, init?: ResponseInit): Response;
}
declare class URLSearchParams implements Iterable<[string, string]> {
  constructor(init?: string | Record<string, string> | Iterable<readonly [string, string]>);
  readonly size: number;
  append(name: string, value: string): void;
  delete(name: string): void;
  get(name: string): string | null;
  getAll(name: string): string[];
  has(name: string): boolean;
  set(name: string, value: string): void;
  sort(): void;
  toString(): string;
  entries(): IterableIterator<[string, string]>;
  keys(): IterableIterator<string>;
  values(): IterableIterator<string>;
  forEach(callback: (value: string, key: string, parent: URLSearchParams) => void, thisArg?: unknown): void;
  [Symbol.iterator](): IterableIterator<[string, string]>;
}
declare class Blob {
  constructor(parts?: Iterable<string | ArrayBuffer | ArrayBufferView | Blob>, options?: { type?: string });
  readonly size: number;
  readonly type: string;
  text(): Promise<string>;
  arrayBuffer(): Promise<ArrayBuffer>;
  slice(start?: number, end?: number, contentType?: string): Blob;
}
declare class FormData implements Iterable<[string, string | Blob]> {
  constructor();
  append(name: string, value: string | Blob, filename?: string): void;
  delete(name: string): void;
  get(name: string): string | Blob | null;
  getAll(name: string): Array<string | Blob>;
  has(name: string): boolean;
  set(name: string, value: string | Blob, filename?: string): void;
  entries(): IterableIterator<[string, string | Blob]>;
  keys(): IterableIterator<string>;
  values(): IterableIterator<string | Blob>;
  forEach(callback: (value: string | Blob, key: string, parent: FormData) => void, thisArg?: unknown): void;
  [Symbol.iterator](): IterableIterator<[string, string | Blob]>;
}
declare class AbortSignal {
  readonly aborted: boolean;
  readonly reason: unknown;
  throwIfAborted(): void;
  addEventListener(type: 'abort', callback: () => void): void;
  removeEventListener(type: 'abort', callback: () => void): void;
  static abort(reason?: unknown): AbortSignal;
}
declare class AbortController { readonly signal: AbortSignal; abort(reason?: unknown): void; }
declare class TextEncoder { readonly encoding: 'utf-8'; encode(input?: string): Uint8Array; }
declare class TextDecoder {
  constructor(label?: string, options?: { fatal?: boolean; ignoreBOM?: boolean });
  readonly encoding: 'utf-8';
  decode(input?: ArrayBuffer | ArrayBufferView): string;
}
