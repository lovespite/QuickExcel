/** QuickExcel data-only additions; no browser, DOM or OS access. */
declare class URL {
  constructor(url: string | URL, base?: string | URL);
  hash: string; host: string; hostname: string; href: string; password: string;
  pathname: string; port: string; protocol: string; search: string; username: string;
  readonly origin: string; readonly searchParams: URLSearchParams;
  toString(): string; toJSON(): string;
  static canParse(url: string | URL, base?: string | URL): boolean;
  static parse(url: string | URL, base?: string | URL): URL | null;
}
interface FilePropertyBag { type?: string; lastModified?: number; }
type BufferSource = ArrayBuffer | ArrayBufferView;
interface File extends Blob {
  readonly name: string; readonly lastModified: number; readonly webkitRelativePath: string;
}
declare var File: { readonly prototype: File; new(parts: (string | Blob | ArrayBuffer | ArrayBufferView)[], name: string, options?: FilePropertyBag): File; };
/** Structural reader accepted by table import functions; there is no global stream constructor. */
interface ReadableStream<R = unknown> {
  readonly locked: boolean;
  getReader(): { read(): Promise<{done: boolean; value?: R}>; cancel(reason?: unknown): Promise<void>; releaseLock(): void; };
}
