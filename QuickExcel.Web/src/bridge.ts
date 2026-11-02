// A document-bound data channel. It exposes no host object or user-code execution API.
export type Message = { protocol: number; instance: string; document: string; generation: number; type: string; requestId: string; payload: any };
const instance = location.pathname.split('/')[1];
const identity = new URLSearchParams(location.hash.slice(1));
const documentId = identity.get('document');
const generation = Number(identity.get('generation'));
const encoder = new TextEncoder();
const decoder = new TextDecoder('utf-8', { fatal: true });
const waiting = new Map<string, { type: string; resolve: (p: any) => void; reject: (e: Error) => void; timer: ReturnType<typeof setTimeout> }>();
let closed = false;
let handler: (message: Message) => Promise<void> = async () => { throw new Error('Page is not initialized'); };
let release: () => Promise<void> = async () => {};
export function post(type: string, payload: unknown, requestId = crypto.randomUUID()) {
  if (closed) throw new Error('Page is closed');
  const body = JSON.stringify({ protocol: 1, instance, document: documentId, generation, type, requestId, payload });
  if (encoder.encode(body).length > 256 * 1024) throw new Error('Message exceeds 256 KiB');
  (window as any).invokeCSharpAction(body);
}
export function request(type: string, expected: string, payload: unknown): Promise<any> {
  if (waiting.size >= 32) return Promise.reject(new Error('Too many pending requests'));
  const id = crypto.randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => { waiting.delete(id); reject(new Error('Host acknowledgment timed out')); }, 10000);
    waiting.set(id, { type: expected, resolve, reject, timer });
    try { post(type, payload, id); } catch (error) { clearTimeout(timer); waiting.delete(id); reject(error); }
  });
}
export function fail(error: unknown) {
  if (!closed) {
    document.body.setAttribute('data-failed', 'true');
    const banner = document.getElementById('status'); if (banner) banner.textContent = '页面已停止，请在主窗口恢复已确认草稿。';
    try { post('error', { message: String(error).slice(0, 2000) }); } catch {}
  }
}
export function start(receive: typeof handler, dispose: typeof release) {
  handler = receive; release = dispose;
  (window as any).quickExcelReceiveBase64 = (data: string) => {
    void (async () => {
      if (closed || typeof data !== 'string' || data.length > 350000) throw new Error('Invalid message');
      const bytes = fromBase64(data); if (bytes.length > 256 * 1024) throw new Error('Message too large');
      const message = JSON.parse(decoder.decode(bytes)) as Message;
      if (Object.keys(message).length !== 7 || message.protocol !== 1 || message.instance !== instance || message.document !== documentId || message.generation !== generation ||
          !/^[A-Za-z0-9_.-]{1,128}$/.test(message.requestId)) throw new Error('Wrong document or page generation');
      const pending = waiting.get(message.requestId);
      if (pending) {
        if (pending.type !== message.type) throw new Error('Unexpected acknowledgment');
        clearTimeout(pending.timer); waiting.delete(message.requestId); pending.resolve(message.payload); return;
      }
      if (message.type === 'ping') { post('pong', {}, message.requestId); return; }
      if (message.type === 'dispose') { await release(); post('disposed', {}, message.requestId); close(); return; }
      await handler(message);
    })().catch(fail);
  };
  window.addEventListener('error', e => fail(e.message));
  window.addEventListener('unhandledrejection', e => fail(e.reason));
  window.addEventListener('pagehide', () => { void release(); close(); }, { once: true });
  post('ready', { userAgent: navigator.userAgent });
}
function close() { closed = true; for (const p of waiting.values()) { clearTimeout(p.timer); p.reject(new Error('Page closed')); } waiting.clear(); }
export function fromBase64(value: string) { return Uint8Array.from(atob(value), c => c.charCodeAt(0)); }
export function toBase64(value: Uint8Array) { let result = ''; for (let i = 0; i < value.length; i += 8192) result += String.fromCharCode(...value.subarray(i, i + 8192)); return btoa(result); }
export async function hash(bytes: Uint8Array) { return [...new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))].map(b => b.toString(16).padStart(2, '0')).join(''); }
