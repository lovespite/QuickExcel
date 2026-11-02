import * as monaco from 'monaco-editor/editor/editor.api';
import * as typescript from 'monaco-editor/language/typescript/monaco.contribution';
import 'monaco-editor/languages/definitions/javascript/register';
import declarations from './quickexcel.d.ts';
import { start, post, request, fail, fromBase64, toBase64, hash, Message } from './bridge';

type Model = { model: monaco.editor.ITextModel; version: number; accepted: number; state: monaco.editor.ICodeEditorViewState | null };
const models = new Map<string, Model>();
const workers = new Set<Worker>();
const utf8 = new TextEncoder();
let editor: monaco.editor.IStandaloneCodeEditor;
let selected: string | undefined;
let readOnly = false, locked = true, initialized = false, disposed = false, composing = false;
let scheduled: ReturnType<typeof setTimeout> | undefined;
let sync: Promise<void> = Promise.resolve();
let receive: { module: string; count: number; totalBytes: number; sha256: string; transfer: string; next: number; written: number; bytes: Uint8Array } | undefined;
let totalSourceBytes = 0;
const status = document.getElementById('status')!;
const baseUri = `file:///quickexcel/${new URLSearchParams(location.hash.slice(1)).get('document')}/`;
declare const __QUICKEXCEL_WEB_TEST__: boolean;
(window as any).MonacoEnvironment = { getWorker(_id: string, label: string) {
  const worker = new Worker(new URL(label === 'typescript' || label === 'javascript' ? 'ts.worker.js' : 'editor.worker.js', location.href), { type: 'module' });
  workers.add(worker); return worker;
} };
typescript.javascriptDefaults.setEagerModelSync(true);
typescript.javascriptDefaults.setDiagnosticsOptions({ noSemanticValidation: false, noSyntaxValidation: false });
const extraLib = typescript.javascriptDefaults.addExtraLib(declarations, baseUri + 'quickexcel.d.ts');
editor = monaco.editor.create(document.getElementById('editor')!, { model: null, automaticLayout: true, readOnly: true,
  minimap: { enabled: false }, fontSize: 14, scrollBeyondLastLine: false, accessibilitySupport: 'auto', ariaLabel: 'QuickExcel JavaScript 脚本编辑器',
  links: false, unicodeHighlight: { ambiguousCharacters: false }, stickyScroll: { enabled: false } });
editor.onDidCompositionStart(() => { composing = true; });
editor.onDidCompositionEnd(() => { composing = false; schedule(); });
editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyS, () => post('save', {}));
editor.addCommand(monaco.KeyCode.F6, () => post('focusHost', {}));

function setLocked(value: boolean) { locked = value; editor.updateOptions({ readOnly: readOnly || locked }); }
function select(id: string) {
  const next = models.get(id); if (!next) throw new Error('Unknown module');
  if (selected) models.get(selected)!.state = editor.saveViewState();
  selected = id; editor.setModel(next.model); if (next.state) editor.restoreViewState(next.state); editor.focus();
  status.textContent = `${id} · 静态检查使用 QuickExcel API；运行结果由独立脚本进程提供`;
}
function schedule() {
  if (scheduled !== undefined || composing || locked || disposed) return;
  scheduled = setTimeout(() => { scheduled = undefined; enqueue().catch(fail); }, 0);
}
function enqueue() { sync = sync.then(pushDrafts); return sync; }
async function pushDrafts() {
  for (const [id, item] of models) {
    if (item.version === item.accepted) continue;
    const version = item.version, baseVersion = item.accepted;
    const bytes = utf8.encode(item.model.getValue());
    if (bytes.length > 1048576 || [...models.values()].reduce((sum, m) => sum + utf8.encode(m.model.getValue()).length, 0) > 1048576) throw new Error('脚本总量超过 1 MiB，请恢复后缩减源码');
    const sha256 = await hash(bytes), transfer = crypto.randomUUID(), count = Math.max(1, Math.ceil(bytes.length / 49152));
    for (let index = 0; index < count; index++) {
      const acknowledgment = await request('draftChunk', 'chunkAck', { module: id, baseVersion, version, transfer, index, count, totalBytes: bytes.length, sha256,
        data: toBase64(bytes.subarray(index * 49152, (index + 1) * 49152)) });
      if (acknowledgment.transfer !== transfer || acknowledgment.index !== index || acknowledgment.version !== version) throw new Error('Stale draft acknowledgment');
    }
    item.accepted = version;
  }
  if (!locked && [...models.values()].some(m => m.version !== m.accepted)) schedule();
}
async function loadChunk(payload: any) {
  const { module, count, index, totalBytes, sha256, transfer, data } = payload;
  if (initialized || !models.has(module) || !Number.isInteger(totalBytes) || totalBytes < 0 || totalBytes > 1048576 || count < 1 || count > 32 || data.length > 65536) throw new Error('Invalid source chunk');
  if (!receive) {
    if (index !== 0) throw new Error('Out-of-order source chunk');
    receive = { module, count, totalBytes, sha256, transfer, next: 0, written: 0, bytes: new Uint8Array(totalBytes) };
  }
  const r = receive;
  if (r.module !== module || r.next !== index || r.count !== count || r.totalBytes !== totalBytes || r.sha256 !== sha256 || r.transfer !== transfer) throw new Error('Mismatched source chunk');
  const bytes = fromBase64(data); if (bytes.length > 49152 || r.written + bytes.length > totalBytes) throw new Error('Invalid source size');
  r.bytes.set(bytes, r.written); r.written += bytes.length; r.next++;
  if (r.next === count) {
    if (r.written !== totalBytes || await hash(r.bytes) !== sha256.toLowerCase()) throw new Error('Source checksum mismatch');
    totalSourceBytes += totalBytes; if (totalSourceBytes > 1048576) throw new Error('Source budget exceeded');
    models.get(module)!.model.setValue(new TextDecoder('utf-8', { fatal: true }).decode(r.bytes)); receive = undefined;
  }
}
async function handle(message: Message) {
  const p = message.payload;
  // Eliminated by esbuild in every production resource bundle. No test mutation capability ships.
  if (__QUICKEXCEL_WEB_TEST__ && message.type === 'test') {
    if (p.action === 'inputSample') {
      const started = performance.now(); const model = editor.getModel()!; const end = model.getPositionAt(model.getValueLength()); const before = model.getPositionAt(model.getValueLength() - 1);
      editor.executeEdits('benchmark', [{ range: new monaco.Range(before.lineNumber, before.column, end.lineNumber, end.column), text: p.character }]);
      await enqueue(); await new Promise(requestAnimationFrame);
      post('testResult', { elapsedMs: performance.now() - started }, message.requestId); return;
    }
    if (p.action === 'completionSample') {
      const started = performance.now(); const getWorker = await typescript.getJavaScriptWorker(); const worker = await getWorker(editor.getModel()!.uri);
      const completion = await worker.getCompletionsAtPosition(editor.getModel()!.uri.toString(), editor.getValue().indexOf('quickExcel.functions.') + 'quickExcel.functions.'.length);
      post('testResult', { elapsedMs: performance.now() - started, registered: completion?.entries.some(e => e.name === 'register') }, message.requestId); return;
    }
    if (p.action === 'edit') {
      select(p.module); editor.executeEdits('smoke', [{ range: editor.getModel()!.getFullModelRange(), text: p.text }]); editor.pushUndoStop();
      post('testResult', {}, message.requestId); return;
    }
    if (p.action === 'undo') { editor.trigger('smoke', 'undo', null); post('testResult', {}, message.requestId); return; }
    if (p.action === 'crash') { models.get(selected!)!.version++; post('draftChanged', { module: selected, version: models.get(selected!)!.version }); fail('Injected page failure before final source acknowledgment'); post('testResult', {}, message.requestId); return; }
    if (p.action === 'inspect') {
      const getWorker = await typescript.getJavaScriptWorker(); const worker = await getWorker(editor.getModel()!.uri);
      const diagnostics = await worker.getSemanticDiagnostics(editor.getModel()!.uri.toString());
      const completion = await worker.getCompletionsAtPosition(editor.getModel()!.uri.toString(), p.offset ?? editor.getModel()!.getValueLength());
      post('testResult', { selected, source: editor.getValue(), workers: workers.size, readOnly: editor.getOption(monaco.editor.EditorOption.readOnly),
        diagnostics: diagnostics.map(d => ({ code: d.code, text: typeof d.messageText === 'string' ? d.messageText : d.messageText.messageText })), completions: completion?.entries.map(e => e.name).slice(0, 500) ?? [] }, message.requestId); return;
    }
    throw new Error('Unknown test operation');
  }
  switch (message.type) {
    case 'initialize': {
      if (models.size || !Array.isArray(p.modules) || p.modules.length < 1 || p.modules.length > 64) throw new Error('Invalid model initialization');
      readOnly = p.readOnly === true;
      const paths: Record<string, string[]> = {};
      for (const id of p.modules) {
        if (typeof id !== 'string' || !/^[A-Za-z0-9_.-]{1,128}$/.test(id) || id === '.' || id === '..' || models.has(id)) throw new Error('Invalid module ID');
        const uri = monaco.Uri.parse(baseUri + 'modules/' + id + '.js');
        models.set(id, { model: monaco.editor.createModel('', 'javascript', uri), version: 1, accepted: 1, state: null });
        paths[id] = ['modules/' + id + '.js'];
      }
      typescript.javascriptDefaults.setCompilerOptions({ allowJs: true, checkJs: true, allowNonTsExtensions: true, target: typescript.ScriptTarget.ES2022,
        module: typescript.ModuleKind.ESNext, moduleResolution: typescript.ModuleResolutionKind.NodeJs, lib: ['es2023'], baseUrl: baseUri, paths });
      post('initialized', {}, message.requestId); return;
    }
    case 'loadChunk': await loadChunk(p); post('loadAck', { transfer: p.transfer, index: p.index }, message.requestId); return;
    case 'activate': {
      if (initialized || receive) throw new Error('Incomplete initialization');
      initialized = true;
      for (const [id, item] of models) item.model.onDidChangeContent(() => {
        if (locked || readOnly || disposed) { fail(new Error('Unexpected read-only edit')); return; }
        item.version++; post('draftChanged', { module: id, version: item.version }); schedule();
      });
      select(p.module);
      // A real language worker must be ready before the host enables the page.
      const getWorker = await typescript.getJavaScriptWorker(); const worker = await getWorker(editor.getModel()!.uri);
      await worker.getSyntacticDiagnostics(editor.getModel()!.uri.toString());
      setLocked(false); post('loaded', { modules: models.size, workers: workers.size }, message.requestId); return;
    }
    case 'flushDraft': {
      if (!initialized) throw new Error('Editor has not loaded');
      // Composition belongs to the original model; do not end it or run incomplete text.
      if (composing) { post('draftAck', { busy: '中文组合输入尚未结束，请完成输入后重试。' }, message.requestId); return; }
      setLocked(true); if (scheduled !== undefined) { clearTimeout(scheduled); scheduled = undefined; }
      await enqueue(); post('draftAck', { versions: Object.fromEntries([...models].map(([id, m]) => [id, m.accepted])) }, message.requestId); return;
    }
    case 'releaseDraft': setLocked(false); post('released', {}, message.requestId); return;
    case 'select': if (!initialized || composing) throw new Error('Cannot switch during composition'); select(p.module); post('selected', {}, message.requestId); return;
    case 'locate': {
      const item = models.get(p.module); if (!item || item.accepted !== p.version || item.version !== p.version) throw new Error('Diagnostic belongs to an older source');
      select(p.module); const line = Math.max(1, Math.min(item.model.getLineCount(), Number(p.line) || 1));
      editor.setPosition({ lineNumber: line, column: Math.max(1, Number(p.column) || 1) }); editor.revealLineInCenter(line); post('located', {}, message.requestId); return;
    }
    default: throw new Error('Unsupported script page message');
  }
}
async function dispose() {
  if (disposed) return; disposed = true; if (scheduled !== undefined) clearTimeout(scheduled);
  editor.dispose(); for (const item of models.values()) item.model.dispose(); models.clear(); extraLib.dispose();
  for (const worker of workers) worker.terminate(); workers.clear();
}
start(handle, dispose);
