import { uiText } from './localization';
import * as monaco from 'monaco-editor/editor/editor.api';
// editor.api is deliberately minimal: language providers alone do not install
// their UI. Keep these registrations explicit for the pinned Monaco package.
// In 0.57 features/suggest/register installs only inline suggestions.
import 'monaco-editor/editor/contrib/suggest/browser/suggestController';
import 'monaco-editor/features/codicon/register';
import 'monaco-editor/features/hover/register';
import 'monaco-editor/features/parameterHints/register';
import 'monaco-editor/features/gotoError/register';
import 'monaco-editor/features/gotoSymbol/register';
import 'monaco-editor/features/contextmenu/register';
import 'monaco-editor/features/find/register';
import * as typescript from 'monaco-editor/language/typescript/monaco.contribution';
import 'monaco-editor/languages/definitions/javascript/register';
import 'monaco-editor/languages/definitions/typescript/register';
import declarations from './quickexcel.d.ts';
import builtinDeclarations from '../../artifacts/script-libraries/declarations.json';
import { start, post, request, fail, fromBase64, toBase64, hash, Message } from './bridge';

type Model = { model: monaco.editor.ITextModel; version: number; accepted: number; state: monaco.editor.ICodeEditorViewState | null };
const models = new Map<string, Model>();
const definitions = new Map<string, { model: monaco.editor.ITextModel; state: monaco.editor.ICodeEditorViewState | null }>();
const workers = new Set<Worker>();
const utf8 = new TextEncoder();
let editor: monaco.editor.IStandaloneCodeEditor;
let selected: string | undefined;
let selectedDefinition: string | undefined, navigating = false;
let readOnly = false, locked = true, initialized = false, disposed = false, composing = false;
let scheduled: ReturnType<typeof setTimeout> | undefined;
let sync: Promise<void> = Promise.resolve();
let receive: { module: string; count: number; totalBytes: number; sha256: string; transfer: string; next: number; written: number; bytes: Uint8Array } | undefined;
let totalSourceBytes = 0;
const status = document.getElementById('status')!;
const assistanceButtons = [...document.querySelectorAll<HTMLButtonElement>('button[data-action]')];
const baseUri = `file:///quickexcel/${new URLSearchParams(location.hash.slice(1)).get('document')}/`;
const declarationUri = monaco.Uri.parse(baseUri + 'quickexcel.d.ts');
const isDefinition = (name: string) => name === 'quickexcel.d.ts' || Object.hasOwn(builtinDeclarations.files, name) && /\.d\.[cm]?ts$/.test(name) || /^lib\.(?:es(?:5|6|20\d\d|next)(?:\.[a-z0-9]+)*|decorators(?:\.legacy)?)\.d\.ts$/.test(name);
declare const __QUICKEXCEL_WEB_TEST__: boolean;
(window as any).MonacoEnvironment = { getWorker(_id: string, label: string) {
  const worker = new Worker(new URL(label === 'typescript' || label === 'javascript' ? 'ts.worker.js' : 'editor.worker.js', location.href), { type: 'module' });
  workers.add(worker); return worker;
} };
for (const defaults of [typescript.javascriptDefaults, typescript.typescriptDefaults]) {
  defaults.setEagerModelSync(true);
  // Revalidate on attachment as well as edits, so a dependency draft change is
  // reflected when returning to its importer. Do not check 64 hidden editors.
  defaults.setDiagnosticsOptions({ noSemanticValidation: false, noSyntaxValidation: false, onlyVisible: true });
}
let extraLibs = [typescript.javascriptDefaults, typescript.typescriptDefaults].map(d => d.addExtraLib(declarations, baseUri + 'quickexcel.d.ts'));
const builtinExtraLibs = [typescript.javascriptDefaults, typescript.typescriptDefaults].flatMap(d =>
  Object.entries(builtinDeclarations.files).map(([name, source]) => d.addExtraLib(source, 'file://' + name)));
editor = monaco.editor.create(document.getElementById('editor')!, { model: null, automaticLayout: true, readOnly: true, theme: 'vs-dark',
  minimap: { enabled: false }, fontSize: 14, scrollBeyondLastLine: false, accessibilitySupport: 'auto', ariaLabel: uiText('QuickExcel JavaScript / TypeScript 脚本编辑器'),
  quickSuggestions: { other: true, comments: false, strings: false }, suggestOnTriggerCharacters: true,
  parameterHints: { enabled: true }, hover: { enabled: true }, renderValidationDecorations: 'on',
  wordBasedSuggestions: 'off',
  gotoLocation: { multipleDefinitions: 'goto', multipleDeclarations: 'goto', multipleTypeDefinitions: 'goto' },
  links: false, unicodeHighlight: { ambiguousCharacters: false }, stickyScroll: { enabled: false } });
editor.onDidCompositionStart(() => { composing = true; updateAssistance(); });
editor.onDidCompositionEnd(() => { composing = false; updateAssistance(); schedule(); });
editor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyS, () => post('save', {}));
editor.addCommand(monaco.KeyCode.F6, () => post('focusHost', {}));
// Native WebView keyboard events do not bubble to the Avalonia window.
// Capture host shortcuts before Monaco/browser shortcuts, including focus in the assistance bar.
window.addEventListener('keydown', event => {
  if (!['F5', 'F12'].includes(event.key) || event.ctrlKey || event.metaKey || event.altKey || event.shiftKey) return;
  event.preventDefault(); event.stopImmediatePropagation();
  if (!event.repeat && !event.isComposing && !composing && initialized && !locked && !disposed && (event.key === 'F12' || !readOnly))
    post(event.key === 'F5' ? 'apply' : 'console', {});
}, true);
// Stay in the current module: model switching is owned by the host and must
// preserve the original cell/module's composition and draft acknowledgement.
editor.addCommand(monaco.KeyCode.F8, () => { if (!composing) void editor.getAction('editor.action.marker.next')?.run(); });
editor.addCommand(monaco.KeyMod.Shift | monaco.KeyCode.F8, () => { if (!composing) void editor.getAction('editor.action.marker.prev')?.run(); });
for (const button of assistanceButtons) {
  button.addEventListener('mousedown', event => event.preventDefault()); // Preserve the caret/selection.
  button.addEventListener('click', () => {
    if (button.disabled || disposed) return;
    editor.focus();
    const action = editor.getAction(button.dataset.action!);
    if (action) void action.run().catch(fail);
    else editor.trigger('toolbar', button.dataset.action!, {}); // Newer Monaco navigation commands use Action2.
  });
}
function updateAssistance() {
  for (const button of assistanceButtons)
    button.disabled = disposed || !initialized || composing || locked || ((readOnly || !!selectedDefinition) && button.dataset.action === 'editor.action.triggerSuggest');
}
function updateStatus() {
  if (disposed || !selected || document.body.hasAttribute('data-failed')) return;
  const model = editor.getModel()!;
  if (selectedDefinition) { status.textContent = `${selectedDefinition} · ${uiText('定义文件（只读）')}`; return; }
  const markers = monaco.editor.getModelMarkers({ resource: model.uri });
  const errors = markers.filter(m => m.severity === monaco.MarkerSeverity.Error).length;
  const warnings = markers.filter(m => m.severity === monaco.MarkerSeverity.Warning).length;
  status.textContent = `${selected} · ${model.getLanguageId() === 'typescript' ? 'TypeScript' : 'JavaScript'} · ${uiText('当前模块 {0} 个错误，{1} 个警告', errors, warnings)}`;
}
const markerListener = monaco.editor.onDidChangeMarkers(resources => {
  if (resources.some(uri => uri.toString() === editor.getModel()?.uri.toString())) updateStatus();
});

function setLocked(value: boolean) { locked = value; editor.updateOptions({ readOnly: readOnly || locked || !!selectedDefinition }); updateAssistance(); }
function languageWorker(model = editor.getModel()!) { return model.getLanguageId() === 'typescript' ? typescript.getTypeScriptWorker() : typescript.getJavaScriptWorker(); }
function saveView() {
  const previous = selectedDefinition ? definitions.get(selectedDefinition) : selected ? models.get(selected) : undefined;
  if (previous) previous.state = editor.saveViewState();
}
function select(id: string) {
  const next = models.get(id); if (!next) throw new Error('Unknown module');
  saveView(); selectedDefinition = undefined;
  selected = id; editor.setModel(next.model); if (next.state) editor.restoreViewState(next.state); setLocked(locked); editor.focus();
  updateStatus();
}
async function openDefinition(name: string) {
  if (!isDefinition(name)) throw new Error('Unknown definition file');
  let next = definitions.get(name);
  if (!next) {
    const uri = name === 'quickexcel.d.ts' ? declarationUri : monaco.Uri.parse('file://' + (name.startsWith('/') ? name : '/' + name));
    let model = monaco.editor.getModel(uri);
    if (!model) {
      let text: string | undefined = declarations;
      if (Object.hasOwn(builtinDeclarations.files, name)) text = (builtinDeclarations.files as Record<string,string>)[name];
      else if (name !== 'quickexcel.d.ts') {
        // Monaco registers each language service lazily. A JavaScript-only book
        // has no TypeScript worker yet; its existing worker also owns the libs.
        const source = models.get(selected!)!.model;
        const getWorker = await languageWorker(source);
        const worker = await getWorker(source.uri);
        text = await worker.getScriptText(name);
      }
      if (text === undefined) throw new Error('Definition file is not bundled');
      model = monaco.editor.createModel(text, 'typescript', uri);
    }
    definitions.set(name, next = { model, state: null });
  }
  if (disposed) return;
  saveView(); selectedDefinition = name; editor.setModel(next.model);
  if (next.state) editor.restoreViewState(next.state);
  setLocked(locked); updateStatus(); editor.focus();
}
function locate(line: number, column: number) {
  const position = editor.getModel()!.validatePosition({ lineNumber: Math.max(1, line || 1), column: Math.max(1, column || 1) });
  editor.setPosition(position); editor.revealPositionInCenter(position);
}
// Resolve only workbook model identities and bundled declarations. Never open a
// file path, external URI or workbook-supplied URL through the browser or host.
const editorOpener = monaco.editor.registerEditorOpener({ async openCodeEditor(_source, resource, position) {
  if (!initialized || disposed || composing || locked || navigating) return true;
  if (resource.toString() === editor.getModel()?.uri.toString()) return false;
  const module = [...models].find(([, item]) => item.model.uri.toString() === resource.toString())?.[0];
  const name = Object.hasOwn(builtinDeclarations.files, resource.path) ? resource.path : resource.path.slice(1);
  const definition = resource.toString() === declarationUri.toString() ? 'quickexcel.d.ts' :
    resource.scheme === 'file' && !resource.authority && !resource.query && !resource.fragment && isDefinition(name) ? name : undefined;
  if (!module && !definition) return true;
  const line = position && ('startLineNumber' in position ? position.startLineNumber : position.lineNumber) || 1;
  const column = position && ('startColumn' in position ? position.startColumn : position.column) || 1;
  navigating = true;
  try { await request('navigate', 'navigated', { module, definition, line, column }); }
  finally { navigating = false; }
  return true;
} });
const linkOpener = monaco.editor.registerLinkOpener({ open(resource) { return resource.scheme !== 'file'; } });
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
    if (bytes.length > 1048576 || [...models.values()].reduce((sum, m) => sum + utf8.encode(m.model.getValue()).length, 0) > 1048576) throw new Error(uiText('脚本总量超过 1 MiB，请恢复后缩减源码'));
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
    if (p.action === 'consoleShortcut' || p.action === 'applyShortcut') {
      const key = p.action === 'applyShortcut' ? 'F5' : 'F12';
      const event = new KeyboardEvent('keydown', { key, code: key, bubbles: true, cancelable: true,
        repeat: !!p.repeat, isComposing: !!p.composing, ctrlKey: !!p.control });
      editor.getDomNode()!.dispatchEvent(event);
      post('testResult', { prevented: event.defaultPrevented }, message.requestId); return;
    }
    if (p.action === 'inputSample') {
      const started = performance.now(); const model = editor.getModel()!; const end = model.getPositionAt(model.getValueLength()); const before = model.getPositionAt(model.getValueLength() - 1);
      editor.executeEdits('benchmark', [{ range: new monaco.Range(before.lineNumber, before.column, end.lineNumber, end.column), text: p.character }]);
      await enqueue(); await new Promise(requestAnimationFrame);
      post('testResult', { elapsedMs: performance.now() - started }, message.requestId); return;
    }
    if (p.action === 'completionSample') {
      const started = performance.now(); const getWorker = await languageWorker(); const worker = await getWorker(editor.getModel()!.uri);
      const completion = await worker.getCompletionsAtPosition(editor.getModel()!.uri.toString(), editor.getValue().indexOf('host.functions.') + 'host.functions.'.length);
      post('testResult', { elapsedMs: performance.now() - started, registered: completion?.entries.some(e => e.name === 'register') }, message.requestId); return;
    }
    if (p.action === 'edit') {
      select(p.module); editor.executeEdits('smoke', [{ range: editor.getModel()!.getFullModelRange(), text: p.text }]); editor.pushUndoStop();
      post('testResult', {}, message.requestId); return;
    }
    if (p.action === 'undo') { editor.trigger('smoke', 'undo', null); post('testResult', {}, message.requestId); return; }
    if (p.action === 'crash') { models.get(selected!)!.version++; post('draftChanged', { module: selected, version: models.get(selected!)!.version }); fail('Injected page failure before final source acknowledgment'); post('testResult', {}, message.requestId); return; }
    if (p.action === 'inspect') {
      const getWorker = await languageWorker(); const worker = await getWorker(editor.getModel()!.uri);
      const diagnostics = await worker.getSemanticDiagnostics(editor.getModel()!.uri.toString());
      const completion = await worker.getCompletionsAtPosition(editor.getModel()!.uri.toString(), p.offset ?? editor.getModel()!.getValueLength());
      post('testResult', { selected, definition: selectedDefinition ?? null, language: editor.getModel()!.getLanguageId(), source: editor.getValue(), workers: workers.size, readOnly: editor.getOption(monaco.editor.EditorOption.readOnly),
        diagnostics: diagnostics.map(d => ({ code: d.code, text: typeof d.messageText === 'string' ? d.messageText : d.messageText.messageText })), completions: completion?.entries.map(e => e.name).slice(0, 500) ?? [] }, message.requestId); return;
    }
    if (p.action === 'assistance') {
      // Exercise the actual editor actions/widgets, not just language-worker RPC.
      if (p.offset !== undefined) {
        // Test fixtures use LF; the editor may preserve CRLF on Windows.
        const lines = editor.getValue({ lineEnding: '\n', preserveBOM: false }).slice(0, p.offset).split('\n');
        editor.setPosition({ lineNumber: lines.length, column: lines[lines.length - 1].length + 1 }); editor.focus();
      }
      if (p.typeText !== undefined) editor.trigger('keyboard', 'type', { text: p.typeText });
      if (p.button) {
        const button = assistanceButtons.find(b => b.dataset.action === p.button);
        if (button) button.click(); else { const action = editor.getAction(p.button); if (action) await action.run(); else editor.trigger('smoke', p.button, {}); }
      }
      if (p.command) editor.trigger('smoke', p.command, {});
      if (p.link) {
        const link = [...document.querySelectorAll<HTMLAnchorElement>('.monaco-hover a, .zone-widget a')]
          .find(a => a.getClientRects().length > 0 && a.textContent?.includes(p.link));
        if (!link) throw new Error('Definition link is not visible');
        link.click();
      }
      const visibleText = (selector: string) => [...document.querySelectorAll<HTMLElement>(selector)]
        .filter(e => e.getClientRects().length > 0 && getComputedStyle(e).visibility !== 'hidden' && getComputedStyle(e).display !== 'none')
        .map(e => e.textContent).join('\n');
      post('testResult', { source: editor.getValue(), selected, definition: selectedDefinition ?? null, readOnly: editor.getOption(monaco.editor.EditorOption.readOnly), position: editor.getPosition(), status: status.textContent,
        actions: ['editor.action.triggerSuggest', 'editor.action.triggerParameterHints', 'editor.action.showHover', 'editor.action.marker.next', 'editor.action.revealDefinition'].filter(id => editor.getAction(id)),
        suggestions: visibleText('.suggest-widget .monaco-list-row'), hover: visibleText('.monaco-hover'),
        parameters: visibleText('.parameter-hints-widget'), problems: visibleText('.monaco-editor .zone-widget'),
        markers: monaco.editor.getModelMarkers({ resource: editor.getModel()!.uri }).map(m => ({ code: m.code, line: m.startLineNumber, column: m.startColumn, message: m.message })),
        suggestDisabled: (disposed || !initialized || composing || locked || readOnly || !!selectedDefinition) }, message.requestId); return;
    }
    throw new Error('Unknown test operation');
  }
  switch (message.type) {
    case 'captureInsertion': {
      const accepted = initialized && !disposed && !readOnly && !locked && !composing && !selectedDefinition && selected === p.module;
      post('insertionCaptured', { accepted, anchor: accepted ? { module: selected, version: models.get(selected!)!.version, position: editor.getPosition() } : null }, message.requestId); return;
    }
    case 'insertTemplate': {
      const anchor = p.anchor;
      const accepted = initialized && !disposed && !readOnly && !locked && !composing && !selectedDefinition &&
        selected === anchor?.module && models.get(selected!)?.version === anchor.version &&
        typeof p.text === 'string' && p.text.length <= 16384 &&
        [...models.values()].reduce((sum, m) => sum + utf8.encode(m.model.getValue()).length, utf8.encode(p.text).length) <= 1048576;
      if (accepted) {
        const position = anchor.position;
        editor.pushUndoStop();
        editor.executeEdits('event-template', [{ range: new monaco.Range(position.lineNumber, position.column, position.lineNumber, position.column), text: p.text, forceMoveMarkers: true }]);
        editor.pushUndoStop(); editor.focus();
      }
      post('templateInserted', { accepted }, message.requestId); return;
    }
    case 'editorCommand': {
      const commands = ['undo', 'redo', 'editor.action.triggerSuggest', 'editor.action.triggerParameterHints', 'editor.action.showHover', 'editor.action.revealDefinition', 'editor.action.marker.next', 'editor.action.formatDocument'];
      const accepted = initialized && !disposed && !locked && !composing && commands.includes(p.command) &&
        (!(readOnly || selectedDefinition) || !['undo', 'redo', 'editor.action.formatDocument'].includes(p.command));
      if (accepted) { editor.focus(); const action = editor.getAction(p.command); if (action) await action.run(); else editor.trigger('menu', p.command, {}); }
      post('commandExecuted', { accepted }, message.requestId); return;
    }
    case 'initialize': {
      if (models.size || !Array.isArray(p.modules) || p.modules.length < 1 || p.modules.length > 64) throw new Error('Invalid model initialization');
      readOnly = p.readOnly === true;
      if (p.apiVersion !== 4) {
        for (const lib of extraLibs) lib.dispose();
        extraLibs = [typescript.javascriptDefaults, typescript.typescriptDefaults].map(d => d.addExtraLib(
          declarations.replace('interface Host {', 'interface Host extends LegacyRegistration {'), baseUri + 'quickexcel.d.ts'));
      }
      const paths: Record<string, string[]> = Object.assign(Object.create(null), Object.fromEntries(Object.entries(builtinDeclarations.paths).map(([id,name])=>[id,['file://'+name]])));
      let hasTypeScript = false;
      for (const id of p.modules) {
        if (typeof id !== 'string' || !/^[A-Za-z0-9_.-]{1,128}$/.test(id) || id === '.' || id === '..' || models.has(id)) throw new Error('Invalid module ID');
        const language = p.languages?.[id] ?? 'javascript';
        if (language !== 'javascript' && language !== 'typescript') throw new Error('Invalid module language');
        hasTypeScript ||= language === 'typescript';
        const extension = language === 'typescript' ? '.ts' : '.js';
        const uri = monaco.Uri.parse(baseUri + 'modules/' + id + extension);
        models.set(id, { model: monaco.editor.createModel('', language, uri), version: 1, accepted: 1, state: null });
        paths[id] = ['modules/' + id + extension];
      }
      // Monaco's public enum lags its pinned TS 5.9.3 worker. These frozen
      // numeric values are ES2022=9, Bundler=100 and Force module detection=3.
      const options = { allowJs: true, checkJs: !hasTypeScript, allowNonTsExtensions: true, target: 9,
        module: typescript.ModuleKind.ESNext, moduleResolution: 100, moduleDetection: 3, lib: ['es2022'], types: [],
        strict: hasTypeScript, isolatedModules: true, verbatimModuleSyntax: true, useDefineForClassFields: true,
        baseUrl: baseUri, paths };
      typescript.javascriptDefaults.setCompilerOptions(options);
      typescript.typescriptDefaults.setCompilerOptions(options);
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
      const getWorker = await languageWorker(); const worker = await getWorker(editor.getModel()!.uri);
      await worker.getSyntacticDiagnostics(editor.getModel()!.uri.toString());
      setLocked(false); post('loaded', { modules: models.size, workers: workers.size }, message.requestId); return;
    }
    case 'flushDraft': {
      if (!initialized) throw new Error('Editor has not loaded');
      // Composition belongs to the original model; do not end it or run incomplete text.
      if (composing) { post('draftAck', { busy: uiText('中文组合输入尚未结束，请完成输入后重试。') }, message.requestId); return; }
      setLocked(true); if (scheduled !== undefined) { clearTimeout(scheduled); scheduled = undefined; }
      await enqueue(); post('draftAck', { versions: Object.fromEntries([...models].map(([id, m]) => [id, m.accepted])) }, message.requestId); return;
    }
    case 'releaseDraft': setLocked(false); post('released', {}, message.requestId); return;
    case 'select': if (!initialized || composing) throw new Error('Cannot switch during composition'); select(p.module); post('selected', {}, message.requestId); return;
    case 'definition': {
      if (!initialized || composing) throw new Error('Cannot switch during composition');
      await openDefinition(p.name);
      if (p.line != null) locate(p.line, p.column);
      post('definitionSelected', {}, message.requestId); return;
    }
    case 'locate': {
      const item = models.get(p.module); if (!item || item.accepted !== p.version || item.version !== p.version) throw new Error('Diagnostic belongs to an older source');
      if (composing) throw new Error('Cannot switch during composition');
      select(p.module); locate(p.line, p.column); post('located', {}, message.requestId); return;
    }
    default: throw new Error('Unsupported script page message');
  }
}
async function dispose() {
  if (disposed) return; disposed = true; if (scheduled !== undefined) clearTimeout(scheduled);
  markerListener.dispose(); editorOpener.dispose(); linkOpener.dispose(); updateAssistance();
  editor.dispose(); for (const model of monaco.editor.getModels()) model.dispose(); models.clear(); definitions.clear(); for (const lib of [...extraLibs,...builtinExtraLibs]) lib.dispose();
  for (const worker of workers) worker.terminate(); workers.clear();
}
start(handle, dispose);
