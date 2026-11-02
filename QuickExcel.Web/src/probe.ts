import * as monaco from 'monaco-editor/editor/editor.api';
import * as typescript from 'monaco-editor/language/typescript/monaco.contribution';
import 'monaco-editor/languages/definitions/javascript/register';
import { getDocument, GlobalWorkerOptions, version as pdfVersion } from 'pdfjs-dist';

// This page accepts only probe requests; no workbook editing or script execution capability.
const instance = location.pathname.split('/')[1];
const workers: string[] = [];
const workerHandles: Worker[] = [];
const NativeWorker = window.Worker;
window.Worker = class extends NativeWorker {
  constructor(url: string | URL, options?: WorkerOptions) {
    super(url, options);
    workers.push(String(url));
    workerHandles.push(this);
  }
};
(window as any).MonacoEnvironment = {
  getWorker(_moduleId: string, label: string) {
    return new Worker(new URL(label === 'typescript' || label === 'javascript' ? 'ts.worker.js' : 'editor.worker.js', location.href), { type: 'module' });
  }
};
GlobalWorkerOptions.workerSrc = new URL('pdf.worker.mjs', location.href).href;
const status = document.querySelector('#status')!;
let release: () => Promise<{ text: string } | undefined> = async () => undefined;
const post = (type: string, payload: unknown) => (window as any).invokeCSharpAction(JSON.stringify({ protocol: 1, instance, type, payload }));
(window as any).quickExcelProbe = async (request: { protocol: number; instance: string; type: string; requestId: string }) => {
  if (request.protocol !== 1 || request.instance !== instance) return;
  if (request.type === 'ping') post('pong', { requestId: request.requestId });
  if (request.type === 'dispose') post('disposed', { requestId: request.requestId, editor: await release() });
};
window.addEventListener('error', event => post('error', String(event.message)));
window.addEventListener('unhandledrejection', event => post('error', String(event.reason)));

async function run() {
  const started = performance.now();
  typescript.javascriptDefaults.setCompilerOptions({ allowJs: true, checkJs: true, noLib: true });
  typescript.javascriptDefaults.setDiagnosticsOptions({ noSemanticValidation: false, noSyntaxValidation: false });
  const model = monaco.editor.createModel('/** @type {number} */\nlet value = "中文";\n', 'javascript', monaco.Uri.parse('inmemory://quickexcel/a0.js'));
  const editor = monaco.editor.create(document.querySelector('#editor') as HTMLElement, { model, automaticLayout: true, minimap: { enabled: false } });
  const getWorker = await typescript.getJavaScriptWorker();
  const service = await getWorker(model.uri);
  const diagnostics = await service.getSemanticDiagnostics(model.uri.toString());
  if (!diagnostics.some(d => d.code === 2322)) throw new Error('Monaco language worker did not report the deliberate type error');
  const documentTask = getDocument({ url: new URL('fixture.pdf', location.href).href,
    cMapUrl: new URL('cmaps/', location.href).href, cMapPacked: true,
    standardFontDataUrl: new URL('standard_fonts/', location.href).href,
    wasmUrl: new URL('wasm/', location.href).href, isEvalSupported: false, enableXfa: false });
  const pdf = await documentTask.promise;
  const page = await pdf.getPage(1);
  const viewport = page.getViewport({ scale: 1 });
  const canvas = document.querySelector('#pdf') as HTMLCanvasElement;
  canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
  await page.render({ canvas, viewport }).promise;
  const text = (await page.getTextContent()).items.map(item => 'str' in item ? item.str : '').join('');
  if (!text.includes('QuickExcel') || !text.normalize('NFKC').includes('中文')) throw new Error('PDF.js did not extract the English and Chinese fixture text');
  if (!workers.some(w => w.endsWith('ts.worker.js')) || !workers.some(w => w.endsWith('pdf.worker.mjs'))) throw new Error('Real workers were not created');
  status.textContent = '离线组件检查通过。可测试中文输入、焦点与弹窗；此程序不是产品宿主。';
  post('ready', { pdfVersion, pages: pdf.numPages, text, diagnostics: diagnostics.map(d => d.code), workers,
    elapsedMs: performance.now() - started, userAgent: navigator.userAgent });
  let released = false;
  release = async () => {
    if (released) return;
    released = true;
    const text = model.getValue();
    editor.dispose(); model.dispose(); await documentTask.destroy();
    for (const worker of workerHandles) worker.terminate();
    return { text };
  };
  window.addEventListener('pagehide', () => { void release(); }, { once: true });
}
run().catch(error => { status.textContent = String(error); post('error', String(error.stack ?? error)); });
