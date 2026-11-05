import { uiText } from './localization';
import { getDocument, GlobalWorkerOptions, AnnotationMode, AnnotationEditorType } from 'pdfjs-dist';
import { PDFViewer, EventBus, PDFLinkService, PDFFindController, RenderingStates } from 'pdfjs-dist/web/pdf_viewer.mjs';
import 'pdfjs-dist/web/pdf_viewer.css';
import { start, post, fail, Message } from './bridge';

declare const __QUICKEXCEL_WEB_TEST__: boolean;
GlobalWorkerOptions.workerSrc = new URL('pdf.worker.mjs', location.href).href;
const container = document.getElementById('viewerContainer')!;
const thumbnails = document.getElementById('thumbnails')!;
const status = document.getElementById('status')!;
const pageInput = document.getElementById('page') as HTMLInputElement;
const search = document.getElementById('find') as HTMLInputElement;
const eventBus = new EventBus();
const life = new AbortController();
// Only internal PDF destinations are active; workbook URLs remain inert.
const linkService = new PDFLinkService({ eventBus });
linkService.addLinkAttributes = (link: HTMLAnchorElement) => { link.removeAttribute('href'); link.addEventListener('click', e => e.preventDefault()); };
const findController = new PDFFindController({ eventBus, linkService });
let viewer: PDFViewer;
let documentTask: ReturnType<typeof getDocument> | undefined;
let pdf: Awaited<ReturnType<typeof getDocument>['promise']> | undefined;
let disposed = false;
let firstPage: ((value?: unknown) => void) | undefined;
let firstFailure: ((reason: unknown) => void) | undefined;
let firstPaintMs = 0, loadStarted = 0;

// Uses PDF.js' existing Viewer/PageView/text/search layers with a bounded rendering queue.
// Destroying old PageViews releases canvases and text layers; no private Viewer cache fields are patched.
class BoundedQueue {
  active: any;
  pending: Promise<unknown> = Promise.resolve();
  cache = new Map<number, any>();
  stopped = false;
  scheduled = false;
  hasViewer() { return true; }
  isHighestPriority(view: any) { return this.active === view; }
  renderHighestPriority(visible?: any) {
    if (this.stopped || this.active || this.scheduled || !pdf) return;
    this.scheduled = true;
    queueMicrotask(() => { this.scheduled = false; if (!this.stopped && !this.active) viewer.forceRendering(visible); });
  }
  getHighestPriority(visible: any, views: any[], ahead: boolean) {
    if (this.stopped || this.active) return null;
    const shown = visible.views.slice(0, 8);
    for (const { view } of shown) {
      if (this.cache.has(view.id)) { this.cache.delete(view.id); this.cache.set(view.id, view); }
      if (view.renderingState !== RenderingStates.FINISHED) return view;
    }
    if (shown.length < 8 && shown.length > 0) {
      const index = ahead ? visible.last.id : visible.first.id - 2;
      const neighbor = views[index]; if (neighbor && neighbor.renderingState !== RenderingStates.FINISHED) return neighbor;
    }
    return null;
  }
  renderView(view: any) {
    if (this.stopped || this.active || view.renderingState === RenderingStates.FINISHED) return false;
    // An old zoom canvas is released before its replacement is allocated.
    view.reset(); this.cache.delete(view.id);
    while (this.cache.size >= 8) { const [id, old] = this.cache.entries().next().value!; old.destroy(); this.cache.delete(id); }
    this.cache.set(view.id, view); this.active = view;
    this.pending = view.draw().catch((error: Error) => { if (error.name !== 'RenderingCancelledException' && !this.stopped) { firstFailure?.(error); fail(error); } })
      .finally(() => { if (this.active === view) this.active = undefined; this.renderHighestPriority(); });
    return true;
  }
  async stop() { this.stopped = true; for (const view of this.cache.values()) view.cancelRendering(); await this.pending; for (const view of this.cache.values()) view.destroy(); this.cache.clear(); }
}
const queue = new BoundedQueue();
viewer = new PDFViewer({ container, viewer: document.getElementById('viewer')!, eventBus, linkService, findController, renderingQueue: queue,
  textLayerMode: 1, annotationMode: AnnotationMode.DISABLE, annotationEditorMode: AnnotationEditorType.DISABLE,
  enableAutoLinking: false, enableDetailCanvas: false, enableOptimizedPartialRendering: false, enableSelectionRendering: false,
  maxCanvasPixels: 2_000_000, maxCanvasDim: 4096, minDurationToUpdateCanvas: 0, supportsPinchToZoom: false, abortSignal: life.signal });
linkService.setViewer(viewer);
eventBus.on('pagesinit', () => { viewer.currentScaleValue = 'page-fit'; viewer.update(); });
eventBus.on('pagerendered', ({ pageNumber, error }: any) => {
  if (error) { firstFailure?.(error); return; }
  if (firstPage) { firstPaintMs = performance.now() - loadStarted; requestAnimationFrame(() => { firstPage?.(); firstPage = undefined; }); }
  status.textContent = uiText('第 {0} 页 · {1} 页 · 预览与导出使用同一打印任务', pageNumber, pdf?.numPages ?? 0);
});
eventBus.on('pagechanging', ({ pageNumber }: any) => {
  pageInput.value = String(pageNumber);
  thumbnails.querySelector('[aria-current=true]')?.removeAttribute('aria-current'); thumbnails.children[pageNumber - 1]?.setAttribute('aria-current', 'true');
});
eventBus.on('updatefindmatchescount', ({ matchesCount }: any) => { status.textContent = uiText('查找：{0} / {1}', matchesCount.current, matchesCount.total); });
const on = (id: string, action: () => void) => document.getElementById(id)!.addEventListener('click', action, { signal: life.signal });
on('previous', () => { if (pdf) viewer.currentPageNumber = Math.max(1, viewer.currentPageNumber - 1); });
on('next', () => { if (pdf) viewer.currentPageNumber = Math.min(pdf.numPages, viewer.currentPageNumber + 1); });
pageInput.addEventListener('change', () => { if (pdf) viewer.currentPageNumber = Math.max(1, Math.min(pdf.numPages, Number(pageInput.value) || 1)); }, { signal: life.signal });
function zoom(value: string) {
  if (!pdf) return;
  if (value !== 'page-fit' && value !== 'page-width' && (!Number.isFinite(Number(value)) || Number(value) < .25 || Number(value) > 4)) throw new Error('Zoom outside supported range');
  for (const view of queue.cache.values()) view.reset();
  viewer.currentScaleValue = value; viewer.update();
}
document.getElementById('zoom')!.addEventListener('change', e => zoom((e.target as HTMLSelectElement).value), { signal: life.signal });
function find(previous = false, again = true) {
  eventBus.dispatch('find', { source: window, type: again ? 'again' : '', query: search.value.slice(0, 512), phraseSearch: true, caseSensitive: false, entireWord: false, highlightAll: true, findPrevious: previous });
}
search.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); find(e.shiftKey); } }, { signal: life.signal });
on('findNext', () => find()); on('findPrevious', () => find(true));
on('export', () => post('exportPdf', {})); on('print', () => post('print', {}));
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') { e.preventDefault(); post('print', {}); }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); post('exportPdf', {}); }
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') { e.preventDefault(); search.focus(); search.select(); }
  if (e.key === 'F6') { e.preventDefault(); post('focusHost', {}); }
}, { signal: life.signal });

const thumbnailCache = new Map<number, HTMLCanvasElement>();
const thumbnailPending = new Set<number>();
let thumbnailWork: Promise<void> = Promise.resolve();
let thumbnailRender: any;
const observer = new IntersectionObserver(entries => {
  for (const entry of entries) if (entry.isIntersecting) thumbnailPending.add(Number((entry.target as HTMLElement).dataset.page));
  thumbnailWork = thumbnailWork.then(renderThumbnails).catch(error => { if (!disposed && error.name !== 'RenderingCancelledException') fail(error); });
}, { root: thumbnails, rootMargin: '160px' });
async function renderThumbnails() {
  for (const number of thumbnailPending) {
    thumbnailPending.delete(number); if (disposed || !pdf) return;
    const button = thumbnails.children[number - 1] as HTMLElement;
    if (!button || button.getBoundingClientRect().bottom < thumbnails.getBoundingClientRect().top - 160 || button.getBoundingClientRect().top > thumbnails.getBoundingClientRect().bottom + 160) continue;
    if (thumbnailCache.has(number)) continue;
    while (thumbnailCache.size >= 12) { const [id, old] = thumbnailCache.entries().next().value!; old.width = old.height = 0; old.remove(); thumbnailCache.delete(id); }
    const page = await pdf.getPage(number); if (disposed) return;
    const natural = page.getViewport({ scale: 1 }); const viewport = page.getViewport({ scale: Math.min(112 / natural.width, 160 / natural.height) });
    const canvas = document.createElement('canvas'); canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
    thumbnailCache.set(number, canvas); button.prepend(canvas);
    thumbnailRender = page.render({ canvas, viewport, annotationMode: AnnotationMode.DISABLE });
    await thumbnailRender.promise; thumbnailRender = undefined;
  }
}
async function handle(message: Message) {
  const p = message.payload;
  if (__QUICKEXCEL_WEB_TEST__ && message.type === 'test') {
    if (p.action === 'extractFixture') {
      const url = new URL(p.url);
      if (url.origin !== location.origin || !url.pathname.startsWith('/' + location.pathname.split('/')[1] + '/documents/')) throw new Error('Invalid test resource');
      const task = getDocument({ url: url.href, isEvalSupported: false, disableFontFace: true, useSystemFonts: false });
      try {
        const fixture = await task.promise, page = await fixture.getPage(1);
        const content = await page.getTextContent({ disableNormalization: true, includeMarkedContent: true });
        post('testResult', { text: content.items.filter((item: any) => typeof item.str === 'string').map((item: any) => item.str).join('') }, message.requestId);
      } finally { await task.destroy(); }
      return;
    }
    if (p.action === 'crash') { fail('Injected PDF page failure'); post('testResult', {}, message.requestId); return; }
    if (p.action === 'copyText') {
      const layer = viewer.getPageView(viewer.currentPageNumber - 1).div.querySelector('.textLayer');
      if (!layer?.textContent) { post('testResult', { copied: '' }, message.requestId); return; }
      const selection = document.getSelection()!, range = document.createRange(); range.selectNodeContents(layer);
      selection.removeAllRanges(); selection.addRange(range);
      let copied = '';
      const event = new Event('copy', { bubbles: true, cancelable: true });
      Object.defineProperty(event, 'clipboardData', { value: { setData(type: string, value: string) { if (type === 'text/plain') copied = value; } } });
      layer.dispatchEvent(event); selection.removeAllRanges();
      post('testResult', { copied }, message.requestId); return;
    }
    if (p.action === 'snapshot') {
      viewer.currentPageNumber = p.page; viewer.update();
      let view = viewer.getPageView(p.page - 1);
      for (let attempt = 0; attempt < 200 && view.renderingState !== RenderingStates.FINISHED; attempt++) await new Promise(resolve => setTimeout(resolve, 50));
      const canvas = view.div.querySelector('canvas') as HTMLCanvasElement;
      if (view.renderingState !== RenderingStates.FINISHED || !canvas?.width) throw new Error('Snapshot page is not rendered');
      const data = canvas.toDataURL('image/png').split(',')[1];
      post('testResult', { data: data.slice(p.offset, p.offset + 49152), length: data.length, width: canvas.width, height: canvas.height,
        scaleX: view.viewport.scale * view.outputScale.sx, scaleY: view.viewport.scale * view.outputScale.sy }, message.requestId); return;
    }
    if (p.action === 'zoomSample') {
      const started = performance.now();
      await new Promise<void>((resolve, reject) => {
        const timer = setTimeout(() => { eventBus.off('pagerendered', complete); reject(new Error('Visible zoom rendering timed out')); }, 10000);
        const complete = (event: any) => { if (!event.error && event.pageNumber === viewer.currentPageNumber) { clearTimeout(timer); eventBus.off('pagerendered', complete); requestAnimationFrame(() => resolve()); } };
        eventBus.on('pagerendered', complete); zoom(p.value);
      });
      post('testResult', { elapsedMs: performance.now() - started }, message.requestId); return;
    }
    if (p.action === 'page') { viewer.currentPageNumber = p.page; viewer.update(); }
    if (p.action === 'zoom') zoom(p.value);
    if (p.action === 'find') { search.value = p.text; find(false, false); }
    const canvases = [...document.querySelectorAll<HTMLCanvasElement>('#viewer canvas')];
    post('testResult', { pages: pdf?.numPages, page: viewer.currentPageNumber, firstPaintMs, canvases: canvases.filter(c => c.width && c.height).length,
      canvasBytes: canvases.reduce((n, c) => n + c.width * c.height * 4, 0), thumbnails: thumbnailCache.size,
      toolbarColor: getComputedStyle(document.getElementById('export')!).color,
      text: (container.textContent ?? '').slice(0, 5000), matches: findController.pageMatches?.reduce((n: number, p: any[]) => n + p.length, 0) ?? 0 }, message.requestId); return;
  }
  if (message.type !== 'loadPdf' || documentTask) throw new Error('Unsupported preview page request');
  const url = new URL(p.url); if (url.origin !== location.origin || !new RegExp('^/' + location.pathname.split('/')[1] + '/documents/[a-f0-9]{48}\\.pdf$').test(url.pathname) || url.search || url.hash) throw new Error('Invalid preview resource');
  loadStarted = performance.now();
  const first = new Promise((resolve, reject) => { firstPage = resolve; firstFailure = reject; });
  documentTask = getDocument({ url: url.href, cMapUrl: new URL('cmaps/', location.href).href, cMapPacked: true,
    standardFontDataUrl: new URL('standard_fonts/', location.href).href, wasmUrl: new URL('wasm/', location.href).href,
    isEvalSupported: false, enableXfa: false, disableFontFace: true, useSystemFonts: false, disableAutoFetch: true, disableStream: true, rangeChunkSize: 65536, maxImageSize: 16_777_216 });
  pdf = await documentTask.promise; if (disposed) return;
  if (pdf.numPages > 2000 || pdf.numPages < 1) throw new Error('PDF page count exceeds the 2,000 page preview budget');
  pageInput.max = String(pdf.numPages); document.getElementById('count')!.textContent = '/ ' + pdf.numPages;
  linkService.setDocument(pdf); viewer.setDocument(pdf);
  for (let index = 1; index <= pdf.numPages; index++) {
    const button = document.createElement('button'); button.className = 'thumbnail'; button.dataset.page = String(index); button.textContent = String(index); button.setAttribute('aria-label', uiText('第 {0} 页', index));
    button.addEventListener('click', () => { viewer.currentPageNumber = index; }, { signal: life.signal }); thumbnails.append(button); observer.observe(button);
  }
  await first; post('pdfLoaded', { pages: pdf.numPages, firstPaintMs }, message.requestId);
}
async function dispose() {
  if (disposed) return; disposed = true; life.abort(); observer.disconnect(); thumbnailPending.clear(); thumbnailRender?.cancel();
  await queue.stop(); await thumbnailWork; viewer.setDocument(null); findController.setDocument(null); linkService.setDocument(null);
  for (const canvas of thumbnailCache.values()) { canvas.width = canvas.height = 0; canvas.remove(); } thumbnailCache.clear();
  await documentTask?.destroy(); pdf = undefined;
}
start(handle, dispose);
