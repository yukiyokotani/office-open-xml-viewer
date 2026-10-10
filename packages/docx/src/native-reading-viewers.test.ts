import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';
import { DocxViewer } from './viewer.js';
import { DocxScrollViewer } from './scroll-viewer.js';
import { publishDocxLayout } from './document-layout-events.js';
import { FakeDocxEngine, installDom, makeContainer, makeEl, type FakeEl } from './scroll-viewer-test-dom.js';
import { nativeReadingNotices, noReadingNotices } from './native-reading-notice.js';
import type { DocumentLayout } from './layout/types.js';
const notices = nativeReadingNotices({ pages: [{ layers: { body: [{ kind: 'paragraph', nativeReadingRelocations: ['invented-scene'] }] } }], diagnostics: [] } as unknown as DocumentLayout);
afterEach(() => { vi.unstubAllGlobals(); vi.restoreAllMocks(); });
function descendants(element: FakeEl): FakeEl[] { return [element, ...element.children.flatMap(descendants)]; }
function noticeElements(root: FakeEl) { return descendants(root).filter(el => el.textContent === notices[0].message); }
function mount(kind: 'viewer' | 'scroll', onPageChange?: (page: number, count: number) => void) {
  installDom();
  // Standard DOM firstChild semantics, absent from the older general-purpose fake.
  const create = globalThis.document.createElement.bind(globalThis.document);
  vi.spyOn(globalThis.document, 'createElement').mockImplementation((tag: string) => {
    const element = create(tag);
    Object.defineProperty(element, 'firstChild', { get: () => (element as unknown as FakeEl).children[0] ?? null });
    return element;
  });
  const root = makeContainer(200, 400), onError = vi.fn();
  Object.defineProperty(root, 'firstChild', { get: () => root.children[0] ?? null });
  let surface: DocxViewer | DocxScrollViewer;
  if (kind === 'viewer') { const canvas = makeEl('canvas'); root.appendChild(canvas); surface = new DocxViewer(canvas as unknown as HTMLCanvasElement, { width: 100, onError, onPageChange }); }
  else surface = new DocxScrollViewer(root as unknown as HTMLElement, { width: 100, onError });
  return { root, surface, onError };
}
function readingEngine(root: FakeEl, visible = true) {
  const engine = new FakeDocxEngine(2, [{ widthPt: 100, heightPt: 200 }]);
  let current = visible ? notices : noReadingNotices;
  Object.defineProperty(engine, 'readingNotices', { get: () => current });
  Object.defineProperty(engine, '_readingPublicationOwned', { get: () => true });
  Object.defineProperty(engine, '_readingPublicationToken', { get: () => 0 });
  const invalidate = (error: Error) => {
    current = noReadingNotices; engine.setPageCount(0); engine.setLayoutFailure(error);
    publishDocxLayout(engine.asDoc(), { pageCount: 0, exact: false, complete: false, error });
  };
  Object.assign(engine, { _invalidateReadingLayout: invalidate });
  const render = engine.renderPage.bind(engine);
  engine.renderPage = (...args) => {
    const disclosure = noticeElements(root);
    expect(disclosure).toHaveLength(visible ? 1 : 0);
    if (visible) expect(disclosure[0].parentElement!.children[0]).toBe(disclosure[0]);
    return render(...args);
  };
  return { engine, invalidate };
}

describe.each(['viewer', 'scroll'] as const)('reading disclosure in built-in %s', kind => {
  it('discloses before paint, then clears every mounted page and notice on complete-publication failure', async () => {
    const { root, surface, onError } = mount(kind), { engine, invalidate } = readingEngine(root);
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    await surface.load('invented.doc');
    expect(noticeElements(root)).toHaveLength(1);
    expect(descendants(root).some(el => el.tag === 'canvas' && el.width > 0)).toBe(true);
    const error = new Error('late reading resource failure'); invalidate(error);
    expect(surface.pageCount).toBe(0); expect(noticeElements(root)).toHaveLength(0);
    expect(descendants(root).filter(el => el.tag === 'canvas').every(el => el.width === 0)).toBe(true);
    expect(onError).toHaveBeenCalledWith(error); surface.destroy();
  });

  it('clears notice-free reading-capable pages when the complete publication is revoked', async () => {
    const { root, surface, onError } = mount(kind), { engine, invalidate } = readingEngine(root, false);
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    await surface.load('invented-notice-free-reading.doc');
    expect(noticeElements(root)).toHaveLength(0);
    expect(descendants(root).some(el => el.tag === 'canvas' && el.width > 0)).toBe(true);
    const error = new Error('notice-free publication failed'); invalidate(error);
    expect(surface.pageCount).toBe(0);
    expect(descendants(root).filter(el => el.tag === 'canvas').every(el => el.width === 0)).toBe(true);
    expect(onError).toHaveBeenCalledWith(error); surface.destroy();
  });

  it('removes old disclosure and pages on ordinary replacement and cancellation', async () => {
    const { root, surface, onError } = mount(kind), { engine, invalidate } = readingEngine(root);
    const destroy = engine.destroy.bind(engine);
    engine.destroy = () => { invalidate(new Error('reading owner cancelled')); destroy(); };
    const ordinary = new FakeDocxEngine(1, [{ widthPt: 100, heightPt: 200 }]);
    vi.spyOn(DocxDocument, 'load').mockResolvedValueOnce(engine.asDoc()).mockResolvedValueOnce(ordinary.asDoc());
    await surface.load('invented-reading.doc');
    expect(noticeElements(root)).toHaveLength(1);
    await surface.load('invented-ordinary.doc');
    expect(engine.destroyed).toBe(true); expect(noticeElements(root)).toHaveLength(0);
    expect(surface.pageCount).toBe(1); surface.destroy();
    expect(noticeElements(root)).toHaveLength(0); expect(ordinary.destroyed).toBe(true);
    expect(onError).not.toHaveBeenCalled();
  });
});

it('reports a background reading render failure exactly once through the real viewer router', async () => {
  const { root, surface, onError } = mount('viewer'), { engine, invalidate } = readingEngine(root);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  await surface.load('invented.doc');
  const error = new Error('background page acquisition failed');
  engine.renderPage = async () => { invalidate(error); throw error; };
  publishDocxLayout(engine.asDoc(), { pageCount: 2, exact: true, complete: true });
  await vi.waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  expect(onError).toHaveBeenCalledWith(error); expect(surface.pageCount).toBe(0); surface.destroy();
});

it('preserves valid reading pages when a public page-change callback throws after paint', async () => {
  const error = new TypeError('page-change consumer failed');
  let fail = false;
  const callback = vi.fn(() => { if (fail) throw error; });
  const { root, surface, onError } = mount('viewer', callback), { engine } = readingEngine(root);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  await surface.load('invented.doc'); fail = true;
  await expect((surface as DocxViewer).nextPage()).rejects.toBe(error);
  expect(surface.pageCount).toBe(2); expect(noticeElements(root)).toHaveLength(1);
  expect(onError).not.toHaveBeenCalled(); surface.destroy();
});

it('allows replacement while a former reading opening window waits for completion', async () => {
  const { root, surface, onError } = mount('scroll'), { engine } = readingEngine(root, false);
  const ordinary = new FakeDocxEngine(1, [{ widthPt: 100, heightPt: 200 }]);
  let waiting!: () => void, rejectCompletion!: (error: Error) => void;
  const started = new Promise<void>(resolve => { waiting = resolve; });
  engine.waitUntilLayoutComplete = () => { waiting(); return new Promise((_, reject) => { rejectCompletion = reject; }); };
  const destroy = engine.destroy.bind(engine);
  engine.destroy = () => { destroy(); rejectCompletion(new Error('former owner cancelled')); };
  vi.spyOn(DocxDocument, 'load').mockResolvedValueOnce(engine.asDoc()).mockResolvedValueOnce(ordinary.asDoc());
  const first = surface.load('invented-reading.doc');
  await started; await surface.load('invented-ordinary.doc'); await first;
  expect(surface.pageCount).toBe(1); expect(onError).not.toHaveBeenCalled(); surface.destroy();
});

it('preserves the awaited reading render failure when revocation invalidates the canvas dispatcher', async () => {
  const { root, surface, onError } = mount('viewer'), { engine, invalidate } = readingEngine(root);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  await surface.load('invented.doc');
  const error = new Error('navigation reading decode failed');
  engine.renderPage = async () => { invalidate(error); throw error; };
  await expect((surface as DocxViewer).nextPage()).rejects.toBe(error);
  expect(onError).not.toHaveBeenCalled(); expect(noticeElements(root)).toHaveLength(0);
  surface.destroy();
});

it('rejects awaited navigation revocation in a notice-free reading-capable view', async () => {
  const { root, surface, onError } = mount('viewer'), { engine, invalidate } = readingEngine(root, false);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  await surface.load('invented-notice-free-reading.doc');
  expect(noticeElements(root)).toHaveLength(0);
  const error = new Error('notice-free navigation failed');
  engine.renderPage = async () => { invalidate(error); throw error; };
  await expect((surface as DocxViewer).nextPage()).rejects.toBe(error);
  expect(surface.pageCount).toBe(0); expect(onError).not.toHaveBeenCalled(); surface.destroy();
});

it('rejects initial scroll load after reading revocation recycles the awaited slots', async () => {
  const { root, surface, onError } = mount('scroll'), { engine, invalidate } = readingEngine(root, false);
  const error = new Error('opening reading page failed');
  engine.renderPage = async () => { invalidate(error); throw error; };
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  await expect(surface.load('invented-opening-reading.doc')).rejects.toBe(error);
  expect(surface.pageCount).toBe(0); expect(onError).not.toHaveBeenCalled();
  expect(noticeElements(root)).toHaveLength(0); surface.destroy();
});
