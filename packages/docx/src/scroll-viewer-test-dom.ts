import { vi } from 'vitest';
import type { DocxDocument } from './document';
import type { DocxTextRunInfo } from './renderer';
import type { RenderPageOptions } from './types';
import type { WireRenderPageOptions } from './worker-protocol';
import type { DocxElementContext, DocxPagePoint } from './selection-context';
import type { DocxElementContextOptions } from './element-context';
import { DocxScrollViewer, type DocxScrollViewerOptions } from './scroll-viewer';

/** Test-only adapter for mechanics cases that need a preloaded engine. Public
 * callers use `DocxScrollViewer.fromDocument()` directly. */
export function makeBorrowedDocxScrollViewer(
  container: HTMLElement,
  options: DocxScrollViewerOptions & { document: DocxDocument },
): DocxScrollViewer {
  const { document, ...viewerOptions } = options;
  return (DocxScrollViewer.fromDocument as (
    target: HTMLElement,
    engine: DocxDocument,
    opts: DocxScrollViewerOptions,
  ) => DocxScrollViewer)(container, document, viewerOptions);
}

/** A recording fake DOM element. Extends the `FakeEl` pattern from
 *  text-layer.test.ts with the surface DocxScrollViewer touches: scrollTop,
 *  clientHeight/Width, event listeners, removeChild/remove, canvas ctx. */
export interface FakeEl {
  tag: string;
  /** Uppercase tag, mirroring the real DOM (`div` → `DIV`) — the viewer's
   *  canvas-container guard reads it. */
  tagName: string;
  textContent: string;
  innerHTML: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  parentElement: FakeEl | null;
  /** DOM alias of `parentElement` (viewer destroy reads `canvas.parentNode`). */
  readonly parentNode: FakeEl | null;
  /** The element following this one under its parent, or null (viewer destroy
   *  captures `canvas.nextSibling` so it can re-`insertBefore` at the original
   *  slot). Computed from the parent's child order, like the real DOM. */
  readonly nextSibling: FakeEl | null;
  // canvas-only
  width: number;
  height: number;
  // scrollHost-only (settable by the test to drive the virtualization loop)
  scrollTop: number;
  scrollLeft: number;
  clientHeight: number;
  clientWidth: number;
  /** Spacer-only: laid-out width. The real DOM derives it from style.width; the
   *  fake keeps it a settable number (default 0) so a horizontal-zoom test can
   *  give the scroll extent a concrete size. */
  offsetWidth: number;
  _listeners: Map<string, Array<(e: unknown) => void>>;
  _bitmapCtx?: { transferFromImageBitmap: (b: unknown) => void; lastBitmap: unknown };
  /** Records every DEVICE-BUFFER resize (a `canvas.width`/`canvas.height`
   *  assignment). The flicker-free CSS-preview path must NOT touch the device
   *  buffer of an in-window slot (it only CSS-resizes via `style.width/height`),
   *  so a preview leaves this array untouched; a settle/mount render appends. */
  _deviceResizes: Array<{ prop: 'width' | 'height'; value: number }>;
  /** Monotonic id assigned at creation, so tests can order operations and tell a
   *  freshly-created spare canvas from the on-screen one it replaces. */
  _uid: number;
  appendChild(c: FakeEl): FakeEl;
  removeChild(c: FakeEl): FakeEl;
  remove(): void;
  insertBefore(n: FakeEl, ref: FakeEl | null): FakeEl;
  contains(node: unknown): boolean;
  addEventListener(type: string, fn: (e: unknown) => void, opts?: unknown): void;
  removeEventListener(type: string, fn: (e: unknown) => void): void;
  getContext(kind: string): unknown;
  getBoundingClientRect(): { top: number; left: number; width: number; height: number };
  /** test-only: fire a recorded listener */
  dispatch(type: string, event?: unknown): void;
}

let _uidSeq = 0;

export function makeEl(tag: string): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
    tagName: tag.toUpperCase(),
    textContent: '',
    innerHTML: '',
    width: 0,
    height: 0,
    scrollTop: 0,
    scrollLeft: 0,
    clientHeight: 0,
    clientWidth: 0,
    offsetWidth: 0,
    children: [],
    parentElement: null,
    // Placeholders for the interface; replaced by computed getters below
    // (Object.defineProperty), mirroring innerHTML/width/height.
    parentNode: null,
    nextSibling: null,
    _listeners: new Map(),
    _deviceResizes: [],
    _uid: _uidSeq++,
    style: new Proxy(style as Record<string, string> & { cssText: string }, {
      set(target, prop: string, value: string) {
        if (prop === 'cssText') {
          for (const decl of value.split(';')) {
            const idx = decl.indexOf(':');
            if (idx > 0) target[decl.slice(0, idx).trim()] = decl.slice(idx + 1).trim();
          }
          target.cssText = value;
        } else {
          target[prop] = value;
        }
        return true;
      },
      get(target, prop: string) {
        return target[prop] ?? '';
      },
    }),
    appendChild(c: FakeEl) {
      // Real-DOM move semantics: detach from the current parent first so a
      // reparent (wrapper.appendChild(canvas)) removes the node from its old
      // slot rather than duplicating it.
      c.parentElement?.removeChild(c);
      c.parentElement = this;
      this.children.push(c);
      return c;
    },
    removeChild(c: FakeEl) {
      const i = this.children.indexOf(c);
      if (i >= 0) this.children.splice(i, 1);
      c.parentElement = null;
      return c;
    },
    remove() {
      this.parentElement?.removeChild(this);
    },
    insertBefore(n: FakeEl, ref: FakeEl | null) {
      // Real-DOM pre-insert validity: a non-null reference that is not a child
      // of this node throws NotFoundError. Modelled so viewer code cannot
      // silently rely on a stale sibling reference (real browsers would throw).
      if (ref && !this.children.includes(ref)) {
        throw new Error('NotFoundError: the node before which the new node is to be inserted is not a child of this node');
      }
      // Real-DOM move semantics: detach from the current parent first (see
      // appendChild). Detach BEFORE resolving `ref`'s index so re-inserting a
      // node relative to a sibling under the same parent still lands correctly.
      n.parentElement?.removeChild(n);
      n.parentElement = this;
      const i = ref ? this.children.indexOf(ref) : -1;
      if (i >= 0) this.children.splice(i, 0, n);
      else this.children.push(n);
      return n;
    },
    contains(node: unknown): boolean {
      if (node === this) return true;
      return this.children.some((child) => child.contains(node));
    },
    addEventListener(type: string, fn: (e: unknown) => void) {
      const arr = this._listeners.get(type) ?? [];
      arr.push(fn);
      this._listeners.set(type, arr);
    },
    removeEventListener(type: string, fn: (e: unknown) => void) {
      const arr = this._listeners.get(type);
      if (arr) this._listeners.set(type, arr.filter((f) => f !== fn));
    },
    getContext(kind: string) {
      if (kind === 'bitmaprenderer') {
        this._bitmapCtx = {
          lastBitmap: null,
          transferFromImageBitmap(b: unknown) {
            this.lastBitmap = b;
          },
        };
        return this._bitmapCtx;
      }
      // A minimal recording 2d context is never actually used by the viewer
      // (renderPage owns the ctx); return a no-op object for safety.
      return {};
    },
    getBoundingClientRect() {
      return { top: 0, left: 0, width: this.clientWidth, height: this.clientHeight };
    },
    dispatch(type: string, event: unknown = {}) {
      for (const fn of this._listeners.get(type) ?? []) fn(event);
    },
  };
  // Mirror the DOM: assigning `innerHTML = ''` detaches all children (both
  // buildDocxTextLayer's leading clear and the viewer's overlay-teardown clear on
  // recycle rely on this). A non-empty assignment only records the string (the
  // viewer/text-layer never set non-empty innerHTML, so no parsing is needed).
  let _html = '';
  Object.defineProperty(el, 'innerHTML', {
    get() {
      return _html;
    },
    set(value: string) {
      _html = value;
      if (value === '') {
        for (const c of el.children) c.parentElement = null;
        el.children.length = 0;
      }
    },
    enumerable: true,
    configurable: true,
  });
  // Record device-buffer resizes. A real canvas.width/height assignment resizes
  // (and CLEARS) the backing store; the flicker-free preview path must never do
  // this to an on-screen slot (it CSS-resizes instead). Making width/height
  // recording accessors lets the preview tests assert "no device-buffer resize
  // happened during the CSS preview" (test a). The value is still stored so the
  // worker transfer path (`canvas.width = bmp.width`) reads back correctly.
  let _w = 0;
  let _h = 0;
  Object.defineProperty(el, 'width', {
    get() {
      return _w;
    },
    set(value: number) {
      _w = value;
      el._deviceResizes.push({ prop: 'width', value });
    },
    enumerable: true,
    configurable: true,
  });
  Object.defineProperty(el, 'height', {
    get() {
      return _h;
    },
    set(value: number) {
      _h = value;
      el._deviceResizes.push({ prop: 'height', value });
    },
    enumerable: true,
    configurable: true,
  });
  // DOM-mirroring read-only relations. `parentNode` aliases `parentElement`;
  // `nextSibling` is derived from the parent's live child order so it reflects
  // reparent/insertBefore mutations exactly as the browser would.
  Object.defineProperty(el, 'parentNode', {
    get() {
      return el.parentElement;
    },
    enumerable: true,
    configurable: true,
  });
  Object.defineProperty(el, 'nextSibling', {
    get() {
      const p = el.parentElement;
      if (!p) return null;
      const i = p.children.indexOf(el);
      return i >= 0 && i + 1 < p.children.length ? p.children[i + 1] : null;
    },
    enumerable: true,
    configurable: true,
  });
  return el;
}

/** A container FakeEl with a nonzero clientWidth so `width` defaults resolve. */
export function makeContainer(clientWidth = 800, clientHeight = 600): FakeEl {
  const c = makeEl('div');
  c.clientWidth = clientWidth;
  c.clientHeight = clientHeight;
  return c;
}

/** Install a recording document + window + ResizeObserver into globals.
 *  Returns the last-constructed ResizeObserver callback so a test can fire a
 *  synthetic resize. Call `vi.unstubAllGlobals()` in afterEach. */
export function installDom(): { resizeCb: () => (() => void) | undefined } {
  let lastResizeCb: (() => void) | undefined;
  vi.stubGlobal('document', { createElement: (t: string) => makeEl(t) });
  vi.stubGlobal('window', { devicePixelRatio: 1 });
  vi.stubGlobal(
    'ResizeObserver',
    class {
      constructor(cb: () => void) {
        lastResizeCb = cb;
      }
      observe(): void {}
      disconnect(): void {}
    },
  );
  return { resizeCb: () => lastResizeCb };
}

export interface RenderCall {
  page: number;
  /** The per-call `width` (px) the viewer passed — asserted by T3 to confirm
   *  each page gets its OWN px width (uniform px-per-pt scale, §7). */
  width?: number;
  currentDate?: Date | number;
  /** §17.13.5 — the tracked-change view flag the viewer passed (absent =
   *  final view), so the option tests can pin the render-path threading. */
  showTrackedChanges?: boolean;
  /** The canvas element the viewer handed to `renderPage` (main mode). The
   *  flicker-free double-buffer settle renders into a SPARE canvas, so this lets
   *  a test confirm the on-screen canvas was NOT the render target until swap. */
  canvas?: FakeEl;
  resolve: () => void;
  reject: (e: Error) => void;
}

/** A fake DocxDocument covering exactly the surface DocxScrollViewer consumes.
 *  Deferred mode: renderPage / renderPageToBitmap return promises the test
 *  resolves via the recorded RenderCall, so coalescing / stale-drop is
 *  observable. */
export class FakeDocxEngine {
  destroyed = false;
  renderCalls: RenderCall[] = [];
  bitmapCalls: RenderCall[] = [];
  createdBitmaps: Array<{ width: number; height: number; close: ReturnType<typeof vi.fn> }> = [];
  elementContext: DocxElementContext | null = null;
  elementContextCalls: Array<{
    pageIndex: number;
    point: DocxPagePoint;
    options: DocxElementContextOptions;
  }> = [];
  /** Runs fed to the `onTextRun` callback of `renderPage` (main) /
   *  `renderPageToBitmap` (worker) and returned by `collectPageRuns`, so a test
   *  can drive the viewer's per-slot text overlay (buildDocxTextLayer) and find
   *  scan without a real render. Empty/undefined ⇒ no runs emitted. IX6 — the
   *  worker path now ships runs back beside the bitmap, so the stub mirrors that
   *  by replaying `feedTextRuns` to `renderPageToBitmap`'s `onTextRun` too. */
  feedTextRuns?: DocxTextRunInfo[];
  constructor(
    private _pageCount: number,
    // Uniform-page convention: a single-element `_sizes` array means EVERY page
    // is that size (pageSize clamps the index). T2 variable-height tests pass a
    // full per-page array instead.
    private _sizes: Array<{ widthPt: number; heightPt: number }>,
    private _mode: 'main' | 'worker' = 'main',
    private deferred = false,
  ) {}
  get pageCount(): number {
    return this._pageCount;
  }
  /** Mirrors the real `DocxDocument.mode` fact (document.ts) — the exact fact the
   *  viewer constructor reads to decide the render path (main ⇒ renderPage,
   *  worker ⇒ renderPageToBitmap). Design §11: no probing / no silent mis-pathing. */
  get mode(): 'main' | 'worker' {
    return this._mode;
  }
  pageSize(i: number): { widthPt: number; heightPt: number } {
    const clamped = Math.max(0, Math.min(i, this._sizes.length - 1));
    const s = this._sizes[clamped] ?? { widthPt: 0, heightPt: 0 };
    return { widthPt: s.widthPt, heightPt: s.heightPt };
  }
  renderPage(_canvas: unknown, page: number, opts?: RenderPageOptions): Promise<void> {
    if (this._mode === 'worker') {
      // Record the attempt BEFORE throwing so a test can assert the viewer never
      // routes a worker-mode engine through renderPage (and detect wrong-path
      // routing). The real DocxDocument.renderPage is a non-async method that
      // THROWS synchronously in worker mode — mirror that exactly (do NOT return
      // Promise.reject), reusing the real error text (document.ts).
      this.renderCalls.push({ page, resolve: () => {}, reject: () => {} });
      throw new Error(
        "renderPage(canvas) is unavailable in mode: 'worker'; use renderPageToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    // Mirror the real renderer's SYNCHRONOUS device-buffer clear: the real
    // renderDocumentToCanvas sets `canvas.width = round(cssWidth × dpr)` (which
    // clears the backing store to blank) up front, BEFORE its first await, then
    // paints after. The flicker-free settle path relies on this happening on a
    // SPARE off-screen canvas (never the on-screen one), so recording it lets the
    // double-buffer test assert the clear landed on the spare (test d).
    const canvas = _canvas as FakeEl | undefined;
    if (canvas && opts?.width && opts.width > 0) {
      const dpr = opts.dpr ?? 1;
      const size = this.pageSize(page);
      const scale = size.widthPt > 0 ? opts.width / size.widthPt : 0;
      canvas.width = Math.round(opts.width * dpr);
      canvas.height = Math.round(size.heightPt * scale * dpr);
    }
    // Emit any fed runs to the viewer's internal onTextRun so it can build the
    // per-slot overlay (mirrors the real renderer emitting run geometry).
    for (const r of this.feedTextRuns ?? []) opts?.onTextRun?.(r);
    return new Promise<void>((resolve, reject) => {
      const call: RenderCall = {
        page,
        width: opts?.width,
        currentDate: opts?.currentDate,
        showTrackedChanges: opts?.showTrackedChanges,
        canvas,
        resolve: () => resolve(),
        reject,
      };
      this.renderCalls.push(call);
      if (!this.deferred) resolve();
    });
  }
  renderPageToBitmap(
    page: number,
    opts?: WireRenderPageOptions & { onTextRun?: (r: DocxTextRunInfo) => void },
  ): Promise<ImageBitmap> {
    const size = this.pageSize(page);
    const bmp = { width: Math.round(size.widthPt), height: Math.round(size.heightPt), close: vi.fn() };
    this.createdBitmaps.push(bmp);
    // IX6 — replay fed runs to the caller's collector, mirroring the real proxy
    // which invokes `onTextRun` with the runs the worker shipped beside the
    // bitmap. Lets a worker-mode test drive the selection overlay from the stub.
    for (const r of this.feedTextRuns ?? []) opts?.onTextRun?.(r);
    return new Promise<ImageBitmap>((resolve, reject) => {
      const call: RenderCall = {
        page,
        width: opts?.width,
        currentDate: opts?.currentDate,
        showTrackedChanges: opts?.showTrackedChanges,
        resolve: () => resolve(bmp as unknown as ImageBitmap),
        reject,
      };
      this.bitmapCalls.push(call);
      if (!this.deferred) resolve(bmp as unknown as ImageBitmap);
    });
  }
  /** IX6 — mode-agnostic run collection for the find scan (mirrors the real
   *  `DocxDocument.collectPageRuns`). Returns the fed runs regardless of mode. */
  collectPageRuns(_page: number, _opts?: WireRenderPageOptions): Promise<DocxTextRunInfo[]> {
    return Promise.resolve([...(this.feedTextRuns ?? [])]);
  }
  getElementContextAt(
    pageIndex: number,
    point: DocxPagePoint,
    options: DocxElementContextOptions = {},
  ): Promise<DocxElementContext | null> {
    this.elementContextCalls.push({ pageIndex, point, options });
    return Promise.resolve(this.elementContext ? structuredClone(this.elementContext) : null);
  }
  /** The per-call `width` (px) recorded for every renderPage call, in call order.
   *  T3 asserts each mounted page received its OWN px width. */
  renderPageWidths(): number[] {
    return this.renderCalls.map((c) => c.width ?? NaN);
  }
  destroy(): void {
    this.destroyed = true;
  }
  /** IX-nav: a `bookmarkName → page index` table the test seeds so the viewer's
   *  internal `<w:anchor>` hyperlink dispatch resolves against a known map.
   *  Records every lookup so a test can assert which anchor was queried. */
  bookmarkPages = new Map<string, number>();
  bookmarkCalls: string[] = [];
  getBookmarkPage(name: string): number | undefined {
    this.bookmarkCalls.push(name);
    return this.bookmarkPages.get(name);
  }
  /** §17.13.4 — comment data the test seeds so the viewers' comment margin can
   *  be driven without a real parse (mirrors `DocxDocument.comments` and
   *  `DocxDocument.commentAnchorRanges`). Empty by default. */
  comments: import('./types').DocComment[] = [];
  feedCommentAnchorRanges: import('./comment-margin-layout').CommentAnchorRange[] = [];
  commentAnchorRanges(): import('./comment-margin-layout').CommentAnchorRange[] {
    return this.feedCommentAnchorRanges;
  }
  asDoc(): DocxDocument {
    return this as unknown as DocxDocument;
  }
}
