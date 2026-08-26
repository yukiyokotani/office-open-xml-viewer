import { vi } from 'vitest';

/**
 * A minimal recording fake DOM for exercising `XlsxViewer`'s construct/destroy
 * lifecycle without jsdom (the repo has no jsdom; the scroll-viewer suites use
 * the same hand-rolled-fake approach). It covers exactly the surface the
 * constructor and `destroy()` touch: element creation, child tree, style,
 * attributes/dataset/classList, event-listener recording, a `document.head`
 * that supports `appendChild` + `querySelector('style[data-…]')`, and
 * document-level `addEventListener`/`removeEventListener`/`dispatchEvent` so a
 * test can confirm the keydown listener is detached on destroy.
 *
 * Geometry getters return 0 — the constructor only registers listeners; it does
 * not fire them, so no real layout is needed.
 */
export interface FakeEl {
  tag: string;
  ownerDocument: FakeDocument | null;
  textContent: string;
  title: string;
  // <input>-only fields the zoom slider sets
  type: string;
  min: string;
  max: string;
  step: string;
  value: string;
  style: Record<string, string> & { cssText: string };
  children: FakeEl[];
  parentElement: FakeEl | null;
  readonly parentNode: FakeEl | null;
  readonly nextSibling: FakeEl | null;
  readonly firstChild: FakeEl | null;
  readonly childNodes: FakeEl[];
  dataset: Record<string, string>;
  classList: { add(...c: string[]): void; remove(...c: string[]): void; contains(c: string): boolean };
  _attrs: Map<string, string>;
  _listeners: Map<string, Array<(e: unknown) => void>>;
  /** Persisted bitmaprenderer context (worker mode). Records the last painted
   *  bitmap so tests can assert which frame reached the canvas. Lazily created
   *  by {@link getContext}('bitmaprenderer'); null until then. */
  _bitmapCtx: { lastBitmap: unknown; transferFromImageBitmap(bmp: unknown): void } | null;
  // geometry (constructor reads a few; 0 is fine — no events are fired)
  scrollTop: number;
  scrollLeft: number;
  clientWidth: number;
  clientHeight: number;
  clientLeft: number;
  clientTop: number;
  scrollWidth: number;
  scrollHeight: number;
  offsetLeft: number;
  offsetWidth: number;
  offsetHeight: number;
  width: number;
  height: number;
  appendChild(c: FakeEl): FakeEl;
  replaceChildren(...children: FakeEl[]): void;
  removeChild(c: FakeEl): FakeEl;
  remove(): void;
  insertBefore(n: FakeEl, ref: FakeEl | null): FakeEl;
  setAttribute(name: string, value: string): void;
  removeAttribute(name: string): void;
  getAttribute(name: string): string | null;
  hasAttribute(name: string): boolean;
  contains(other: FakeEl | null): boolean;
  querySelector(sel: string): FakeEl | null;
  addEventListener(type: string, fn: (e: unknown) => void, opts?: unknown): void;
  removeEventListener(type: string, fn: (e: unknown) => void, opts?: unknown): void;
  setPointerCapture(id: number): void;
  releasePointerCapture(id: number): void;
  getContext(kind: string): unknown;
  getBoundingClientRect(): { top: number; left: number; width: number; height: number; bottom: number; right: number };
  /** test-only: fire a recorded listener */
  dispatch(type: string, event?: unknown): void;
}

/** Depth-first match of `style[data-<attr>]`-style selectors and bare tag/attr
 *  combos used by the viewer. Supports `tag[data-foo]` and `[data-foo]`. */
function matchSelector(el: FakeEl, sel: string): boolean {
  const m = sel.match(/^([a-zA-Z]*)(?:\[([^\]=]+)(?:=["']?([^"'\]]*)["']?)?\])?$/);
  if (!m) return false;
  const [, tag, attr, val] = m;
  if (tag && el.tag !== tag) return false;
  if (attr) {
    if (!el._attrs.has(attr)) return false;
    if (val !== undefined && el._attrs.get(attr) !== val) return false;
  }
  return true;
}

function queryDeep(root: FakeEl, sel: string): FakeEl | null {
  for (const c of root.children) {
    if (matchSelector(c, sel)) return c;
    const found = queryDeep(c, sel);
    if (found) return found;
  }
  return null;
}

export function makeEl(tag: string, ownerDocument: FakeDocument | null = null): FakeEl {
  const style: Record<string, string> = {};
  const el: FakeEl = {
    tag,
    ownerDocument,
    textContent: '',
    title: '',
    type: '',
    min: '',
    max: '',
    step: '',
    value: '',
    children: [],
    parentElement: null,
    parentNode: null,
    nextSibling: null,
    firstChild: null,
    childNodes: [],
    dataset: {},
    _attrs: new Map(),
    _listeners: new Map(),
    _bitmapCtx: null,
    scrollTop: 0,
    scrollLeft: 0,
    clientWidth: 0,
    clientHeight: 0,
    clientLeft: 0,
    clientTop: 0,
    scrollWidth: 0,
    scrollHeight: 0,
    offsetLeft: 0,
    offsetWidth: 0,
    offsetHeight: 0,
    width: 0,
    height: 0,
    classList: {
      add() {},
      remove() {},
      contains() {
        return false;
      },
    },
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
      // Real-DOM move semantics: detach from the current parent first.
      c.parentElement?.removeChild(c);
      c.parentElement = this;
      this.children.push(c);
      return c;
    },
    replaceChildren(...children: FakeEl[]) {
      for (const child of [...this.children]) this.removeChild(child);
      for (const child of children) this.appendChild(child);
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
      // of this node throws NotFoundError (kept consistent with the pptx/docx
      // test DOMs so no fake is more permissive than a browser).
      if (ref && !this.children.includes(ref)) {
        throw new Error('NotFoundError: the node before which the new node is to be inserted is not a child of this node');
      }
      n.parentElement?.removeChild(n);
      n.parentElement = this;
      const i = ref ? this.children.indexOf(ref) : -1;
      if (i >= 0) this.children.splice(i, 0, n);
      else this.children.push(n);
      return n;
    },
    setAttribute(name: string, value: string) {
      this._attrs.set(name, value);
    },
    removeAttribute(name: string) {
      this._attrs.delete(name);
    },
    getAttribute(name: string) {
      return this._attrs.has(name) ? (this._attrs.get(name) as string) : null;
    },
    hasAttribute(name: string) {
      return this._attrs.has(name);
    },
    contains(other: FakeEl | null) {
      if (!other) return false;
      if (other === this) return true;
      return this.children.some((c) => c.contains(other));
    },
    querySelector(sel: string) {
      return queryDeep(this, sel);
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
    setPointerCapture() {},
    releasePointerCapture() {},
    getContext(kind: string) {
      if (kind === 'bitmaprenderer') {
        // Persist one bitmaprenderer context per canvas (a real canvas holds a
        // single context type for its lifetime) and record the last painted
        // bitmap, so worker-mode tests can assert which frame reached the canvas.
        this._bitmapCtx ??= {
          lastBitmap: null,
          transferFromImageBitmap(bmp: unknown) {
            this.lastBitmap = bmp;
          },
        };
        return this._bitmapCtx;
      }
      return {};
    },
    getBoundingClientRect() {
      return { top: 0, left: 0, width: this.clientWidth, height: this.clientHeight, bottom: 0, right: 0 };
    },
    dispatch(type: string, event: unknown = {}) {
      for (const fn of this._listeners.get(type) ?? []) fn(event);
    },
  };
  // Mirror the DOM: clearing textContent removes all descendants. Overlay hosts
  // rely on this when switching between cell and element selection frames.
  let text = '';
  Object.defineProperty(el, 'textContent', {
    get() { return text; },
    set(value: string) {
      text = value;
      if (value === '') {
        for (const child of el.children) child.parentElement = null;
        el.children.length = 0;
      }
    },
    enumerable: true,
    configurable: true,
  });
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
  Object.defineProperty(el, 'firstChild', {
    get() {
      return el.children[0] ?? null;
    },
    enumerable: true,
    configurable: true,
  });
  Object.defineProperty(el, 'childNodes', {
    get() {
      return el.children;
    },
    enumerable: true,
    configurable: true,
  });
  // dataset writes mirror into attributes as data-* so querySelector can see them.
  const dataProxy = new Proxy(
    {},
    {
      set(target: Record<string, string>, prop: string, value: string) {
        target[prop] = value;
        el._attrs.set('data-' + prop.replace(/[A-Z]/g, (m) => '-' + m.toLowerCase()), value);
        return true;
      },
      get(target: Record<string, string>, prop: string) {
        return target[prop];
      },
    },
  );
  Object.defineProperty(el, 'dataset', {
    get() {
      return dataProxy;
    },
    enumerable: true,
    configurable: true,
  });
  return el;
}

/** A recording `document.head` that supports appendChild + querySelector. */
export interface FakeDocument {
  head: FakeEl;
  defaultView: {
    devicePixelRatio: number;
    ResizeObserver: new (callback: () => void) => {
      observe(): void;
      disconnect(): void;
      unobserve(): void;
    };
  };
  createElement(tag: string): FakeEl;
  createElementNS(namespace: string, tag: string): FakeEl;
  addEventListener(type: string, fn: (e: unknown) => void, opts?: unknown): void;
  removeEventListener(type: string, fn: (e: unknown) => void, opts?: unknown): void;
  dispatchEvent(type: string, event?: unknown): void;
  /** test-only: current document-level listener count for a type */
  listenerCount(type: string): number;
}

/** Install a recording document + window + ResizeObserver into globals.
 *  Returns the fake document so a test can query head / dispatch keydown.
 *  Call `vi.unstubAllGlobals()` in afterEach. */
export function makeDocument(devicePixelRatio = 1): FakeDocument {
  const docListeners = new Map<string, Array<(e: unknown) => void>>();
  const ResizeObserver = class {
    constructor(_callback: () => void) {}
    observe(): void {}
    disconnect(): void {}
    unobserve(): void {}
  };
  let doc: FakeDocument;
  doc = {
    head: null as unknown as FakeEl,
    defaultView: { devicePixelRatio, ResizeObserver },
    createElement: (t: string) => makeEl(t, doc),
    createElementNS: (_namespace: string, t: string) => makeEl(t, doc),
    addEventListener(type: string, fn: (e: unknown) => void) {
      const arr = docListeners.get(type) ?? [];
      arr.push(fn);
      docListeners.set(type, arr);
    },
    removeEventListener(type: string, fn: (e: unknown) => void) {
      const arr = docListeners.get(type);
      if (arr) docListeners.set(type, arr.filter((f) => f !== fn));
    },
    dispatchEvent(type: string, event: unknown = {}) {
      for (const fn of docListeners.get(type) ?? []) fn(event);
    },
    listenerCount(type: string) {
      return docListeners.get(type)?.length ?? 0;
    },
  };
  doc.head = makeEl('head', doc);
  return doc;
}

export function installDom(): FakeDocument {
  const doc = makeDocument();
  vi.stubGlobal('document', doc);
  vi.stubGlobal('window', doc.defaultView);
  vi.stubGlobal('ResizeObserver', doc.defaultView.ResizeObserver);
  return doc;
}

/** A container FakeEl with nonzero client size so width defaults resolve. */
export function makeContainer(
  clientWidth = 800,
  clientHeight = 600,
  ownerDocument: FakeDocument | null = null,
): FakeEl {
  const c = makeEl('div', ownerDocument);
  c.clientWidth = clientWidth;
  c.clientHeight = clientHeight;
  return c;
}
