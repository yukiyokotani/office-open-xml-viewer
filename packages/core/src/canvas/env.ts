/**
 * Worker-safe environment guards shared by the renderers.
 *
 * `HTMLCanvasElement` and `window` are not defined inside a Web Worker, so a
 * bare `x instanceof HTMLCanvasElement` (or a bare `devicePixelRatio`
 * identifier) throws a ReferenceError there. Route every render-path check
 * through these helpers so the same code runs on the main thread and in the
 * render worker.
 */

/** True when `target` is a DOM canvas, including one from a popup/iframe realm.
 *  A cross-realm canvas fails `instanceof` against this realm's constructor;
 *  its owning document supplies the correct constructor and FontFaceSet. */
export function isHTMLCanvas(target: unknown): target is HTMLCanvasElement {
  if (typeof HTMLCanvasElement !== 'undefined' && target instanceof HTMLCanvasElement) return true;
  if (typeof target !== 'object' || target === null) return false;
  const element = target as {
    nodeType?: unknown;
    localName?: unknown;
    ownerDocument?: { defaultView?: { HTMLCanvasElement?: unknown } | null };
  };
  if (element.nodeType !== 1 || element.localName !== 'canvas') return false;
  const ownerConstructor = element.ownerDocument?.defaultView?.HTMLCanvasElement;
  return typeof ownerConstructor === 'function' && target instanceof ownerConstructor;
}

/** `window.devicePixelRatio` on the main thread; `fallback` in a worker. */
export function defaultDpr(fallback = 1): number {
  return typeof window !== 'undefined' ? (window.devicePixelRatio || fallback) : fallback;
}
