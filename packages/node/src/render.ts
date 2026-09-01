/**
 * Server-side rendering helpers. These adapt the browser-bound canvas
 * renderers in `@silurus/ooxml-{pptx,docx,xlsx}` to a user-supplied
 * Node canvas implementation (e.g. `skia-canvas`).
 *
 * The original renderers reach for a few browser-only globals:
 *   - `createImageBitmap` — used to paint embedded raster pictures
 *   - `document.fonts.add(new FontFace(...))` — used by the Google-Fonts loader
 *
 * We provide minimal Node shims for bitmap decode and auxiliary canvases. The
 * shared SVG loader detects the absence of `Image` and uses this same bitmap
 * decoder. Font loading stays opt-in
 * via `useGoogleFonts: false`. The user passes in a canvas factory so the
 * package itself does not pin a particular Node canvas implementation;
 * `skia-canvas` is recommended in the README.
 */

import type { Presentation } from '@silurus/ooxml-pptx';

let nodeCanvasRuntimeTail: Promise<void> = Promise.resolve();

/** DOM-free text metrics required from a Node canvas implementation. */
export interface NodeTextMetricsLike {
  readonly width: number;
}

/** DOM-free structural seam for the Node canvas 2D context. */
export interface NodeCanvasContext2D {
  measureText(text: string): NodeTextMetricsLike;
  /** Optional here for lightweight metric-only adapters; required by the
   * image shim when it must materialize a resized fallback surface. */
  drawImage?(image: NodeImageLike, dx: number, dy: number, dw: number, dh: number): void;
}

/** A subset of the Node-canvas API that the renderers actually need. The
 *  `skia-canvas` `Canvas` (and `@napi-rs/canvas`'s `Canvas`) both satisfy
 *  this — they expose the same `getContext('2d')` shape as the browser. */
export interface NodeCanvasLike {
  width: number;
  height: number;
  getContext(kind: '2d'): NodeCanvasContext2D;
  /** Encode to PNG bytes. skia-canvas: `canvas.png` async getter or
   *  `toBuffer('png')`. @napi-rs/canvas: `toBuffer('image/png')`. */
  toBuffer?(format?: string): Uint8Array | Promise<Uint8Array>;
}

export interface NodeImageLike {
  width: number;
  height: number;
}

export interface NodeCanvasFactory {
  /** Create a blank canvas of the given pixel size. */
  createCanvas(width: number, height: number): NodeCanvasLike;
  /** Decode a buffer (PNG/JPEG/etc.) into something the canvas can `drawImage`.
   * Implementations should honor `target` in the decoder when possible; the
   * shim falls back to a target-sized canvas when they return native pixels. */
  loadImage(
    buffer: ArrayBuffer | Uint8Array,
    target?: Readonly<{ width?: number; height?: number }>,
  ): Promise<NodeImageLike>;
}

/**
 * Polyfill `globalThis.OffscreenCanvas` so the shared rendering primitives can
 * allocate auxiliary canvases under Node.
 *
 * `packages/core/src/shape/effects.ts`'s `createAuxCanvas` probes
 * `typeof OffscreenCanvas !== 'undefined'` and otherwise `document`. Under Node
 * neither exists, so it returns `null` and the pptx renderer's beveled-flat
 * path, scene3d projection, and the inner-shadow / soft-edge / reflection
 * effects all *silently* degrade to flat output (no rim shading, no 3D warp,
 * no blur). This shim makes `new OffscreenCanvas(w, h)` allocate a real
 * skia-canvas (via the user's `factory.createCanvas`), so those paths light up
 * server-side exactly as they do in the browser.
 *
 * `new OffscreenCanvas(w, h)` simply *returns* a backing canvas from
 * `factory.createCanvas` (a class constructor may return a different object).
 * Returning the real canvas — rather than a wrapper — matters: the effect
 * helpers pass the allocated canvas straight to `ctx.drawImage(aux, …)`, and
 * skia-canvas's `drawImage` only accepts a real `Image`/`Canvas`, rejecting a
 * forwarding wrapper. The backing canvas already exposes everything
 * `createAuxCanvas` consumers touch (getContext, getImageData/putImageData,
 * drawImage as a source, and `ctx.filter = 'blur(Npx)'`); skia-canvas supports
 * all of them, including the `filter` blur used by the effect helpers.
 *
 * If `globalThis.OffscreenCanvas` is already defined (e.g. on a real DOM/worker
 * runtime, or Node ≥ a future version that ships it) the existing value is left
 * untouched. The returned function restores the global to its pre-call value.
 */
export function installOffscreenCanvasShim(factory: NodeCanvasFactory): () => void {
  const g = globalThis as unknown as { OffscreenCanvas?: unknown };
  const prev = g.OffscreenCanvas;
  const hadOwn = Object.prototype.hasOwnProperty.call(globalThis, 'OffscreenCanvas');

  // Respect a pre-existing implementation — never overwrite a real one.
  if (typeof prev !== 'undefined') {
    return () => {
      /* nothing to restore: we never touched the global */
    };
  }

  class OffscreenCanvasShim {
    constructor(width: number, height: number) {
      // Return the backing canvas itself (constructors may return another
      // object). This keeps `aux instanceof <skia Canvas>` true so skia's
      // `drawImage` accepts it as an image source.
      return factory.createCanvas(width, height) as unknown as OffscreenCanvasShim;
    }
  }

  g.OffscreenCanvas = OffscreenCanvasShim as unknown;

  return () => {
    if (hadOwn) {
      g.OffscreenCanvas = prev;
    } else {
      delete g.OffscreenCanvas;
    }
  };
}

/** Polyfill `globalThis.createImageBitmap` so the existing renderers can
 *  decode raster pictures. Wires it to the user's `loadImage`. Returns the
 *  previous global (if any) so the caller can restore it. */
export function installImageBitmapShim(factory: NodeCanvasFactory): () => void {
  const g = globalThis as unknown as { createImageBitmap?: unknown };
  const prev = g.createImageBitmap;
  // The source may be raw bytes (the raster-decode path) OR a canvas-like
  // surface with a 2D context (core's `applyDuotone`, §20.1.8.23, which bakes a
  // recoloured offscreen surface back into an "ImageBitmap"). Widen the param so
  // the canvas branch is a real member and needs no double-cast.
  type CanvasLike = { getContext(id: '2d'): unknown };
  g.createImageBitmap = async (
    source: Blob | ArrayBuffer | Uint8Array | CanvasLike,
    options?: ImageBitmapOptions,
  ) => {
    // A canvas-like source is already a drawable image source in node-canvas —
    // return it directly. The surface IS the skia Canvas the factory made, so no
    // byte round-trip is needed (and skia has no `createImageBitmap(canvas)`).
    if (source && typeof (source as CanvasLike).getContext === 'function') {
      return source as CanvasLike;
    }
    let buf: ArrayBuffer | Uint8Array;
    if (source instanceof Uint8Array || source instanceof ArrayBuffer) {
      buf = source;
    } else if (typeof (source as Blob).arrayBuffer === 'function') {
      buf = await (source as Blob).arrayBuffer();
    } else {
      throw new Error('createImageBitmap shim: unsupported source type');
    }
    const resizeWidth = options?.resizeWidth;
    const resizeHeight = options?.resizeHeight;
    const target = typeof resizeWidth === 'number' && resizeWidth > 0
      ? { width: resizeWidth }
      : typeof resizeHeight === 'number' && resizeHeight > 0
        ? { height: resizeHeight }
        : undefined;
    const image = await factory.loadImage(buf, target);
    if (!target) return image;
    if (target.width !== undefined && image.width === target.width) return image;
    if (target.height !== undefined && image.height === target.height) return image;
    const scale = target.width !== undefined
      ? target.width / image.width
      : typeof target.height === 'number'
        ? target.height / image.height
        : 1;
    const width = Math.max(1, Math.round(image.width * scale));
    const height = Math.max(1, Math.round(image.height * scale));
    const surface = factory.createCanvas(width, height);
    const context = surface.getContext('2d');
    if (typeof context.drawImage !== 'function') return image;
    context.drawImage(image, 0, 0, width, height);
    (image as NodeImageLike & { close?: () => void }).close?.();
    return surface;
  };
  return () => { g.createImageBitmap = prev as typeof globalThis.createImageBitmap; };
}

/**
 * Run one renderer while this process owns the browser-canvas compatibility
 * globals. Node globals are process-wide, so DOCX and PPTX operations share one
 * queue and always restore exactly the values they observed.
 *
 * @internal
 */
export function withNodeCanvasRuntime<T>(
  factory: NodeCanvasFactory,
  operation: () => Promise<T>,
): Promise<T> {
  const run = async (): Promise<T> => {
    const restoreImageBitmap = typeof globalThis.createImageBitmap === 'function'
      ? () => undefined
      : installImageBitmapShim(factory);
    const restoreOffscreen = installOffscreenCanvasShim(factory);
    try {
      return await operation();
    } finally {
      restoreOffscreen();
      restoreImageBitmap();
    }
  };
  const result = nodeCanvasRuntimeTail.then(run, run);
  nodeCanvasRuntimeTail = result.then(() => undefined, () => undefined);
  return result;
}

/** Render a materialized slide into a user-supplied Node canvas. The
 *  caller must:
 *   - provide a `Presentation` obtained from `materializePptxPresentation()`,
 *     or use `openPptxPresentation().renderSlide()` when embedded parts matter
 *   - pass `opts.factory` (recommended) or install a compatible
 *     `createImageBitmap` implementation yourself
 *   - load fonts they want available into the canvas implementation's font
 *     registry (e.g. `Font.use(...)` for skia-canvas) BEFORE calling render
 *
 *  Resolves after painting the caller-owned canvas; encode it with the canvas
 *  implementation's PNG API (for example `canvas.toBuffer('png')`).
 *
 *  Note: the underlying browser renderer is `async` and imports Vite-only
 *  worker assets at the top of `presentation.ts`. The Node path bypasses
 *  `PptxPresentation` and `worker.ts` entirely and calls the pure
 *  `renderSlide` function from `@silurus/ooxml-pptx`.
 *
 *  Pass `opts.factory` so this function can install both image-decoding and
 *  auxiliary-canvas shims for the duration of the render. Without a factory,
 *  the caller must provide `createImageBitmap` itself and effects that require
 *  auxiliary canvases can degrade to flat output. Shared Node renders are
 *  serialized while the process-wide shims are installed, then the previous
 *  globals are restored. */
export async function renderSlideNode(
  canvas: NodeCanvasLike,
  presentation: Presentation,
  slideIndex: number,
  opts: {
    width?: number;
    dpr?: number;
    factory?: NodeCanvasFactory;
    /**
     * Lazily resolve an embedded image (by zip path + MIME) to a Blob. Pictures
     * and blip fills carry only zip paths now (no inlined base64). Supplying
     * Use the owned presentation session when bytes must come from the source
     * package. When omitted, images decode to nothing.
     */
    fetchImage?: (path: string, mimeType: string) => Promise<Blob>;
    /** Lazily resolve embedded media/poster bytes. Defaults to an empty Blob. */
    fetchMedia?: (path: string) => Promise<Blob>;
  } = {},
): Promise<void> {
  // Direct import of the pure renderer module — avoids `presentation.ts`
  // and `viewer.ts`, both of which pull Vite-specific worker / asset
  // imports that don't resolve under Node.
  const { renderSlide } = (await import('@silurus/ooxml-pptx/internal/session')) as unknown as {
    renderSlide: (
      canvas: HTMLCanvasElement,
      slide: Presentation['slides'][number],
      slideWidth: number,
      slideHeight: number,
      opts: Record<string, unknown>,
    ) => Promise<HTMLCanvasElement>;
  };
  const slide = presentation.slides[slideIndex];
  if (!slide) throw new Error(`Slide index ${slideIndex} out of range`);
  const width = opts.width ?? 960;
  const dpr = opts.dpr ?? 2;
  // Light up bevel/scene3d/effects auxiliary-canvas allocation for the render,
  // then restore the global. No-op restore if a factory was not supplied or an
  // OffscreenCanvas already exists.
  const fetchImage = opts.fetchImage ?? (async () => new Blob([]));
  const paint = async (): Promise<void> => {
    await renderSlide(
      canvas as unknown as HTMLCanvasElement,
      slide,
      presentation.slideWidth,
      presentation.slideHeight,
      {
        width,
        dpr,
        defaultTextColor: presentation.defaultTextColor,
        majorFont: presentation.majorFont,
        minorFont: presentation.minorFont,
        hlinkColor: presentation.hlinkColor ?? null,
        // Node-side renderers don't run media playback, so an empty fetcher
        // is fine for posters.
        fetchMedia: opts.fetchMedia ?? (async () => new Blob([])),
        // Owned sessions provide the package-backed fetcher. Direct callers may
        // inject another source; text/shape-only rendering uses an empty Blob.
        fetchImage,
        skipMediaControls: true,
      },
    );
  };
  await (opts.factory ? withNodeCanvasRuntime(opts.factory, paint) : paint());
}
