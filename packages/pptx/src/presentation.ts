import type { MediaElement, Presentation, WorkerRequest, WorkerResponse } from './types';
import { renderSlide } from './renderer';
import { createPresentationHandle, type PresentationHandle } from './presentation-handle';
import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/pptx_parser_bg.wasm?url';

/** Available via Google Fonts; key is lowercase font family name. */
const GOOGLE_FONTS_MAP: Record<string, string> = {
  'nunito sans': 'https://fonts.googleapis.com/css2?family=Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  'nunito': 'https://fonts.googleapis.com/css2?family=Nunito:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  'open sans': 'https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  'roboto': 'https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  'lato': 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  'montserrat': 'https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  'poppins': 'https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  'raleway': 'https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,400;0,700;1,400;1,700&display=swap',
  'playfair display': 'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&display=swap',
};

async function preloadThemeFonts(majorFont: string | null, minorFont: string | null): Promise<void> {
  if (typeof document === 'undefined') return;
  const loaded = new Set<string>();
  const families: string[] = [];
  for (const fontName of [majorFont, minorFont]) {
    if (!fontName) continue;
    const key = fontName.toLowerCase();
    if (loaded.has(key)) continue;
    loaded.add(key);
    families.push(fontName);
    const url = GOOGLE_FONTS_MAP[key];
    if (!url) continue;
    if (document.querySelector(`link[href="${url}"]`)) continue;
    try {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = url;
      document.head.appendChild(link);
    } catch {
      // silently ignore font loading errors
    }
  }
  // @font-face declarations are passive — the font file is only fetched when a
  // glyph is requested. Trigger an explicit load for the weights we care about
  // so the canvas does not measure/draw against a system fallback.
  const loads: Promise<unknown>[] = [];
  for (const family of families) {
    for (const weight of ['400', '700']) {
      for (const style of ['normal', 'italic']) {
        loads.push(
          document.fonts.load(`${style} ${weight} 16px "${family}"`).catch(() => undefined)
        );
      }
    }
  }
  await Promise.race([
    Promise.all(loads).then(() => document.fonts.ready),
    new Promise<void>((resolve) => setTimeout(resolve, 3000)),
  ]);
}

/** Options for {@link PptxPresentation.load}. */
export interface LoadOptions {
  /**
   * Opt in to loading theme-declared webfonts from Google Fonts
   * (`fonts.googleapis.com`). When enabled, end-user IP/User-Agent is sent to
   * Google, which may have privacy/GDPR implications for your application.
   *
   * Default: `false` — the canvas falls back to locally available fonts. Host
   * the required webfonts yourself and reference them via `@font-face` in your
   * application CSS to match the document's theme fonts.
   */
  useGoogleFonts?: boolean;
}

/** Options for rendering a single slide onto a canvas. */
export interface RenderSlideOptions {
  /** Display width in CSS pixels. Defaults to canvas.offsetWidth or 960. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio or 1. */
  dpr?: number;
  /**
   * Skip drawing the play badge overlay on media elements. Used internally by
   * {@link PptxPresentation.presentSlide} so its interactive handle can draw
   * its own play/pause chrome without duplication.
   */
  skipMediaControls?: boolean;
}

/**
 * Headless PPTX rendering engine.
 *
 * Parses `.pptx` archives in a background worker (WASM) but renders slides
 * synchronously on the main thread, so the canvas shares the document's
 * `FontFaceSet` — avoiding subtle wrap differences between system fallback
 * fonts and theme-declared webfonts (e.g. Nunito Sans).
 *
 * Construct via the static `load` factory. A single instance can drive any
 * number of canvases (scroll view, thumbnail grid, master-detail, etc.).
 *
 * @example
 * const pres = await PptxPresentation.load(buffer);
 * await pres.renderSlide(canvas, 0, { width: 960 });
 */
export class PptxPresentation {
  private readonly _worker: Worker;
  private _presentation: Presentation | null = null;
  private _pendingParseCallbacks = new Map<
    number,
    { resolve: (p: Presentation) => void; reject: (e: Error) => void }
  >();
  private _pendingMediaCallbacks = new Map<
    number,
    { resolve: (b: ArrayBuffer) => void; reject: (e: Error) => void }
  >();
  private _mediaCache = new Map<string, Promise<Blob>>();
  private _nextId = 1;
  private _workerReady = false;
  private _workerReadyCallbacks: Array<() => void> = [];

  private constructor() {
    this._worker = new InlineWorker();
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this._worker.postMessage({ kind: 'init', wasmUrl } satisfies WorkerRequest);

    this._worker.onmessage = (e: MessageEvent<WorkerResponse>) => {
      const msg = e.data;

      if (msg.kind === 'ready') {
        this._workerReady = true;
        for (const cb of this._workerReadyCallbacks) cb();
        this._workerReadyCallbacks = [];
        return;
      }

      if (msg.kind === 'parsed') {
        const cb = this._pendingParseCallbacks.get(msg.id);
        if (cb) {
          this._pendingParseCallbacks.delete(msg.id);
          cb.resolve(msg.presentation);
        }
        return;
      }

      if (msg.kind === 'mediaExtracted') {
        const cb = this._pendingMediaCallbacks.get(msg.id);
        if (cb) {
          this._pendingMediaCallbacks.delete(msg.id);
          cb.resolve(msg.bytes);
        }
        return;
      }

      if (msg.kind === 'error') {
        const err = new Error(msg.message);
        const parseCb = this._pendingParseCallbacks.get(msg.id);
        if (parseCb) {
          this._pendingParseCallbacks.delete(msg.id);
          parseCb.reject(err);
          return;
        }
        const mediaCb = this._pendingMediaCallbacks.get(msg.id);
        if (mediaCb) {
          this._pendingMediaCallbacks.delete(msg.id);
          mediaCb.reject(err);
        }
      }
    };
  }

  /** Parse a PPTX from URL or ArrayBuffer. */
  static async load(
    source: string | ArrayBuffer,
    opts: LoadOptions = {},
  ): Promise<PptxPresentation> {
    const pres = new PptxPresentation();
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    await pres._parse(buffer);
    if (opts.useGoogleFonts) {
      await preloadThemeFonts(pres._presentation!.majorFont, pres._presentation!.minorFont);
    }
    return pres;
  }

  private _waitForWorker(): Promise<void> {
    if (this._workerReady) return Promise.resolve();
    return new Promise((resolve) => this._workerReadyCallbacks.push(resolve));
  }

  private async _parse(buffer: ArrayBuffer): Promise<void> {
    await this._waitForWorker();
    const id = this._nextId++;
    const presentation = await new Promise<Presentation>((resolve, reject) => {
      this._pendingParseCallbacks.set(id, { resolve, reject });
      this._worker.postMessage({ kind: 'parse', id, buffer } satisfies WorkerRequest, [buffer]);
    });
    this._presentation = presentation;
  }

  /** Total number of slides in the loaded presentation. */
  get slideCount(): number { return this._presentation?.slides.length ?? 0; }

  /** Slide width in EMU. */
  get slideWidth(): number { return this._presentation?.slideWidth ?? 0; }

  /** Slide height in EMU. */
  get slideHeight(): number { return this._presentation?.slideHeight ?? 0; }

  /** Render a slide onto the given canvas. */
  async renderSlide(
    canvas: HTMLCanvasElement,
    slideIndex: number,
    opts: RenderSlideOptions = {},
  ): Promise<void> {
    if (!this._presentation) throw new Error('Presentation not loaded');
    const slide = this._presentation.slides[slideIndex];
    if (!slide) throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    const dpr = opts.dpr ?? (typeof window !== 'undefined' ? (window.devicePixelRatio || 1) : 1);
    const width = opts.width ?? (canvas.offsetWidth || 960);
    await renderSlide(
      canvas,
      slide,
      this._presentation.slideWidth,
      this._presentation.slideHeight,
      {
        width,
        dpr,
        defaultTextColor: this._presentation.defaultTextColor,
        majorFont: this._presentation.majorFont,
        minorFont: this._presentation.minorFont,
        fetchMedia: (path) => this.getMedia(path),
        skipMediaControls: opts.skipMediaControls,
      },
    );
  }

  /**
   * Extract raw media bytes for a zip path referenced by {@link MediaElement}.
   * Results are cached by path for the lifetime of this instance.
   */
  async getMedia(mediaPath: string): Promise<Blob> {
    const hit = this._mediaCache.get(mediaPath);
    if (hit) return hit;
    const mimeType = this._findMimeTypeForPath(mediaPath);
    const p = (async () => {
      await this._waitForWorker();
      const id = this._nextId++;
      const bytes = await new Promise<ArrayBuffer>((resolve, reject) => {
        this._pendingMediaCallbacks.set(id, { resolve, reject });
        this._worker.postMessage({ kind: 'extractMedia', id, path: mediaPath } satisfies WorkerRequest);
      });
      return new Blob([bytes], { type: mimeType });
    })();
    this._mediaCache.set(mediaPath, p);
    return p;
  }

  private _findMimeTypeForPath(mediaPath: string): string {
    if (!this._presentation) return '';
    for (const slide of this._presentation.slides) {
      for (const el of slide.elements) {
        if (el.type !== 'media') continue;
        const m = el as MediaElement;
        if (m.mediaPath === mediaPath) return m.mimeType;
        if (m.posterPath === mediaPath) return m.posterMimeType;
      }
    }
    return '';
  }

  /**
   * Render a slide and attach canvas-native playback controls for any
   * embedded audio/video. Returns a disposable handle that owns the RAF loop,
   * media elements, and object URLs. Unlike {@link renderSlide}, this method
   * is stateful — always call `handle.dispose()` when leaving the slide.
   */
  async presentSlide(
    canvas: HTMLCanvasElement,
    slideIndex: number,
    opts: RenderSlideOptions = {},
  ): Promise<PresentationHandle> {
    if (!this._presentation) throw new Error('Presentation not loaded');
    const slide = this._presentation.slides[slideIndex];
    if (!slide) throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    const dpr = opts.dpr ?? (typeof window !== 'undefined' ? (window.devicePixelRatio || 1) : 1);
    const width = opts.width ?? (canvas.offsetWidth || 960);
    return createPresentationHandle(canvas, slide, {
      width,
      dpr,
      slideWidthEmu: this._presentation.slideWidth,
      fetchMedia: (path) => this.getMedia(path),
      drawBase: () => this.renderSlide(canvas, slideIndex, { width, dpr, skipMediaControls: true }),
    });
  }

  /** Terminate the worker and release all resources. */
  destroy(): void {
    this._worker.terminate();
    this._presentation = null;
    this._mediaCache.clear();
  }
}
