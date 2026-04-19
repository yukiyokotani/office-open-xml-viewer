import type { Presentation, WorkerRequest, WorkerResponse } from './types';
import { renderSlide } from './renderer';
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

/** Render target: a canvas plus optional sizing hints. */
export interface RenderTarget {
  canvas: HTMLCanvasElement;
  /** Display width in CSS pixels. Defaults to canvas.offsetWidth or 960. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio or 1. */
  dpr?: number;
}

export interface PptxPresentationOptions {
  /** Called when the WASM worker is ready. */
  onReady?: () => void;
  /** Called on parse or render errors. */
  onError?: (err: Error) => void;
}

/**
 * Headless PPTX rendering engine.
 *
 * Parses .pptx files in a background worker (WASM) but renders slides
 * synchronously on the main thread, so text measurement uses the document's
 * FontFaceSet — avoiding subtle wrap differences between system fallback fonts
 * and theme-declared webfonts (e.g. Nunito Sans).
 *
 * @example
 * const pres = new PptxPresentation();
 * await pres.load(buffer);
 * await pres.renderSlide({ canvas, width: 960 }, 0);
 */
export class PptxPresentation {
  private worker: Worker | null = null;
  private presentation: Presentation | null = null;
  private pendingParseCallbacks = new Map<
    number,
    { resolve: (p: Presentation) => void; reject: (e: Error) => void }
  >();
  private nextId = 1;
  private workerReady = false;
  private workerReadyCallbacks: Array<() => void> = [];
  private readonly opts: PptxPresentationOptions;

  constructor(opts: PptxPresentationOptions = {}) {
    this.opts = opts;
    this.initWorker();
  }

  private initWorker() {
    this.worker = new InlineWorker();
    const wasmUrl = new URL(wasmAssetUrl, location.href).href;
    this.worker.postMessage({ kind: 'init', wasmUrl } satisfies WorkerRequest);

    this.worker.onmessage = (e: MessageEvent<WorkerResponse>) => {
      const msg = e.data;

      if (msg.kind === 'ready') {
        this.workerReady = true;
        for (const cb of this.workerReadyCallbacks) cb();
        this.workerReadyCallbacks = [];
        this.opts.onReady?.();
        return;
      }

      if (msg.kind === 'parsed') {
        const cb = this.pendingParseCallbacks.get(msg.id);
        if (cb) {
          this.pendingParseCallbacks.delete(msg.id);
          cb.resolve(msg.presentation);
        }
        return;
      }

      if (msg.kind === 'error') {
        const parseCb = this.pendingParseCallbacks.get(msg.id);
        const err = new Error(msg.message);
        if (parseCb) {
          this.pendingParseCallbacks.delete(msg.id);
          parseCb.reject(err);
        }
        this.opts.onError?.(err);
      }
    };

    this.worker.onerror = (e) => {
      this.opts.onError?.(new Error(e.message));
    };
  }

  private waitForWorker(): Promise<void> {
    if (this.workerReady) return Promise.resolve();
    return new Promise((resolve) => this.workerReadyCallbacks.push(resolve));
  }

  /** Parse a PPTX ArrayBuffer. Resolves when parsing is complete. */
  async load(buffer: ArrayBuffer): Promise<void> {
    await this.waitForWorker();
    const id = this.nextId++;
    const presentation = await new Promise<Presentation>((resolve, reject) => {
      this.pendingParseCallbacks.set(id, { resolve, reject });
      this.worker!.postMessage({ kind: 'parse', id, buffer } satisfies WorkerRequest, [buffer]);
    });
    this.presentation = presentation;
    await preloadThemeFonts(presentation.majorFont, presentation.minorFont);
  }

  /** Total number of slides in the loaded presentation (0 if not loaded). */
  get slideCount(): number { return this.presentation?.slides.length ?? 0; }

  /** Slide width in EMU (0 if not loaded). */
  get slideWidth(): number { return this.presentation?.slideWidth ?? 0; }

  /** Slide height in EMU (0 if not loaded). */
  get slideHeight(): number { return this.presentation?.slideHeight ?? 0; }

  /** Render a slide directly onto a canvas on the main thread. */
  async renderSlide(target: RenderTarget, slideIndex: number): Promise<void> {
    if (!this.presentation) throw new Error('No presentation loaded. Call load() first.');
    const slide = this.presentation.slides[slideIndex];
    if (!slide) throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    const dpr = target.dpr ?? (typeof window !== 'undefined' ? (window.devicePixelRatio || 1) : 1);
    const width = target.width ?? (target.canvas.offsetWidth || 960);
    await renderSlide(
      target.canvas,
      slide,
      this.presentation.slideWidth,
      this.presentation.slideHeight,
      {
        width,
        dpr,
        defaultTextColor: this.presentation.defaultTextColor,
        majorFont: this.presentation.majorFont,
        minorFont: this.presentation.minorFont,
      },
    );
  }

  /** Terminate the worker and release all resources. */
  destroy(): void {
    this.worker?.terminate();
    this.worker = null;
    this.presentation = null;
  }
}
