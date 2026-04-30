import type { MediaElement, Presentation, WorkerRequest, WorkerResponse } from './types';
import { renderSlide, type TextRunCallback } from './renderer';
import { createPresentationHandle, type PresentationHandle } from './presentation-handle';
import { preloadGoogleFonts, type FontPreloadEntry, type LoadOptions as CoreLoadOptions } from '@silurus/ooxml-core';
import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/pptx_parser_bg.wasm?url';

/** Theme-referenced typefaces commonly used by PPTX templates. Keys are
 *  lower-cased family names; loadFamily is omitted because Google Fonts
 *  ships the same family name in those entries. */
const PPTX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  'nunito sans':       { url: 'https://fonts.googleapis.com/css2?family=Nunito+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'nunito':            { url: 'https://fonts.googleapis.com/css2?family=Nunito:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'open sans':         { url: 'https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'roboto':            { url: 'https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'lato':              { url: 'https://fonts.googleapis.com/css2?family=Lato:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'montserrat':        { url: 'https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'poppins':           { url: 'https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'raleway':           { url: 'https://fonts.googleapis.com/css2?family=Raleway:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
  'playfair display':  { url: 'https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400;1,700&display=swap' },
};

/** Options for {@link PptxPresentation.load}. Re-exports the shared
 *  `LoadOptions` shape from `@silurus/ooxml-core`. */
export type LoadOptions = CoreLoadOptions;

/** Options for rendering a single slide onto a canvas. */
export interface RenderSlideOptions {
  /** Display width in CSS pixels. Defaults to canvas.offsetWidth or 960. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio or 1. */
  dpr?: number;
  /** Called for each rendered text segment. Used to build a transparent text selection overlay. */
  onTextRun?: TextRunCallback;
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
      await preloadGoogleFonts(
        [pres._presentation!.majorFont, pres._presentation!.minorFont],
        PPTX_GOOGLE_FONTS,
      );
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
      opts.onTextRun,
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
      drawBase: () => this.renderSlide(canvas, slideIndex, {
        width,
        dpr,
        skipMediaControls: true,
        onTextRun: opts.onTextRun,
      }),
    });
  }

  /** Terminate the worker and release all resources. */
  destroy(): void {
    this._worker.terminate();
    this._presentation = null;
    this._mediaCache.clear();
  }
}
