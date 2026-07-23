import type { DimOptions, MediaElement, Presentation, ShapeElement, WorkerRequest, WorkerResponse } from './types';
import { renderSlide, dropImageBitmapCache, dropDuotoneBitmapCache, type TextRunCallback, type PptxTextRunInfo } from './renderer';
import { createPresentationHandle, type PresentationHandle } from './presentation-handle';
import { selectNotes } from './notes';
import { selectHidden } from './hidden';
import {
  buildSlidePartIndex,
  resolveInternalSlideTarget,
  type SlidePartNames,
} from './slide-nav';
import {
  preloadGoogleFonts,
  unloadGoogleFonts,
  WorkerBridge,
  defaultDpr,
  isHTMLCanvas,
  dropSvgImageCache,
  resolveOoxmlContainer,
  toArrayBuffer,
  type LoadOptions as CoreLoadOptions,
  type MathRenderer,
} from '@silurus/ooxml-core';
import { PPTX_GOOGLE_FONTS, pptxFontPreloadNames } from './google-fonts';
import { findMimeTypeForPath } from './media-mime';
import type {
  PresentationMeta,
  RenderWorkerRequest,
  RenderWorkerResponse,
} from './worker-protocol';
import InlineWorker from './worker.ts?worker&inline';
import wasmAssetUrl from './wasm/pptx_parser_bg.wasm?url';
import {
  applyPptxShapeChanges,
  type PptxApplyShapeChangesRequest,
  type PptxShapeChange,
} from './shape-changes';

/** Options for {@link PptxPresentation.load}. */
export type LoadOptions = CoreLoadOptions & {
  /**
   * 'main' (default): parse in a worker, render on the main thread (current
   * behaviour). 'worker': parse AND render inside the worker; use
   * {@link PptxPresentation.renderSlideToBitmap} and paint the returned
   * ImageBitmap via an `ImageBitmapRenderingContext`. Requires OffscreenCanvas.
   */
  mode?: 'main' | 'worker';
};

/** Options for {@link PptxPresentation.renderSlideToBitmap}. */
export interface RenderSlideToBitmapOptions {
  /** Slide width in CSS pixels. Defaults to 960. */
  width?: number;
  /** Device pixel ratio. Defaults to window.devicePixelRatio (workers have none). */
  dpr?: number;
  /**
   * Skip the static media play-badge so a live overlay can draw its own
   * controls. Used internally by {@link PptxPresentation.presentSlide}.
   * @internal
   */
  skipMediaControls?: boolean;
  /** Translucent overlay drawn over the finished slide (hidden-slide dimming). */
  dim?: DimOptions;
  /**
   * IX6 — receives the slide's text-run geometry (the same stream `renderSlide`
   * emits in main mode). Stays main-thread (never crosses the wire); in worker
   * mode the proxy invokes it with the runs the worker shipped back beside the
   * bitmap, so a caller builds the selection / find overlay on the SAME code
   * path in both modes.
   */
  onTextRun?: TextRunCallback;
}

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
  /** Translucent overlay drawn over the finished slide (hidden-slide dimming). */
  dim?: DimOptions;
}

/** Result of a successful {@link PptxPresentation.applyShapeChanges} call. */
export interface PptxShapeChangesResult {
  /** Zero-based index of the only slide whose model changed. */
  slideIndex: number;
  /** Slide-local DrawingML `cNvPr@id` of the updated shape. */
  shapeId: string;
  /**
   * Detached snapshot of the committed shape. Mutating this object does not
   * mutate the presentation; call `applyShapeChanges` for another change.
   */
  shape: ShapeElement;
  /** Detached, normalized copies of the changes that were applied. */
  applied: PptxShapeChange[];
  /**
   * Detached changes that restore the previous shape. Pass this array directly
   * to `applyShapeChanges` to undo the batch.
   */
  inverse: PptxShapeChange[];
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
  private readonly _bridge: WorkerBridge<WorkerResponse | RenderWorkerResponse>;
  private _mode: 'main' | 'worker' = 'main';
  private _presentation: Presentation | null = null;
  private _meta: PresentationMeta | null = null;
  /** Lazily-built `partName → slide index` map for internal hyperlink slide
   *  jumps (IX-nav). Cleared on {@link destroy}; built on first
   *  {@link getSlideIndexByPartName}/{@link resolveInternalTarget} from either
   *  the parsed slides (main) or the worker meta's `partNames` (worker). */
  private _slidePartIndex: Map<string, number> | null = null;
  private _mediaCache = new Map<string, Promise<Blob>>();
  private _imageCache = new Map<string, Promise<Blob>>();
  /** Google-Fonts `FontFace` objects this deck preloaded into `document.fonts`
   *  (main mode only — in worker mode the worker owns them and terminates with
   *  its own FontFaceSet). Released in {@link destroy} so they do not leak into
   *  the shared FontFaceSet for the lifetime of the SPA (deduped + refcounted in
   *  core, so a web font shared with another open deck survives until both go). */
  private _googleFontFaces: FontFace[] = [];
  /** One stable closure per instance: the decoded-bitmap and SVG caches key on
   *  this identity to scope decodes per deck (so two open decks never swap
   *  images for a shared zip path like ppt/media/image1.png). Reusing the same
   *  reference across every render also lets those caches hit across slides. */
  private readonly _fetchImage = (path: string, mime: string): Promise<Blob> =>
    this.getImage(path, mime);
  /** Opt-in OMML equation engine, injected once at {@link load}. Every
   *  `renderSlide` / `presentSlide` reuses it — equations render when present,
   *  and are skipped (engine tree-shaken) when omitted. */
  private _math: MathRenderer | undefined;

  private constructor(worker: Worker, mode: 'main' | 'worker', wasmUrlOverride?: string | URL) {
    this._worker = worker;
    this._mode = mode;
    this._bridge = new WorkerBridge<WorkerResponse | RenderWorkerResponse>(this._worker, {
      // Every response carries an id (no `ready` handshake — the worker `await`s
      // its own init promise before each request, docx/xlsx pattern).
      correlate: (msg) => msg.id,
      toError: (msg) => (msg.kind === 'error' ? msg.message : undefined),
    });
    // Default: the parser WASM emitted next to this bundle, resolved relative to
    // the document URL. `wasmUrl` overrides it (CDN / self-hosted copy); a
    // relative override is still resolved against `location.href`.
    const wasmUrl = new URL(wasmUrlOverride ?? wasmAssetUrl, location.href).href;
    this._bridge.post({ kind: 'init', wasmUrl } satisfies WorkerRequest);
  }

  /** Parse a PPTX from URL or ArrayBuffer. */
  static async load(
    source: string | ArrayBuffer,
    opts: LoadOptions = {},
  ): Promise<PptxPresentation> {
    const mode = opts.mode ?? 'main';
    if (mode === 'worker' && (typeof Worker === 'undefined' || typeof OffscreenCanvas === 'undefined')) {
      throw new Error("mode: 'worker' requires Worker and OffscreenCanvas support");
    }
    let buffer: ArrayBuffer;
    if (typeof source === 'string') {
      const res = await fetch(source);
      if (!res.ok) throw new Error(`Failed to fetch: ${res.status} ${res.statusText}`);
      buffer = await res.arrayBuffer();
    } else {
      buffer = source;
    }
    // Resolve the container on the main thread — before spinning up the worker.
    // A normal ZIP passes through unchanged; an Agile-encrypted CFB is decrypted
    // when `opts.password` is supplied ([MS-OFFCRYPTO]); a password-protected
    // file without a password, or a legacy-binary / unknown CFB, becomes a typed
    // OoxmlError (whose `instanceof` would not survive the worker boundary).
    buffer = toArrayBuffer(await resolveOoxmlContainer(buffer, opts.password));
    // The render worker is reachable only through this dynamic import, so
    // main-mode bundles never pull in its (renderer-bearing) chunk.
    const worker =
      mode === 'worker'
        ? (await import('./render-worker-host')).createRenderWorker()
        : new InlineWorker();
    const pres = new PptxPresentation(worker, mode, opts.wasmUrl);
    if (opts.math && mode === 'worker') {
      console.warn(
        "[ooxml] the math engine is unavailable in mode: 'worker'; equations will be skipped. Use mode: 'main' for documents with equations.",
      );
    }
    pres._math = mode === 'worker' ? undefined : opts.math;
    await pres._parse(
      buffer,
      opts.maxZipEntryBytes,
      mode === 'worker' ? !!opts.useGoogleFonts : false,
      opts.workerTimeoutMs,
    );
    if (mode === 'main' && opts.useGoogleFonts && pres._presentation) {
      pres._googleFontFaces = await preloadGoogleFonts(
        pptxFontPreloadNames(pres._presentation),
        PPTX_GOOGLE_FONTS,
      );
    }
    return pres;
  }

  private async _parse(
    buffer: ArrayBuffer,
    maxZipEntryBytes?: number,
    useGoogleFonts = false,
    timeoutMs?: number,
  ): Promise<void> {
    const res = await this._bridge.request(
      (id) =>
        this._mode === 'worker'
          ? ({ kind: 'parse', id, buffer, maxZipEntryBytes, useGoogleFonts } satisfies RenderWorkerRequest)
          : ({ kind: 'parse', id, buffer, maxZipEntryBytes } satisfies WorkerRequest),
      [buffer],
      { timeoutMs },
    );
    if (this._mode === 'worker') {
      this._meta = (res as Extract<RenderWorkerResponse, { kind: 'parsedMeta' }>).meta;
    } else {
      // The model arrives as transferred UTF-8 JSON bytes; decode + parse once
      // here (the only serialization on the parse-mode path).
      const { presentationJson } = res as Extract<WorkerResponse, { kind: 'parsed' }>;
      this._presentation = JSON.parse(
        new TextDecoder().decode(new Uint8Array(presentationJson)),
      ) as Presentation;
    }
  }

  /** Total number of slides in the loaded presentation. */
  get slideCount(): number { return this._presentation?.slides.length ?? this._meta?.slideCount ?? 0; }

  /** Slide width in EMU. */
  get slideWidth(): number { return this._presentation?.slideWidth ?? this._meta?.slideWidth ?? 0; }

  /** Slide height in EMU. */
  get slideHeight(): number { return this._presentation?.slideHeight ?? this._meta?.slideHeight ?? 0; }

  /** The render mode this engine was loaded with ('main' | 'worker'). A fact for
   *  integrators and the scroll viewer: an injected engine's mode decides whether
   *  slides render via renderSlide (main) or renderSlideToBitmap (worker) — no
   *  probing (design §11: no silent mis-pathing). */
  get mode(): 'main' | 'worker' {
    return this._mode;
  }

  /**
   * Atomically apply serializable top-level deltas to one slide-owned shape and
   * return the inverse batch needed for undo.
   *
   * Nested properties such as `textBody` are replaced as complete values. Shape
   * identity (`type` and `id`) is immutable.
   *
   * Main mode only: in worker mode the parsed presentation lives inside the
   * render worker and is not synchronously mutable from this instance.
   */
  applyShapeChanges(request: PptxApplyShapeChangesRequest): PptxShapeChangesResult {
    const { slideIndex, shapeId, changes } = request;
    if (this._mode !== 'main') {
      throw new Error('PptxPresentation.applyShapeChanges requires mode: "main"');
    }
    if (!this._presentation) throw new Error('Presentation not loaded');
    const slide = this._presentation.slides[slideIndex];
    if (!slide) {
      throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    }
    const elementIndex = slide.elements.findIndex(
      (element) => element.type === 'shape' && element.id === shapeId,
    );
    if (elementIndex < 0) {
      throw new Error(`Shape ${shapeId} not found on slide ${slideIndex}`);
    }

    const current = slide.elements[elementIndex] as ShapeElement;
    const draft = structuredClone(current);
    const { applied, inverse } = applyPptxShapeChanges(draft, changes);
    if (draft.type !== 'shape' || draft.id !== current.id) {
      throw new Error(
        'PptxPresentation.applyShapeChanges cannot change shape identity (type or id)',
      );
    }

    if (applied.length > 0) slide.elements[elementIndex] = draft;
    return {
      slideIndex,
      shapeId,
      shape: structuredClone(applied.length > 0 ? draft : current),
      applied: structuredClone(applied),
      inverse: structuredClone(inverse),
    };
  }

  /**
   * Speaker-notes text for a slide (`ppt/notesSlides/notesSlideN.xml`,
   * ECMA-376 §13.3.5 — Notes Slide). Returns the notes-body text as a single
   * string (paragraphs joined with `\n`), or `null` when the slide has no
   * notes part. The notes are parsed at {@link load} time, so this is a
   * synchronous lookup.
   *
   * `slideIndex` is 0-based. Unlike navigation methods it is *not* clamped:
   * an out-of-range or non-integer index returns `null` rather than the notes
   * of the nearest slide (so a tool iterating by index gets an honest "no
   * notes" instead of a duplicated neighbour).
   *
   * @example
   * const pres = await PptxPresentation.load(buffer);
   * for (let i = 0; i < pres.slideCount; i++) {
   *   const notes = pres.getNotes(i);
   *   if (notes) console.log(`Slide ${i + 1} notes:`, notes);
   * }
   */
  getNotes(slideIndex: number): string | null {
    if (this._meta) {
      // Worker mode: the model lives in the worker, so honour the same
      // non-clamped contract against the per-slide notes array.
      return Number.isInteger(slideIndex) ? (this._meta.notes[slideIndex] ?? null) : null;
    }
    return selectNotes(this._presentation?.slides ?? [], slideIndex);
  }

  /**
   * Whether the slide at `slideIndex` (0-based, absolute) is marked hidden
   * (`<p:sld show="0">`, ECMA-376 §19.3.1.38). Like {@link getNotes} the index
   * is NOT clamped — out-of-range / non-integer ⇒ `false`. This is a *fact*
   * about the model; deciding what to do with a hidden slide (skip / dim) is the
   * caller's policy (see {@link PptxViewer}'s `hiddenSlideMode` modes).
   */
  isHidden(slideIndex: number): boolean {
    if (this._meta) {
      return Number.isInteger(slideIndex) ? (this._meta.hidden[slideIndex] ?? false) : false;
    }
    return selectHidden(this._presentation?.slides ?? [], slideIndex);
  }

  /** The per-slide `partName` array (`sldIdLst` order) from either the parsed
   *  model (main) or the worker meta (worker). Backs the lazy part-index map. */
  private _partNames(): SlidePartNames {
    if (this._meta) return this._meta.partNames;
    return (this._presentation?.slides ?? []).map((s) => s.partName);
  }

  /** Lazily build (and cache) the `partName → index` map. Nulled by
   *  {@link destroy} so a reused reference never serves a stale deck's indices. */
  private _partIndex(): Map<string, number> {
    if (!this._slidePartIndex) this._slidePartIndex = buildSlidePartIndex(this._partNames());
    return this._slidePartIndex;
  }

  /**
   * Resolve a slide's OPC part name (e.g. `ppt/slides/slide3.xml`) to its
   * 0-based index in `sldIdLst` order, or `undefined` when no slide has that
   * part name. This is the map an internal hyperlink slide jump
   * (`<a:hlinkClick action="ppaction://hlinksldjump" r:id>`, ECMA-376
   * §21.1.2.3.5) resolves against: the click's rel Target names a slide part, and
   * this turns it into the index a viewer can navigate to. Works in both `main`
   * and `worker` mode (the part names ride along in the worker meta).
   */
  getSlideIndexByPartName(partName: string): number | undefined {
    return this._partIndex().get(partName);
  }

  /**
   * Resolve an internal hyperlink target string to a 0-based slide index, or
   * `undefined` when it names no reachable slide. Handles both
   * `<a:hlinkClick @action>` classes (§21.1.2.3.5):
   *
   *   - a **relative** show jump — `ppaction://hlinkshowjump?jump=firstslide |
   *     lastslide | nextslide | previousslide` — resolved arithmetically from
   *     `currentIndex` (clamped at the deck ends);
   *   - a **specific** slide-part jump — `ppaction://hlinksldjump`, whose
   *     resolved target is a slide-rel part name like `../slides/slide3.xml` —
   *     resolved through {@link getSlideIndexByPartName}.
   *
   * `ref` is the internal reference a `HyperlinkTarget` of kind `'internal'`
   * carries: the raw `ppaction://…` action string for a relative jump, or the
   * resolved slide-part target string for a specific jump. A viewer's
   * `onHyperlinkClick` default calls this with `ref` and the current slide, then
   * navigates to the returned index.
   *
   * @param ref          the internal action/target string.
   * @param currentIndex the 0-based slide the jump is relative to (default 0).
   */
  resolveInternalTarget(ref: string, currentIndex = 0): number | undefined {
    return resolveInternalSlideTarget(ref, this._partIndex(), currentIndex);
  }

  /** Render a slide onto the given canvas. */
  async renderSlide(
    canvas: HTMLCanvasElement | OffscreenCanvas,
    slideIndex: number,
    opts: RenderSlideOptions = {},
  ): Promise<void> {
    if (this._mode === 'worker') {
      throw new Error(
        "renderSlide(canvas) is unavailable in mode: 'worker'; use renderSlideToBitmap() and paint it via an ImageBitmapRenderingContext",
      );
    }
    if (!this._presentation) throw new Error('Presentation not loaded');
    const slide = this._presentation.slides[slideIndex];
    if (!slide) throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    const dpr = opts.dpr ?? defaultDpr();
    const width = opts.width ?? ((isHTMLCanvas(canvas) ? canvas.offsetWidth : 0) || 960);
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
        hlinkColor: this._presentation.hlinkColor ?? null,
        fetchMedia: (path) => this.getMedia(path),
        fetchImage: this._fetchImage,
        skipMediaControls: opts.skipMediaControls,
        dim: opts.dim,
        math: this._math,
      },
      opts.onTextRun,
    );
  }

  /**
   * Render a slide and return it as an ImageBitmap. Works in both modes; in
   * worker mode the entire render runs off the main thread. Paint with:
   * `canvas.getContext('bitmaprenderer').transferFromImageBitmap(bitmap)`.
   *
   * The returned ImageBitmap is owned by the caller: pass it to
   * `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`
   * when done, or its backing memory is held until GC.
   */
  async renderSlideToBitmap(
    slideIndex: number,
    opts: RenderSlideToBitmapOptions = {},
  ): Promise<ImageBitmap> {
    const width = opts.width ?? 960;
    const dpr = opts.dpr ?? defaultDpr();
    if (this._mode === 'worker') {
      if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
        throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
      }
      const res = await this._bridge.request(
        (id) => ({ kind: 'renderSlide', id, slideIndex, width, dpr, skipMediaControls: opts.skipMediaControls, dim: opts.dim }) satisfies RenderWorkerRequest,
      );
      const rendered = res as Extract<RenderWorkerResponse, { kind: 'slideRendered' }>;
      // IX6 — replay the worker's run geometry to the caller's collector so the
      // selection / find overlay is built on the same path as main mode.
      if (opts.onTextRun) for (const r of rendered.runs) opts.onTextRun(r);
      return rendered.bitmap;
    }
    const off = new OffscreenCanvas(1, 1);
    await this.renderSlide(off, slideIndex, {
      width,
      dpr,
      skipMediaControls: opts.skipMediaControls,
      dim: opts.dim,
      onTextRun: opts.onTextRun,
    });
    return off.transferToImageBitmap();
  }

  /**
   * IX6 — collect a slide's text-run geometry (`PptxTextRunInfo[]`) without
   * painting a visible canvas. Works in BOTH modes: worker mode renders the
   * slide off-thread and ships only the runs (no bitmap transfer); main mode
   * renders to a throwaway offscreen canvas. Used by the find controller to scan
   * every slide for matches. Run geometry is in CSS px (independent of dpr) and
   * dimming does not move glyphs, so only `width` is threaded — matching the
   * historical main-mode `_collectSlideRuns`.
   */
  async collectSlideRuns(slideIndex: number, width = 960): Promise<PptxTextRunInfo[]> {
    if (this._mode === 'worker') {
      if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
        throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
      }
      const res = await this._bridge.request(
        (id) => ({ kind: 'collectRuns', id, slideIndex, width }) satisfies RenderWorkerRequest,
      );
      return (res as Extract<RenderWorkerResponse, { kind: 'runsCollected' }>).runs;
    }
    const runs: PptxTextRunInfo[] = [];
    const off = new OffscreenCanvas(1, 1);
    await this.renderSlide(off, slideIndex, { width, onTextRun: (r) => runs.push(r) });
    return runs;
  }

  /**
   * Extract raw media bytes for a zip path referenced by {@link MediaElement}.
   * Results are cached by path for the lifetime of this instance.
   */
  async getMedia(mediaPath: string): Promise<Blob> {
    const hit = this._mediaCache.get(mediaPath);
    if (hit) return hit;
    // Worker mode has no main-thread model, so the mime lookup is skipped and
    // the Blob carries an empty type. That is fine: presentation-handle.ts
    // re-types blobs from MediaElement.mimeType when it builds media controls.
    const mimeType = this._findMimeTypeForPath(mediaPath);
    const p = (async () => {
      const res = await this._bridge.request(
        (id) => ({ kind: 'extractMedia', id, path: mediaPath }) satisfies WorkerRequest,
      );
      const bytes = (res as Extract<WorkerResponse, { kind: 'mediaExtracted' }>).bytes;
      return new Blob([bytes], { type: mimeType });
    })();
    this._mediaCache.set(mediaPath, p);
    return p;
  }

  private _findMimeTypeForPath(mediaPath: string): string {
    if (!this._presentation) return '';
    return findMimeTypeForPath(this._presentation, mediaPath);
  }

  /**
   * Extract raw bytes for an embedded image by zip path (e.g.
   * "ppt/media/image1.png"), wrapped in a Blob of the given MIME type. Mirrors
   * {@link getMedia}; results are cached by path for the lifetime of this
   * instance. The renderer routes its `fetchImage` option here so images are
   * decoded lazily rather than inlined as base64 at parse time.
   */
  async getImage(imagePath: string, mimeType: string): Promise<Blob> {
    const hit = this._imageCache.get(imagePath);
    if (hit) return hit;
    const p = (async () => {
      const res = await this._bridge.request(
        (id) => ({ kind: 'extractImage', id, path: imagePath }) satisfies WorkerRequest,
      );
      const bytes = (res as Extract<WorkerResponse, { kind: 'imageExtracted' }>).bytes;
      return new Blob([bytes], { type: mimeType });
    })();
    this._imageCache.set(imagePath, p);
    return p;
  }

  /**
   * Project the presentation to GitHub-flavoured markdown: title slides become
   * `#` headings, body shapes become nested bullets at each paragraph's `lvl`,
   * tables become pipe tables, charts become summarised bullets, and speaker
   * notes and comments are collated. Positioning, animations, images, and
   * drawing detail are discarded — the projection is meant for AI ingestion and
   * full-text search, not layout.
   *
   * Runs entirely in the worker off the archive opened at {@link load} (no
   * re-copy of the file, no re-parse of the model on the main thread), so it
   * works in BOTH `mode: 'main'` and `mode: 'worker'`.
   *
   * @example
   * const pres = await PptxPresentation.load(buffer);
   * const md = await pres.toMarkdown();
   */
  async toMarkdown(): Promise<string> {
    const res = await this._bridge.request(
      (id) => ({ kind: 'toMarkdown', id }) satisfies WorkerRequest,
    );
    return (res as Extract<WorkerResponse, { kind: 'markdownRendered' }>).markdown;
  }

  /**
   * Render a slide and attach canvas-native playback controls for any
   * embedded audio/video. Returns a {@link PresentationHandle} that owns the
   * RAF loop, media elements, and object URLs. Unlike {@link renderSlide}, this
   * method is stateful — always call `handle.destroy()` when leaving the slide.
   */
  async presentSlide(
    canvas: HTMLCanvasElement,
    slideIndex: number,
    opts: RenderSlideOptions = {},
  ): Promise<PresentationHandle> {
    if (this._mode === 'main' && !this._presentation) {
      throw new Error('Presentation not loaded');
    }
    if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= this.slideCount) {
      throw new Error(`Slide index ${slideIndex} out of range (count: ${this.slideCount})`);
    }
    const dpr = opts.dpr ?? defaultDpr();
    const width = opts.width ?? (canvas.offsetWidth || 960);

    const drawBase =
      this._mode === 'worker'
        ? async () => {
            // Whole slide rendered off-thread; the handle snapshots this paint
            // into its own base copy, so the bitmap can be closed right after.
            // IX6 — the run geometry rides back beside the bitmap, so a media
            // slide is as selectable/searchable in worker mode as in main mode.
            const bmp = await this.renderSlideToBitmap(slideIndex, { width, dpr, skipMediaControls: true, dim: opts.dim, onTextRun: opts.onTextRun });
            canvas.width = bmp.width;
            canvas.height = bmp.height;
            // Set only the CSS width and let height follow the intrinsic aspect
            // ratio — mirrors the main renderer (renderer.ts), which avoids an
            // explicit style.height that could fight the ratio.
            canvas.style.width = `${Math.round(bmp.width / dpr)}px`;
            if (!canvas.style.display) canvas.style.display = 'block';
            const ctx = canvas.getContext('2d');
            if (!ctx) throw new Error('2D context not available');
            ctx.drawImage(bmp, 0, 0);
            bmp.close();
          }
        : () =>
            this.renderSlide(canvas, slideIndex, {
              width,
              dpr,
              skipMediaControls: true,
              dim: opts.dim,
              onTextRun: opts.onTextRun,
            });

    const mediaElements =
      this._mode === 'worker'
        ? (this._meta?.mediaElements[slideIndex] ?? [])
        : (this._presentation as Presentation).slides[slideIndex].elements.filter(
            (el): el is MediaElement => el.type === 'media',
          );

    return createPresentationHandle(canvas, mediaElements, {
      width,
      dpr,
      slideWidthEmu: this.slideWidth,
      fetchMedia: (path) => this.getMedia(path),
      fetchImage: this._fetchImage,
      drawBase,
    });
  }

  /** Terminate the worker and release all resources. */
  destroy(): void {
    this._bridge.terminate();
    this._presentation = null;
    this._meta = null;
    this._slidePartIndex = null;
    this._mediaCache.clear();
    this._imageCache.clear();
    // Release the Google-Fonts substitutes this deck preloaded into the shared
    // FontFaceSet (main mode). Refcounted in core: a web font also used by another
    // open deck stays until that one is destroyed too. Without this, every opened
    // deck left its Google FontFace objects in `document.fonts` forever (SPA leak).
    if (this._googleFontFaces.length > 0) {
      unloadGoogleFonts(this._googleFontFaces);
      this._googleFontFaces = [];
    }
    // Release this deck's decoded raster bitmaps (GPU-backed), duotone-recoloured
    // rasters, and SVG object URLs promptly; all three caches are keyed by
    // `_fetchImage`.
    dropImageBitmapCache(this._fetchImage);
    dropDuotoneBitmapCache(this._fetchImage);
    dropSvgImageCache(this._fetchImage);
  }
}
