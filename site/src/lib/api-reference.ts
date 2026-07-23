// API reference data for the per-format pages. Hand-extracted from the real
// source (viewer.ts / presentation.ts / document.ts + the shared RenderOptions
// / LoadOptions). Keep in sync when the public types change.

export interface ApiOption {
  name: string;
  type: string;
  def?: string;
  desc: string;
}
export interface ApiMethod {
  sig: string;
  desc: string;
}
export interface ApiClass {
  name: string;
  ctor: string;
  note?: string;
  options?: ApiOption[];
  methods: ApiMethod[];
}

const ZIP = { name: 'maxZipEntryBytes', type: 'number', def: '512 MiB', desc: 'Per-entry ZIP decompression cap (zip-bomb guard). Lower it for untrusted input. Zero / negative values fall back to the default.' };
const GFONTS = { name: 'useGoogleFonts', type: 'boolean', def: 'false', desc: 'Load metric-compatible webfonts and non-Latin script fallbacks (Noto Arabic / CJK KR·SC·TC·JP / Cyrillic / Hebrew / Thai / Devanagari) from Google Fonts so layout matches Office and non-Latin text never falls back to tofu. Off by default for privacy.' };
const DPR = { name: 'dpr', type: 'number', def: 'devicePixelRatio', desc: 'Device pixel ratio for the backing store (crispness on HiDPI).' };
const WASM_URL = { name: 'wasmUrl', type: 'string | URL', def: 'bundled asset', desc: 'Override the URL the parser worker fetches the WebAssembly module from. By default each format resolves the `*_parser_bg.wasm` asset that ships next to its bundle (relative to the module URL); set this to serve it from a CDN or a self-hosted path instead (a relative value resolves against the document URL). Pointing it at a mismatched or missing file makes load() reject when the worker instantiates it.' };
const WORKER_TIMEOUT = { name: 'workerTimeoutMs', type: 'number', def: 'unlimited', desc: 'Reject the parse if the worker does not answer within this many ms — an opt-in safety net for a wedged / crashed worker that would otherwise leave load() pending forever. Unlimited by default (a large document with heavy media can legitimately take tens of seconds). A worker that throws or fails to load already rejects immediately regardless; this only covers the "silent, never-responds" case.' };
const MATH = { name: 'math', type: 'MathRenderer', def: 'undefined', desc: 'Opt-in OMML equation engine (MathJax + STIX Two Math, ~3 MB). Import it from the separate @silurus/ooxml/math entry — `import { math } from "@silurus/ooxml/math"` — and pass it to render equations. Omit it and equations are skipped, and the engine is left out of your build. When passed, the engine ships as a standalone asset fetched lazily the first time a document contains an equation.' };
const MODE = { name: 'mode', type: "'main' | 'worker'", def: "'main'", desc: "'main' parses in a worker and renders on the main thread (default). 'worker' parses AND renders entirely inside the worker; the main thread only paints the ImageBitmap returned by the render*ToBitmap method via a `bitmaprenderer` context. Requires Worker + OffscreenCanvas. The canvas-target render methods are unavailable in 'worker' mode, and equations require 'main'. Trade-off: each frame is transferred from the worker as an ImageBitmap, so a single render can be marginally slower than 'main' — the win is that the main thread never blocks." };
const VIEWER_MODE = { name: 'mode', type: "'main' | 'worker'", def: "'main'", desc: "'main' renders on the main thread (default). 'worker' renders the whole viewer off the main thread — every frame is produced in a Web Worker and painted via a `bitmaprenderer` context — so document rendering never blocks the UI. Scroll, sheet tabs, zoom and (xlsx) cell selection are unchanged. Requires Worker + OffscreenCanvas. The pptx/docx text-selection overlay and in-document find work in 'worker' mode too (per-run geometry crosses the worker boundary); equations still require 'main'. Trade-off: each frame crosses the worker boundary as an ImageBitmap, so an individual render can be marginally slower than 'main' — the win is a responsive main thread, not raw render speed." };
const ZOOM_MIN_MAX = { name: 'zoomMin / zoomMax', type: 'number', def: '0.1 / 4', desc: 'Zoom factor bounds for setScale / fitWidth / fitPage (10%–400%).' };
const ON_SCALE_CHANGE = { name: 'onScaleChange', type: '(scale: number) => void', desc: 'Called when the zoom factor changes (setScale / fitWidth / fitPage / zoomIn / zoomOut), with the clamped factor (1 = 100%).' };
const ON_HYPERLINK_CLICK = { name: 'onHyperlinkClick', type: '(target: HyperlinkTarget) => void', desc: "Called when a hyperlink is clicked. `target` is `{ kind: 'external', url }` or `{ kind: 'internal', ref, slideIndex? }`. When supplied, the callback fully owns the click (the default external-open / internal-navigation is not run). External URLs are scheme-sanitized (http / https / mailto / tel only); internal targets resolve to a docx bookmark / pptx slide jump / xlsx defined name or cell." };
const ENABLE_HYPERLINKS = { name: 'enableHyperlinks', type: 'boolean', def: 'true', desc: "Master switch for hyperlink interactivity. Set `false` to disable it entirely: no hit-testing, no pointer cursor over links, no default navigation, and `onHyperlinkClick` is never called. Links still render exactly as authored but are inert, like plain text." };

// Shared zoom methods (IX9) — same contract across all three viewers; the return
// type differs (docx/pptx re-render asynchronously → Promise<void>; xlsx is sync).
const zoomMethods = (asyncSet: boolean): ApiMethod[] => [
  { sig: 'getScale(): number', desc: 'The current zoom factor (1 = 100%).' },
  { sig: `setScale(scale: number): ${asyncSet ? 'Promise<void>' : 'void'}`, desc: 'Set the absolute zoom factor (1 = 100%), clamped to [zoomMin, zoomMax]; re-renders at the new size and fires onScaleChange when it changes. View-only.' },
  { sig: `fitWidth(): ${asyncSet ? 'Promise<void>' : 'void'}`, desc: "Fit the content WIDTH to the host container and re-render (routes through setScale). Defers when nothing is loaded or the container is unlaid-out." },
  { sig: `fitPage(): ${asyncSet ? 'Promise<void>' : 'void'}`, desc: 'Fit the WHOLE content (width and height) inside the container so it is visible without scrolling — takes the tighter of the two fits. Defers when unloaded / unlaid-out.' },
];

// Shared find methods (IX2) — identical shape across all three viewers; only the
// match location type differs (docx page / pptx slide / xlsx sheet+cell).
const findMethods = (loc: string): ApiMethod[] => [
  { sig: `findText(query: string, opts?: { caseSensitive?: boolean }): Promise<FindMatch<${loc}>[]>`, desc: 'Full-text search across the whole document; highlights every hit and returns them in document order. Each match carries `matchIndex`, the matched `text`, and its `location`. Case-insensitive by default.' },
  { sig: `findNext(): Promise<FindMatch<${loc}> | null>`, desc: 'Move to the next match (wrap-around), navigate to it if needed, and draw it in the active-match colour. Returns the now-active match, or null when there are none. Call findText first.' },
  { sig: `findPrev(): Promise<FindMatch<${loc}> | null>`, desc: 'Move to the previous match (wrap-around from first to last).' },
  { sig: 'clearFind(): void', desc: 'Clear all highlights and reset the find state.' },
];

export const apiReference: Record<'docx' | 'xlsx' | 'pptx', ApiClass[]> = {
  pptx: [
    {
      name: 'PptxViewer',
      ctor: 'new PptxViewer(canvas: HTMLCanvasElement, options?: PptxViewerOptions)',
      note: 'Opinionated single-canvas viewer. Hand it a <canvas>; it owns parsing, rendering and the current slide.',
      options: [
        { name: 'width', type: 'number', def: '960', desc: 'Canvas CSS width in px; height is derived from the slide aspect ratio.' },
        DPR,
        GFONTS,
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent text layer so users can select & copy slide text.' },
        { name: 'enableMediaPlayback', type: 'boolean', def: 'false', desc: 'Make embedded audio/video interactive (the viewer draws its own play chrome).' },
        { name: 'hiddenSlideMode', type: "'show' | 'skip' | 'dim'", def: "'show'", desc: 'How hidden slides (`<p:sld show="0">`, §19.3.1.38) are presented. `show` draws them like any other slide; `skip` makes sequential navigation (nextSlide/prevSlide and the initial load) jump over them while keeping absolute indices unchanged (an explicit goToSlide to a hidden slide is still honored); `dim` draws them under a translucent overlay (the PowerPoint thumbnail look).' },
        { name: 'hiddenSlideDim', type: 'Partial<DimOptions>', def: "{ color: '#ffffff', opacity: 0.6 }", desc: 'Overrides for the `dim` overlay, merged over the default white 60% wash. A partial so it stays in sync if DimOptions gains a field.' },
        ZIP,
        MATH,
        VIEWER_MODE,
        ZOOM_MIN_MAX,
        ON_SCALE_CHANGE,
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        { name: 'onSlideChange', type: '(index: number, total: number) => void', desc: 'Called after a slide finishes rendering.' },
        { name: 'onError', type: '(err: Error) => void', desc: 'Called on parse or render errors.' },
      ],
      methods: [
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load from a URL or ArrayBuffer and render the first slide.' },
        { sig: 'goToSlide(index: number): Promise<void>', desc: 'Render a specific slide (0-indexed, clamped).' },
        { sig: 'nextSlide(): Promise<void>', desc: 'Advance one slide.' },
        { sig: 'prevSlide(): Promise<void>', desc: 'Go back one slide.' },
        ...zoomMethods(true),
        ...findMethods('PptxMatchLocation'),
        { sig: 'get slideIndex(): number', desc: 'Current slide index.' },
        { sig: 'get slideCount(): number', desc: 'Total slides (0 until loaded).' },
        { sig: 'get hiddenSlideMode(): "show" | "skip" | "dim"', desc: 'The current hidden-slide mode.' },
        { sig: 'setHiddenSlideMode(mode: "show" | "skip" | "dim"): Promise<void>', desc: 'Switch the hidden-slide mode at runtime and re-render. Entering `skip` while on a hidden slide advances to the nearest visible slide.' },
        { sig: 'get visibleSlideCount(): number', desc: 'Number of non-hidden slides (the absolute slideCount is unchanged).' },
        { sig: 'getNotes(slideIndex: number): string | null', desc: 'Speaker-notes text for a slide (0-based); null when the slide has no notes part or the index is out of range.' },
        { sig: 'get canvasElement(): HTMLCanvasElement', desc: 'The underlying canvas.' },
        { sig: 'destroy(): void', desc: 'Tear down the worker and release resources.' },
      ],
    },
    {
      name: 'PptxPresentation',
      ctor: 'await PptxPresentation.load(source, options?)',
      note: 'Headless engine — parse once, render any slide into any canvas you supply (scroll views, thumbnail grids, master–detail).',
      options: [GFONTS, WASM_URL, ZIP, WORKER_TIMEOUT, MATH, MODE],
      methods: [
        { sig: 'static load(source, options?): Promise<PptxPresentation>', desc: 'Parse a deck from a URL or ArrayBuffer.' },
        { sig: 'get slideCount(): number', desc: 'Total slides.' },
        { sig: 'applyShapeChanges(slideIndex, shapeId, changes): PptxShapeChangesResult', desc: 'Atomically apply serializable add/replace/remove operations to a slide-owned ShapeElement in `mode: "main"`. Returns detached applied and inverse batches for request logging and undo/redo. Call renderSlide with the returned slideIndex to redraw only that slide.' },
        { sig: 'updateShape(slideIndex, shapeId, updater): PptxShapeUpdateResult', desc: 'Atomically update any property of a slide-owned ShapeElement in `mode: "main"`. The updater receives an isolated draft, including nested text bodies/runs for text, font, and colour edits. Identity (`type` / `id`) is immutable. Call renderSlide with the returned slideIndex to redraw only that slide.' },
        { sig: 'renderSlide(canvas, index, opts?: { width?, dpr?, onTextRun?, dim? }): Promise<void>', desc: 'Render one slide into the given canvas at the given width. `onTextRun` is called per rendered text segment so a caller can build a transparent selection overlay; `dim` (a DimOptions) paints a translucent wash over the finished slide (hidden-slide dimming). Equations render when a `math` engine was passed to `load`. Unavailable in `mode: "worker"` — use renderSlideToBitmap.' },
        { sig: 'renderSlideToBitmap(index, opts?: { width?, dpr?, dim? }): Promise<ImageBitmap>', desc: 'Render one slide and return it as an ImageBitmap (both modes; in worker mode the render runs off the main thread). `dim` paints a translucent overlay over the slide (hidden-slide dimming). Equations are skipped in `mode: "worker"` (they require `mode: "main"`). The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.' },
        { sig: 'presentSlide(canvas, index, opts?: { width?, dpr?, onTextRun? }): Promise<PresentationHandle>', desc: 'Render a slide and attach canvas-native audio/video playback, returning a handle with play() / pause() / destroy(). Works in both modes — in `mode: "worker"` the base slide is rendered off the main thread and the video overlay is composited on the main thread; `onTextRun` is unavailable there (it cannot cross the worker boundary).' },
        { sig: 'getNotes(slideIndex: number): string | null', desc: 'Speaker-notes text for a slide (0-based; ECMA-376 §13.3.5). Returns null when the slide has no notes part or the index is out of range.' },
        { sig: 'get slideWidth(): number', desc: 'Slide width in EMU (0 until loaded).' },
        { sig: 'get slideHeight(): number', desc: 'Slide height in EMU (0 until loaded).' },
        { sig: 'get mode(): "main" | "worker"', desc: 'The render mode this engine was loaded with. An injected engine’s mode decides whether slides render via renderSlide (main) or renderSlideToBitmap (worker).' },
        { sig: 'destroy(): void', desc: 'Release the worker.' },
      ],
    },
    {
      name: 'PptxScrollViewer',
      ctor: 'new PptxScrollViewer(container: HTMLElement, options?: PptxScrollViewerOptions)',
      note: 'Container-owning continuous-scroll viewer. Takes a <div> (not a canvas) and renders the whole deck as one vertically-scrolling, virtualized surface (only the visible window + overscan is mounted). Zoom is view-only.',
      options: [
        { name: 'width', type: 'number', def: 'container width', desc: 'Base fit width in CSS px. Default: the container width at first non-zero layout.' },
        { name: 'gap', type: 'number', def: '16', desc: 'Vertical gap (px) between consecutive slides.' },
        { name: 'paddingTop / paddingBottom', type: 'number', def: 'gap', desc: 'Desk padding (px) above the first slide / below the last. Pass 0 for a flush edge.' },
        { name: 'paddingLeft / paddingRight', type: 'number', def: 'gap', desc: 'Horizontal desk gutters (px); also shrink the container-derived fit width so a slide sits inside them at 100%. Pass 0 for a flush edge.' },
        { name: 'overscan', type: 'number', def: '1', desc: 'Slides kept mounted beyond the viewport on each side.' },
        { name: 'background', type: 'string', def: 'undefined', desc: 'CSS background for the scroll surface (the desk behind/between slides). Default transparent (the container shows through).' },
        { name: 'pageShadow', type: 'string | false', def: "'0 1px 3px rgba(0,0,0,0.2)'", desc: 'CSS box-shadow painted on every slide canvas. A spread-only ring (e.g. `0 0 0 1px #c8ccd0`) gives a crisp 1px border look. `false` disables it (flat slides).' },
        { name: 'enableZoom', type: 'boolean', def: 'true', desc: 'Enable Ctrl/⌘ + wheel (and trackpad pinch) zoom. View-only.' },
        { name: 'zoomMin / zoomMax', type: 'number', def: '0.1 / 4', desc: 'Absolute zoom scale bounds (10%–400%).' },
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent, selectable text layer per slide for native copy. `mode: "main"` only — in worker mode the overlay stays empty and the viewer warns once.' },
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        { name: 'presentation', type: 'PptxPresentation', def: 'undefined', desc: 'Inject an already-loaded engine to share one parse across panes. When set, load() is unsupported, the engine’s own mode wins, and destroy() does NOT destroy it (the caller owns its lifecycle).' },
        GFONTS,
        ZIP,
        MATH,
        DPR,
        MODE,
        { name: 'onVisibleSlideChange', type: '(topIndex: number, total: number) => void', desc: 'Fires when the top-most visible slide changes.' },
        { name: 'onError', type: '(err: Error) => void', desc: 'Called on load errors and async per-slide render failures (a failed slide is left blank rather than crashing the scroll loop).' },
      ],
      methods: [
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load a deck from a URL or ArrayBuffer and render the first window. Throws when an engine was injected via `presentation`.' },
        { sig: 'scrollToSlide(index: number, opts?: { behavior?: "auto" | "smooth" }): void', desc: 'Scroll so slide index’s top edge sits at the viewport top (index clamped).' },
        { sig: 'setScale(scale: number): void', desc: 'Set the absolute zoom scale at runtime (clamped to zoomMin/zoomMax). Flicker-free. View-only.' },
        { sig: 'relayout(): void', desc: 'Force a re-fit + re-mount of the visible window. Called automatically after load / resize / zoom; use it when the container resizes in a way a ResizeObserver cannot observe (e.g. a late web-font load). Idempotent.' },
        { sig: 'get slideCount(): number', desc: 'Total slides (0 until loaded).' },
        { sig: 'get topVisibleSlide(): number', desc: 'Index of the top-most visible slide.' },
        { sig: 'destroy(): void', desc: 'Tear down the DOM subtree. Destroys a self-loaded engine; an injected one is left intact.' },
      ],
    },
  ],

  docx: [
    {
      name: 'DocxViewer',
      ctor: 'new DocxViewer(canvas: HTMLCanvasElement, options?: DocxViewerOptions)',
      note: 'Single-canvas viewer that paginates the document and tracks the current page.',
      options: [
        { name: 'width', type: 'number', desc: 'Canvas CSS width in px; height is auto-computed from the page aspect ratio.' },
        DPR,
        GFONTS,
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent text layer for native selection & copy.' },
        { name: 'showTrackChanges', type: 'boolean', desc: 'Render tracked insertions/deletions with author colours.' },
        ZIP,
        MATH,
        VIEWER_MODE,
        ZOOM_MIN_MAX,
        ON_SCALE_CHANGE,
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        { name: 'onPageChange', type: '(index: number, total: number) => void', desc: 'Called after a page finishes rendering.' },
        { name: 'onError', type: '(err: Error) => void', desc: 'Called on parse or render errors.' },
      ],
      methods: [
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load from a URL or ArrayBuffer and render the first page.' },
        { sig: 'goToPage(index: number): Promise<void>', desc: 'Render a specific page (0-indexed, clamped).' },
        { sig: 'nextPage(): Promise<void>', desc: 'Advance one page.' },
        { sig: 'prevPage(): Promise<void>', desc: 'Go back one page.' },
        ...zoomMethods(true),
        ...findMethods('DocxMatchLocation'),
        { sig: 'get pageCount(): number', desc: 'Total pages (0 until loaded).' },
        { sig: 'get currentPage(): number', desc: 'Current page index.' },
        { sig: 'get canvasElement(): HTMLCanvasElement', desc: 'The underlying canvas.' },
        { sig: 'destroy(): void', desc: 'Tear down the worker and release resources.' },
      ],
    },
    {
      name: 'DocxDocument',
      ctor: 'await DocxDocument.load(source, options?)',
      note: 'Headless engine — render any page into any canvas you supply.',
      options: [GFONTS, WASM_URL, ZIP, WORKER_TIMEOUT, MATH, MODE],
      methods: [
        { sig: 'static load(source, options?): Promise<DocxDocument>', desc: 'Parse a document from a URL or ArrayBuffer.' },
        { sig: 'get pageCount(): number', desc: 'Total pages.' },
        { sig: 'pageSize(pageIndex: number): { widthPt, heightPt }', desc: 'Page size in pt for a page (ECMA-376 §17.6.13 / §17.6.11 — per section, so a mixed portrait/landscape document returns different sizes per page). Available in both modes; index is clamped. `{ 0, 0 }` means "not loaded". Returns a fresh object per call.' },
        { sig: 'get mode(): "main" | "worker"', desc: 'The render mode this engine was loaded with. An injected engine’s mode decides whether pages render via renderPage (main) or renderPageToBitmap (worker).' },
        { sig: 'renderPage(canvas, index, opts?: { width?, dpr?, showTrackChanges? }): Promise<void>', desc: 'Render one page into the given canvas. Unavailable in `mode: "worker"` — use renderPageToBitmap.' },
        { sig: 'renderPageToBitmap(index, opts?: { width?, dpr?, showTrackChanges? }): Promise<ImageBitmap>', desc: 'Render one page and return it as an ImageBitmap (both modes; in worker mode the render runs off the main thread). Equations are skipped in `mode: "worker"` (they require `mode: "main"`). The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.' },
      ],
    },
    {
      name: 'DocxScrollViewer',
      ctor: 'new DocxScrollViewer(container: HTMLElement, options?: DocxScrollViewerOptions)',
      note: 'Container-owning continuous-scroll viewer. Takes a <div> (not a canvas) and renders the whole document as one vertically-scrolling, virtualized surface (only the visible window + overscan is mounted). Zoom is view-only.',
      options: [
        { name: 'width', type: 'number', def: 'container width', desc: 'Base fit width in CSS px. Default: the container width at first non-zero layout.' },
        { name: 'gap', type: 'number', def: '16', desc: 'Vertical gap (px) between consecutive pages.' },
        { name: 'paddingTop / paddingBottom', type: 'number', def: 'gap', desc: 'Desk padding (px) above the first page / below the last. Pass 0 for a flush edge.' },
        { name: 'paddingLeft / paddingRight', type: 'number', def: 'gap', desc: 'Horizontal desk gutters (px); also shrink the container-derived fit width so a page sits inside them at 100%. Pass 0 for a flush edge.' },
        { name: 'overscan', type: 'number', def: '1', desc: 'Pages kept mounted beyond the viewport on each side.' },
        { name: 'background', type: 'string', def: 'undefined', desc: 'CSS background for the scroll surface (the desk behind/between pages). Default transparent (the container shows through).' },
        { name: 'pageShadow', type: 'string | false', def: "'0 1px 3px rgba(0,0,0,0.2)'", desc: 'CSS box-shadow painted on every page canvas. A spread-only ring (e.g. `0 0 0 1px #c8ccd0`) gives a crisp 1px border look. `false` disables it (flat pages).' },
        { name: 'enableZoom', type: 'boolean', def: 'true', desc: 'Enable Ctrl/⌘ + wheel (and trackpad pinch) zoom. View-only.' },
        { name: 'zoomMin / zoomMax', type: 'number', def: '0.1 / 4', desc: 'Absolute zoom scale bounds (10%–400%).' },
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent, selectable text layer per page for native copy. `mode: "main"` only — in worker mode the overlay stays empty and the viewer warns once.' },
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        { name: 'showTrackChanges', type: 'boolean', desc: 'Render tracked insertions/deletions with author colours (forwarded to each page render).' },
        { name: 'document', type: 'DocxDocument', def: 'undefined', desc: 'Inject an already-loaded engine to share one parse across panes. When set, load() is unsupported, the engine’s own mode wins, and destroy() does NOT destroy it (the caller owns its lifecycle).' },
        GFONTS,
        ZIP,
        MATH,
        DPR,
        MODE,
        { name: 'onVisiblePageChange', type: '(topIndex: number, total: number) => void', desc: 'Fires when the top-most visible page changes.' },
        { name: 'onError', type: '(err: Error) => void', desc: 'Called on load errors and async per-page render failures (a failed page is left blank rather than crashing the scroll loop).' },
      ],
      methods: [
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load a document from a URL or ArrayBuffer and render the first window. Throws when an engine was injected via `document`.' },
        { sig: 'scrollToPage(index: number, opts?: { behavior?: "auto" | "smooth" }): void', desc: 'Scroll so page index’s top edge sits at the viewport top (index clamped).' },
        { sig: 'setScale(scale: number): void', desc: 'Set the absolute zoom scale at runtime (clamped to zoomMin/zoomMax). Flicker-free. View-only.' },
        { sig: 'relayout(): void', desc: 'Force a re-fit + re-mount of the visible window. Called automatically after load / resize / zoom; use it when the container resizes in a way a ResizeObserver cannot observe (e.g. a late web-font load). Idempotent.' },
        { sig: 'get pageCount(): number', desc: 'Total pages (0 until loaded).' },
        { sig: 'get topVisiblePage(): number', desc: 'Index of the top-most visible page.' },
        { sig: 'destroy(): void', desc: 'Tear down the DOM subtree. Destroys a self-loaded engine; an injected one is left intact.' },
      ],
    },
  ],

  xlsx: [
    {
      name: 'XlsxViewer',
      ctor: 'new XlsxViewer(container: HTMLElement, options?: XlsxViewerOptions)',
      note: 'Full workbook viewer. Takes a container <div> (not a canvas) — it manages its own canvas, sheet-tab bar and zoom slider. Drag-to-resize columns/rows and zoom are view-only: they change the on-screen view only and never modify the loaded file.',
      options: [
        { name: 'cellScale', type: 'number', def: '1', desc: 'Scale factor for cell/header dimensions (0.5 = half size).' },
        { name: 'showZoomSlider', type: 'boolean', def: 'true', desc: 'Show the Excel-style zoom slider at the end of the tab bar. Zooming (slider, Ctrl/⌘+wheel, trackpad pinch) is view-only.' },
        { name: 'zoomMin / zoomMax', type: 'number', def: '0.1 / 4', desc: 'Zoom slider bounds as scale factors (10%–400%).' },
        { name: 'resizable', type: 'boolean', def: 'true', desc: 'Allow resizing columns/rows by dragging header borders. View-only — it changes the on-screen view only and never modifies the loaded file. Set false to disable.' },
        { name: 'selectionColor', type: 'string', def: "'#1a73e8'", desc: 'Accent color for the cell-selection rectangle (any CSS color). The fill is the same color at 8% opacity.' },
        { name: 'hiddenSheetMode', type: "'show' | 'skip' | 'dim'", def: "'show'", desc: 'How hidden / very-hidden sheets (`<sheet state>`, §18.2.19) appear in the tab bar. `show` renders a tab like any other; `skip` hides the tab (`display:none`) and makes sequential navigation jump over it; `dim` renders the tab at reduced opacity. Mirrors pptx `hiddenSlideMode`.' },
        GFONTS,
        ZIP,
        MATH,
        VIEWER_MODE,
        ON_SCALE_CHANGE,
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        { name: 'onReady', type: '(sheetNames: string[]) => void', desc: 'Called once the workbook is parsed.' },
        { name: 'onSheetChange', type: '(index: number, total: number) => void', desc: 'Called when the active sheet changes; `total` is the sheet count. Read the name via `sheetNames[index]`.' },
        { name: 'onSelectionChange', type: '(sel: CellRange | null) => void', desc: 'Called when the selected range changes; null clears it.' },
        { name: 'onError', type: '(err: Error) => void', desc: 'Called on parse or render errors.' },
      ],
      methods: [
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load a workbook from a URL or ArrayBuffer and render the first sheet.' },
        { sig: 'goToSheet(index: number): Promise<void>', desc: 'Show a specific sheet (0-indexed, clamped).' },
        { sig: 'nextSheet(): Promise<void>', desc: 'Advance one sheet.' },
        { sig: 'prevSheet(): Promise<void>', desc: 'Go back one sheet.' },
        { sig: 'get sheetIndex(): number', desc: 'Current sheet index.' },
        { sig: 'get sheetCount(): number', desc: 'Total sheets (0 until loaded).' },
        { sig: 'get sheetNames(): string[]', desc: 'Names of all sheets.' },
        { sig: 'get selection(): CellRange | null', desc: 'The current selected range.' },
        { sig: 'getScale(): number', desc: 'The current zoom factor (1 = 100%).' },
        { sig: 'setScale(scale: number): void', desc: 'Set the zoom factor (1 = 100%), clamped to [zoomMin, zoomMax] and snapped to whole percent; re-renders and fires onScaleChange when it changes. View-only.' },
        { sig: 'fitWidth(): void', desc: 'Fit the used data range WIDTH (row header + used columns) to the canvas area (routes through setScale). Defers when unloaded / unlaid-out.' },
        { sig: 'fitPage(): void', desc: 'Fit the used data range WIDTH and HEIGHT inside the canvas area so the whole used range is visible without scrolling — takes the tighter of the two fits. Defers when unloaded / unlaid-out.' },
        ...findMethods('XlsxMatchLocation'),
        { sig: 'setSelectionColor(color: string): void', desc: 'Change the selection accent color at runtime (any CSS color).' },
        { sig: 'get hiddenSheetMode(): "show" | "skip" | "dim"', desc: 'The current hidden-sheet mode.' },
        { sig: 'setHiddenSheetMode(mode: "show" | "skip" | "dim"): Promise<void>', desc: 'Switch the hidden-sheet mode at runtime: restyle the tabs and re-render. Entering `skip` while on a hidden sheet advances to the nearest visible sheet.' },
        { sig: 'getCellAt(clientX: number, clientY: number): CellAddress | null', desc: 'Hit-test a viewport coordinate to a cell address.' },
        { sig: 'get canvasElement(): HTMLCanvasElement', desc: 'The underlying canvas the grid is drawn on.' },
        { sig: 'destroy(): void', desc: 'Tear down the worker and release resources.' },
      ],
    },
    {
      name: 'XlsxWorkbook',
      ctor: 'await XlsxWorkbook.load(source, options?)',
      note: 'Headless engine — parse once, render any sheet viewport into any canvas you supply.',
      options: [GFONTS, WASM_URL, ZIP, WORKER_TIMEOUT, MATH, MODE],
      methods: [
        { sig: 'static load(source, options?): Promise<XlsxWorkbook>', desc: 'Parse a workbook from a URL or ArrayBuffer.' },
        { sig: 'get sheetNames(): string[]', desc: 'Names of all sheets.' },
        { sig: 'get sheetCount(): number', desc: 'Total sheets.' },
        { sig: 'renderViewport(canvas, sheetIndex, viewport, opts?: { width?, height?, dpr?, cellScale? }): Promise<void>', desc: 'Render a row/col window of a sheet into the given canvas. Equations in shapes render when a `math` engine was passed to `load`. Unavailable in `mode: "worker"` — use renderViewportToBitmap.' },
        { sig: 'renderViewportToBitmap(sheetIndex, viewport, opts: { width, height, dpr?, cellScale? }): Promise<ImageBitmap>', desc: 'Render a sheet viewport and return it as an ImageBitmap (both modes; in worker mode the render runs off the main thread). `width` and `height` are required — a worker has no DOM element to measure. Equations in shapes are skipped in `mode: "worker"` (they require `mode: "main"`). The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.' },
        { sig: 'resolveValidationList(sheetIndex, formula1): Promise<ResolvedList>', desc: 'Resolve a list-type data-validation `formula1` (ECMA-376 §18.3.1.32) into the allowed values to display — inline quoted list, a range reference (each cell’s display string), or `{ kind: \'formula\' }` for named ranges. Read-only.' },
        { sig: 'destroy(): void', desc: 'Release the worker.' },
      ],
    },
  ],
};
