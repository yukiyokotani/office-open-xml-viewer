// API reference data for the per-format pages. Hand-extracted from the real
// source (viewer.ts / presentation.ts / document.ts + the shared RenderOptions
// / LoadOptions). Keep in sync when the public types change.

export interface ApiOption {
  name: string;
  type: string;
  def?: string;
  desc: string;
  /** Exact, contract-critical substring rendered with semantic emphasis. */
  emphasis?: string;
  /** Optional route to the longer contract documentation for this option. */
  detailsHref?: string;
  detailsLabel?: string;
}
export interface ApiMethod {
  sig: string;
  desc: string;
  /** Exact, contract-critical substring rendered with semantic emphasis. */
  emphasis?: string;
}
export interface ApiClass {
  name: string;
  ctor: string;
  note?: string;
  options?: ApiOption[];
  methods: ApiMethod[];
}

export interface OptionalRendererReference {
  name: string;
  entry: string;
  exportName: string;
  contract: string;
  desc: string;
}

export const formatRenderModeGuidance: Readonly<Record<'docx' | 'xlsx' | 'pptx', string>> = {
  docx: 'DOCX keeps page navigation, virtualized scrolling, selection, find, hyperlinks, equations, ChartEx, 3-D charts, Region Maps and the built-in TIFF codec in both modes. Worker mode moves pagination and page paint away from the UI thread. Documents that require DOM OpenType vertical-glyph selection automatically use main mode; read the loaded document\'s mode to observe that fallback.',
  xlsx: 'XLSX keeps sheet tabs, frozen panes, scrolling, selection, find, hyperlinks, equations, ChartEx, 3-D charts, Region Maps and the built-in TIFF codec in both modes. Worker mode paints each requested sheet viewport away from the UI thread.',
  pptx: 'PPTX keeps slide navigation, virtualized scrolling, selection, find, hyperlinks, media playback, equations, ChartEx, 3-D charts, Region Maps and the built-in TIFF codec in both modes. Worker mode moves slide paint away from the UI thread; media controls and overlays remain interactive in the Viewer.',
};

export const optionalRenderers: readonly OptionalRendererReference[] = [
  {
    name: 'Equation renderer',
    entry: '@silurus/ooxml/math',
    exportName: 'math',
    contract: 'MathRenderer',
    desc: 'Renders OMML equations with the separately loaded MathJax and STIX Two Math asset. The asset is fetched lazily only when an equation is present.',
  },
  {
    name: 'Microsoft ChartEx renderer',
    entry: '@silurus/ooxml/chart-ex',
    exportName: 'chartEx',
    contract: 'ChartExRenderer',
    desc: 'Renders the newer cx:* waterfall, histogram, Pareto, funnel, box-and-whisker, sunburst and treemap families. Classic c:* 2-D charts remain in the format entries.',
  },
  {
    name: '3-D chart renderer',
    entry: '@silurus/ooxml/three-d',
    exportName: 'threeD',
    contract: 'ChartThreeDRenderer',
    desc: 'Renders the authored OOXML view with one model-space camera and bounded mesh pipeline. Supports the documented cartesian and pie 3-D families and authored bar/column shape meshes in main and worker modes.',
  },
  {
    name: 'Offline Region Map renderer',
    entry: '@silurus/ooxml/region-map',
    exportName: 'regionMap',
    contract: 'ChartRegionMapRenderer',
    desc: 'Renders supported country-level ChartEx maps without network access using a pinned public-domain Natural Earth asset. Cached provider identities and unsupported sub-country/view-specific layouts fail closed.',
  },
  {
    name: 'TIFF image codec',
    entry: '@silurus/ooxml/tiff',
    exportName: 'tiff',
    contract: 'TiffRenderer',
    desc: 'Decodes bounded stripped TIFF 6.0 images in DOCX, XLSX and PPTX. Supported classes are uncompressed bilevel, 8-bit grayscale, RGB, RGBA and process-CMYK, plus 1-bit CCITT Group 4. Unsupported or malformed classes make standalone calls and DOCX/PPTX rendering report TiffDecodeError. XLSX rendering, including XlsxViewer, contains the failure at that picture and shows an unavailable-image placeholder.',
  },
];

const RESOURCE_LIMITS = { name: 'resourceLimits', type: 'OoxmlResourceLimits', def: '128 MiB per entry / 256 MiB distinct total / 4,096 entries', desc: 'Shared DOCX/XLSX/PPTX package budgets. maxArchiveEntryBytes caps each package part; maxTotalInflatedBytes counts the largest amount read from every distinct part without charging repeat reads twice; maxArchiveEntries bounds central-directory entries before ZIP index allocation. Supply positive safe integers, or null to disable one configurable budget (internal hard ceilings remain). Violations reject with OoxmlResourceLimitError. These deterministic counters reduce OOM risk but do not measure or guarantee peak memory.', emphasis: 'Violations reject with OoxmlResourceLimitError.', detailsHref: '/errors#ooxml-resource-limit-error', detailsLabel: 'Error fields' };
const IMAGE_RESOURCES = { name: 'imageResources', type: 'ImageResourceOptions', def: "{ decodedByteBudget: 128 MiB, strategy: 'adaptive', resolution: 'native-if-fit' }", desc: "Decoded-raster policy shared by DOCX, XLSX and PPTX paints. Ordinary browser rasters receive a geometry-weighted share of decodedByteBudget before source extraction, allowing each source to flow directly into decode. A source keeps native resolution when it fits its share and otherwise uses up to a 2x canvas/DPR grid when that share has headroom. If the complete set of display grids exceeds the budget, adaptive mode reduces them by one uniform quality ratio. Set resolution: 'display' to minimize retained pixels. Natural-size consumers, pixel effects that require the authored grid, and non-resizable formats retain their guarded source-specific paths. Set strategy: 'strict' to preserve requested targets and receive OoxmlDecodedImageLimitError on an aggregate crossing. The budget accepts 4 bytes through 512 MiB; encoded-source, per-axis and per-surface hard safety ceilings remain non-disableable.", emphasis: 'A source keeps native resolution when it fits its share and otherwise uses up to a 2x canvas/DPR grid when that share has headroom.', detailsHref: '/errors#decoded-image-limit-error', detailsLabel: 'Safety boundaries' };
const RESOURCE_METRICS = { name: 'onResourceMetrics', type: '(metrics: OoxmlResourceMetrics) => void', desc: 'Receives the content-free initial-load report used by the debug card, without enabling console output. It reports the configured public policy, timing checkpoints, format/mode, success or typed failure discriminants, source bytes, and observed archive counters when available. It does not wait for a Viewer\'s first paint. On success, call getResourceMetrics() on the engine or Viewer for a fresh snapshot after lazy package work. Callback exceptions never change load results.', emphasis: 'Receives the content-free initial-load report used by the debug card, without enabling console output.' };
const RESOURCE_METRICS_METHOD = { sig: 'getResourceMetrics(): Promise<OoxmlResourceMetrics>', desc: 'Return a fresh, content-free package-usage snapshot, including lazy archive work observed since load. Collection is always active; debug controls only console output.', emphasis: 'Collection is always active; debug controls only console output.' };
const DEBUG = { name: 'debug', type: 'boolean', def: 'false', desc: 'Print one content-free, Ratatui-inspired resource report when the measured load or Node session finishes or fails. Browser DevTools use typography-only %c styling to keep Unicode borders and gauges aligned without changing foreground or background colours; Node and Worker consoles receive one plain argument. Use onResourceMetrics instead for production collection.', emphasis: 'Use onResourceMetrics instead for production collection.' };
const ZIP = { name: 'maxZipEntryBytes', type: 'number', def: 'resource policy default', desc: 'Deprecated compatibility alias for resourceLimits.maxArchiveEntryBytes. It is scheduled for removal in a future breaking release; new code should use resourceLimits. Existing positive values retain their per-entry meaning; zero / negative values fall back to the standard default.', emphasis: 'It is scheduled for removal in a future breaking release; new code should use resourceLimits.' };
const GFONTS = { name: 'useGoogleFonts', type: 'boolean', def: 'false', desc: 'Load metric-compatible webfonts and non-Latin script fallbacks (Noto Arabic / CJK KR·SC·TC·JP / Cyrillic / Hebrew / Thai / Devanagari) from Google Fonts so layout matches Office and non-Latin text never falls back to tofu. Off by default for privacy.', emphasis: 'Off by default for privacy.' };
const PASSWORD = { name: 'password', type: 'string', def: 'undefined', desc: 'Password for an Agile-encrypted OOXML file. Available on self-loading Viewer constructors and headless load(); borrowed fromDocument(), fromPresentation(), and fromWorkbook() factories omit load-only options because their engine is already loaded.', emphasis: 'Available on self-loading Viewer constructors and headless load()' };
const DPR = { name: 'dpr', type: 'number', def: 'devicePixelRatio', desc: 'Device pixel ratio for the backing store (crispness on HiDPI).' };
const WASM_URL = { name: 'wasmUrl', type: 'string | URL', def: 'bundled asset', desc: 'Override the URL the parser worker fetches the WebAssembly module from. By default each format resolves the `*_parser_bg.wasm` asset that ships next to its bundle (relative to the module URL); set this to serve it from a CDN or a self-hosted path instead (a relative value resolves against the document URL). Pointing it at a mismatched or missing file makes load() reject when the worker instantiates it.', emphasis: 'Override the URL the parser worker fetches the WebAssembly module from.' };
const WORKER_TIMEOUT = { name: 'workerTimeoutMs', type: 'number', def: 'unlimited', desc: 'Opt-in worker liveness limit. Ordinary worker requests use it as their response deadline. Worker-mode progressive loads restart this silence interval whenever the worker reports progress. This allows active long-running work to continue. Silence before first paint rejects load(); silence afterward keeps layoutComplete false and rejects waitUntilLayoutComplete(), while configured completion/error callbacks receive the failure. Worker exceptions still reject immediately. Unlimited by default.', emphasis: 'Worker-mode progressive loads restart this silence interval whenever the worker reports progress.' };
const MATH = { name: 'math', type: 'MathRenderer', def: 'undefined', desc: 'Opt-in OMML equation engine (MathJax + STIX Two Math, ~3 MB). Import it from the separate @silurus/ooxml/math entry — `import { math } from "@silurus/ooxml/math"` — and pass it to render equations in either mode. Omit it and equations are skipped; the MathJax asset is not fetched. When passed, that standalone asset is fetched lazily the first time a document contains an equation.', emphasis: 'Opt-in OMML equation engine (MathJax + STIX Two Math, ~3 MB).' };
const THREE_D = { name: 'threeD', type: 'ChartThreeDRenderer', def: 'undefined', desc: 'Opt-in model-space 3-D chart renderer. Import `threeD` from the separate `@silurus/ooxml/three-d` entry and inject it once. Omit it to use the canonical 2-D fallback and avoid loading or evaluating the mesh/camera implementation in main mode. The self-contained worker asset retains the worker-side implementation. It renders the view angle authored in OOXML in main and worker modes.', emphasis: 'Opt-in model-space 3-D chart renderer.' };
const REGION_MAP = { name: 'regionMap', type: 'ChartRegionMapRenderer', def: 'undefined', desc: 'Opt-in offline ChartEx Region Map renderer using a pinned, public-domain Natural Earth country asset. Import `regionMap` from `@silurus/ooxml/region-map` and inject it once. Unsupported cached or sub-country views fail closed. The built-in renderer works in main and worker modes.', emphasis: 'Opt-in offline ChartEx Region Map renderer' };
const CHART_EX = { name: 'chartEx', type: 'ChartExRenderer', def: 'undefined', desc: 'Opt-in renderer for Microsoft ChartEx (`cx:*`) chart families. Import `chartEx` from `@silurus/ooxml/chart-ex` and inject it once. Classic 2-D charts stay in the default format entries; ChartEx is opt-in. The built-in renderer works in main and worker modes.', emphasis: 'Classic 2-D charts stay in the default format entries; ChartEx is opt-in.' };
const TIFF = { name: 'tiff', type: 'TiffRenderer', def: 'undefined', desc: 'Opt-in TIFF image codec shared by DOCX, XLSX and PPTX. Import `tiff` from `@silurus/ooxml/tiff` and inject it once. The bounded codec accepts stripped TIFF 6.0 bilevel, grayscale, RGB, RGBA and process-CMYK images, plus CCITT Group 4 bilevel images. Omit it to keep the implementation out of ordinary format bundles; recognized TIFF images then use an unavailable-image placeholder while the rest of the document keeps rendering. Unsupported or malformed input makes standalone codec calls and DOCX/PPTX rendering report TiffDecodeError; XLSX rendering, including XlsxViewer, contains it at that picture and shows the placeholder. The built-in codec works in main and worker modes.', emphasis: 'Opt-in TIFF image codec shared by DOCX, XLSX and PPTX.' };
const MODE = { name: 'mode', type: "'main' | 'worker'", def: "'main'", desc: "Use 'main' for the smallest worker download, the lowest single-frame overhead or custom renderer objects; parsing still runs in a Worker, while Canvas rendering runs on the main thread. Use 'worker' when document layout and paint would compete with application UI responsiveness. It requires Worker and OffscreenCanvas, downloads a larger render worker and transfers an ImageBitmap per frame. Built-in math, ChartEx, 3-D, Region Map and TIFF renderers use the same options in both modes. In worker mode, use the bitmap render methods instead of methods that accept a Canvas.", emphasis: "Use 'worker' when document layout and paint would compete with application UI responsiveness." };
const VIEWER_MODE = { name: 'mode', type: "'main' | 'worker'", def: "'main'", desc: "Use 'main' for ordinary previews, the smallest worker download or custom renderer objects. Use 'worker' when rendering larger or more complex documents would compete with scrolling, navigation or other application UI. Worker mode requires Worker and OffscreenCanvas, downloads a larger render worker and transfers an ImageBitmap per frame. Viewer navigation, zoom, virtualized scrolling, selection, find, hyperlinks and the built-in math, ChartEx, 3-D, Region Map and TIFF renderers remain available in both modes.", emphasis: "Use 'worker' when rendering larger or more complex documents would compete with scrolling, navigation or other application UI." };
const ZOOM_MIN_MAX = { name: 'zoomMin / zoomMax', type: 'number', def: '0.1 / 4', desc: 'Zoom factor bounds for setScale / fitWidth / fitPage (10%–400%).' };
const ON_SCALE_CHANGE = { name: 'onScaleChange', type: '(scale: number) => void', desc: 'Called when the zoom factor changes (setScale / fitWidth / fitPage / zoomIn / zoomOut), with the clamped factor (1 = 100%).' };
const ON_HYPERLINK_CLICK = { name: 'onHyperlinkClick', type: '(target: HyperlinkTarget) => void', desc: "Called when a hyperlink is clicked. `target` is `{ kind: 'external', url }` or `{ kind: 'internal', ref, slideIndex? }`. When supplied, the callback fully owns the click (the default external-open / internal-navigation is not run). External URLs are scheme-sanitized (http / https / mailto / tel only); internal targets resolve to a docx bookmark, pptx slide jump, or xlsx defined name / cell reference. XLSX switches sheets before scrolling the destination into view and uses the first cell of a range.", emphasis: 'When supplied, the callback fully owns the click (the default external-open / internal-navigation is not run).' };
const ENABLE_HYPERLINKS = { name: 'enableHyperlinks', type: 'boolean', def: 'true', desc: "Master switch for hyperlink interactivity. Set `false` to disable it entirely: no hit-testing, no pointer cursor over links, no default navigation, and `onHyperlinkClick` is never called. Links still render exactly as authored but are inert, like plain text.", emphasis: 'Set `false` to disable it entirely' };
const CONTEXT_MENU = (contextType: string): ApiOption => ({
  name: 'onContextMenu',
  type: `(event: ViewerContextMenuEvent<${contextType}>) => void`,
  desc: 'Called synchronously for the native contextmenu event. Use originalEvent.preventDefault() before returning to replace the browser menu; getContext() starts one memoized target lookup on first call. Omit the callback to keep native browser behavior unchanged.',
  emphasis: 'Use originalEvent.preventDefault() before returning to replace the browser menu',
});
const VIEWER_ON_ERROR = {
  name: 'onError',
  type: '(err: Error) => void',
  desc: 'Receives Viewer-managed failures that have no directly awaitable result, such as virtualized rendering or embedded-media playback. load(), navigation, and other awaitable operations reject their own Promise whether or not this callback is supplied; the same failure is never delivered twice. Background failures are logged with console.error when the callback is omitted. Narrow stable cases with OoxmlError, OoxmlResourceLimitError, OoxmlDecodedImageLimitError or TiffDecodeError; other failures remain Error values and message text is not a stable discriminator.',
  emphasis: 'the same failure is never delivered twice.',
  detailsHref: '/errors#delivery',
  detailsLabel: 'Error reference',
};
const DOCX_PROGRESSIVE_LAYOUT: ApiOption = {
  name: 'progressiveLayout',
  type: 'boolean',
  def: 'false',
  desc: 'Resolve load() when the opening pages are paintable and continue the same canonical pagination session in the background. While layoutComplete is false, pageCount is the pages available so far rather than the final total. Works in main and worker modes; worker mode also keeps the remaining pagination off the UI thread.',
  emphasis: 'While layoutComplete is false, pageCount is the pages available so far rather than the final total.',
  detailsHref: '/docx#progressive-layout',
  detailsLabel: 'Progressive layout guide',
};
const DOCX_SHOW_TRACKED_CHANGES: ApiOption = {
  name: 'showTrackedChanges',
  type: 'boolean',
  def: 'false',
  desc: 'Select the tracked-change markup layout from the initial load. Insertions use author-coloured underlines, deletions use author-coloured strikethroughs, and changed lines receive margin bars. Because this changes line breaks and pagination, set the initial view here and use setLayoutView() or setShowTrackedChanges() for later changes.',
};
const DOCX_CURRENT_DATE: ApiOption = {
  name: 'currentDate',
  type: 'Date | number',
  def: 'load time',
  desc: 'Date used to resolve DATE and TIME fields. It participates in the retained layout variant, so pass it at load time when deterministic field values or pagination are required.',
};
const DOCX_LAYOUT_VIEW_OPTIONS: readonly ApiOption[] = [
  DOCX_SHOW_TRACKED_CHANGES,
  DOCX_CURRENT_DATE,
];
const DOCX_SLICE_LAYOUT: ApiOption = {
  name: 'sliceLayout',
  type: 'boolean',
  def: 'false',
  desc: 'In main mode, yield to the browser between pagination slices while load() still waits for the complete document. Worker mode already keeps pagination off the UI thread; without progressiveLayout this option has no additional effect there. Use progressiveLayout when opening pages should become available before full pagination finishes.',
  detailsHref: '/docx#progressive-layout',
  detailsLabel: 'Progressive layout guide',
};
const DOCX_LAYOUT_PROGRESS: ApiOption = {
  name: 'onLayoutProgress',
  type: '(progress: Readonly<{ committedUnits: number }>) => void',
  desc: 'Receive pagination telemetry from resumable layout passes. It implies sliced layout in main mode and is also delivered by progressiveLayout in main or worker mode; a non-progressive worker load does not publish intermediate progress. committedUnits can move backward while convergence revises provisional work; it counts pages, so use pageCount and the Viewer page callbacks for application navigation UI. Observer exceptions are reported once and never change the layout result.',
  emphasis: 'committedUnits can move backward while convergence revises provisional work',
  detailsHref: '/docx#progressive-layout',
  detailsLabel: 'Progressive layout guide',
};
const DOCX_LAYOUT_PARTIAL: ApiOption = {
  name: 'onLayoutPartial',
  type: '(progress: Readonly<{ availableUnits: number; totalUnits?: number; exact: boolean }>) => void',
  desc: 'Receive each later provisional page publication after the initial load publication. availableUnits is the pages available so far; totalUnits is omitted until DOCX pagination knows the final count, and exact is currently always false. Observer exceptions are reported once and never change the layout result.',
  detailsHref: '/docx#progressive-layout',
  detailsLabel: 'Progressive layout guide',
};
const DOCX_LAYOUT_COMPLETE: ApiOption = {
  name: 'onLayoutComplete',
  type: '(error?: unknown) => void',
  desc: 'With progressiveLayout enabled, called exactly once when a successful load reaches its authoritative full layout, even when load() itself waited for completion. A failure after an early publication is passed as the argument; a failure before the first publication rejects load() directly without calling this observer. Observer exceptions are reported once and never change the layout result.',
  detailsHref: '/docx#progressive-layout',
  detailsLabel: 'Progressive layout guide',
};
const PPTX_PROGRESSIVE_LAYOUT: ApiOption = {
  name: 'progressiveLayout',
  type: 'boolean',
  def: 'false',
  desc: 'Resolve load() when the opening slide is paintable and continue sequential preflight in the background. slideCount and the ScrollViewer extent are final from first paint; availableSlideCount is the paintable opening prefix. Works in main and worker modes.',
  emphasis: 'slideCount and the ScrollViewer extent are final from first paint',
  detailsHref: '/pptx#progressive-layout',
  detailsLabel: 'Progressive layout guide',
};
const PPTX_LAYOUT_PROGRESS: ApiOption = {
  name: 'onLayoutProgress',
  type: '(progress: Readonly<{ committedUnits: number }>) => void',
  desc: 'Called as the sequential preflight commits each paintable slide. committedUnits counts slides. Observer exceptions are reported once and never change the layout result.',
  detailsHref: '/pptx#progressive-layout',
  detailsLabel: 'Progressive layout guide',
};
const PPTX_LAYOUT_PARTIAL: ApiOption = {
  name: 'onLayoutPartial',
  type: '(progress: Readonly<{ availableUnits: number; totalUnits?: number; exact: boolean }>) => void',
  desc: 'Called for each additional paintable prefix after load() resolves. availableUnits counts paintable slides, totalUnits is the final slide count, and exact is false until completion. Observer exceptions are reported once and never change the layout result.',
  detailsHref: '/pptx#progressive-layout',
  detailsLabel: 'Progressive layout guide',
};
const PPTX_LAYOUT_COMPLETE: ApiOption = {
  name: 'onLayoutComplete',
  type: '(error?: unknown) => void',
  desc: 'With progressiveLayout enabled, called exactly once when a successful load makes every slide paintable, even when load() itself waited for completion. A failure after an early publication is passed as the argument; a failure before the first publication rejects load() directly without calling this observer. Observer exceptions are reported once and never change the layout result.',
  detailsHref: '/pptx#progressive-layout',
  detailsLabel: 'Progressive layout guide',
};
const PPTX_PROGRESSIVE_OPTIONS = [
  PPTX_PROGRESSIVE_LAYOUT,
  PPTX_LAYOUT_PROGRESS,
  PPTX_LAYOUT_PARTIAL,
  PPTX_LAYOUT_COMPLETE,
];

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

const FIND_HIGHLIGHT_COLORS: ApiOption = {
  name: 'findHighlightColors',
  type: '{ match?: string; active?: string }',
  def: 'yellow / orange',
  desc: 'CSS backgrounds for ordinary and active find matches. Values are applied verbatim; use an alpha color to keep the canvas text visible through the overlay.',
  emphasis: 'use an alpha color to keep the canvas text visible through the overlay.',
};

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
        PASSWORD,
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent text layer so users can select & copy slide text.' },
        { name: 'enableElementSelection', type: 'boolean', def: 'false', desc: 'Enable read-only slide-element selection with a non-editable outline and element context; no editor model is exposed.' },
        { name: 'elementHitTolerance', type: 'number', def: '6', desc: 'Straight-line hit tolerance in CSS pixels for element context clicks.' },
        { name: 'onSelectionContextChange', type: '(context: PptxSelectionContext | null) => void', desc: 'Receive bounded detached text-selection or element-click context for AI/MCP handoff. Text selection takes precedence; this callback does not enable element hit-testing.' },
        CONTEXT_MENU('PptxSelectionContext'),
        FIND_HIGHLIGHT_COLORS,
        { name: 'enableMediaPlayback', type: 'boolean', def: 'false', desc: 'Make embedded audio/video interactive (the viewer draws its own play chrome).' },
        { name: 'hiddenSlideMode', type: "'show' | 'skip' | 'dim'", def: "'show'", desc: 'How hidden slides (`<p:sld show="0">`, §19.3.1.38) are presented. `show` draws them like any other slide; `skip` makes sequential navigation (nextSlide/prevSlide and the initial load) jump over them while keeping absolute indices unchanged (an explicit goToSlide to a hidden slide is still honored); `dim` draws them under a translucent overlay (the PowerPoint thumbnail look).' },
        { name: 'hiddenSlideDim', type: 'Partial<DimOptions>', def: "{ color: '#ffffff', opacity: 0.6 }", desc: 'Overrides for the `dim` overlay, merged over the default white 60% wash. A partial so it stays in sync if DimOptions gains a field.' },
        ZIP,
        RESOURCE_LIMITS,
        IMAGE_RESOURCES,
        RESOURCE_METRICS,
        DEBUG,
        MATH,
        THREE_D,
        REGION_MAP,
        CHART_EX,
        TIFF,
        VIEWER_MODE,
        ...PPTX_PROGRESSIVE_OPTIONS,
        ZOOM_MIN_MAX,
        ON_SCALE_CHANGE,
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        { name: 'onSlideChange', type: '(index: number, total: number, layoutComplete: boolean) => void', desc: 'Called after a slide finishes rendering and again when progressive availability changes completion state.' },
        VIEWER_ON_ERROR,
      ],
      methods: [
        { sig: 'static fromPresentation(canvas, presentation, options?): Omit<PptxViewer, "load">', desc: 'Synchronously create a Viewer that borrows an already-loaded presentation. Render with goToSlide(); destroy() leaves the presentation open.' },
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load a Viewer-owned URL or ArrayBuffer. With progressiveLayout, resolve when the opening slide is paintable.' },
        { sig: 'goToSlide(index: number): Promise<void>', desc: 'Render a specific slide (0-indexed, clamped).' },
        { sig: 'nextSlide(): Promise<void>', desc: 'Advance one slide.' },
        { sig: 'prevSlide(): Promise<void>', desc: 'Go back one slide.' },
        ...zoomMethods(true),
        ...findMethods('PptxMatchLocation'),
        { sig: 'get slideIndex(): number', desc: 'Current slide index.' },
        { sig: 'get slideCount(): number', desc: 'Total slides (0 until loaded).' },
        { sig: 'get availableSlideCount(): number', desc: 'Paintable opening-slide prefix. Equals slideCount outside progressive loading.' },
        { sig: 'get layoutComplete(): boolean', desc: 'True only when every slide is paintable. It remains false if background preparation fails; waitUntilLayoutComplete() reports that failure.' },
        { sig: 'waitUntilLayoutComplete(): Promise<void>', desc: 'Wait until every slide is paintable; rejects if background preflight fails.' },
        { sig: 'get hiddenSlideMode(): "show" | "skip" | "dim"', desc: 'The current hidden-slide mode.' },
        { sig: 'setHiddenSlideMode(mode: "show" | "skip" | "dim"): Promise<void>', desc: 'Switch the hidden-slide mode at runtime and re-render. Entering `skip` while on a hidden slide advances to the nearest visible slide.' },
        { sig: 'get visibleSlideCount(): number', desc: 'Number of non-hidden slides. During progressive loading this is provisional until layoutComplete; the absolute slideCount remains unchanged.' },
        { sig: 'getNotes(slideIndex: number): string | null', desc: 'Speaker-notes text for a slide (0-based). During progressive loading the answer is authoritative only below availableSlideCount; await waitUntilLayoutComplete() before scanning the whole deck.' },
        { sig: 'get canvasElement(): HTMLCanvasElement', desc: 'The underlying canvas.' },
        { sig: 'getSelectionContext(options?: PptxSelectionContextOptions): PptxSelectionContext | null', desc: 'Return the current bounded, JSON-serializable text or element focus snapshot. Throws after destroy().' },
        RESOURCE_METRICS_METHOD,
        { sig: 'destroy(): void', desc: 'Tear down the worker and release resources.' },
      ],
    },
    {
      name: 'PptxPresentation',
      ctor: 'await PptxPresentation.load(source, options?)',
      note: 'Headless engine — parse once, render any slide into any canvas you supply (scroll views, thumbnail grids, master–detail).',
      options: [GFONTS, PASSWORD, WASM_URL, ZIP, RESOURCE_LIMITS, RESOURCE_METRICS, DEBUG, WORKER_TIMEOUT, MATH, THREE_D, REGION_MAP, CHART_EX, TIFF, MODE, ...PPTX_PROGRESSIVE_OPTIONS],
      methods: [
        { sig: 'static load(source, options?): Promise<PptxPresentation>', desc: 'Parse a deck from a URL or ArrayBuffer. With progressiveLayout, resolve when the opening slide is paintable.' },
        { sig: 'get slideCount(): number', desc: 'Total slides.' },
        { sig: 'get availableSlideCount(): number', desc: 'Paintable opening-slide prefix; slideCount remains final throughout.' },
        { sig: 'get layoutComplete(): boolean', desc: 'True only when every slide is paintable. It remains false if background preparation fails; waitUntilLayoutComplete() reports that failure.' },
        { sig: 'waitUntilLayoutComplete(): Promise<void>', desc: 'Wait until every slide is paintable; rejects if background preflight fails.' },
        { sig: 'renderSlide(canvas, index, opts?: { width?, dpr?, onTextRun?, dim? }): Promise<void>', desc: 'Render one slide into the given canvas at the given width. `onTextRun` receives each rendered segment as `PptxTextRunInfo`, including the source shape’s slide-local `shapeId`, optional frame flips, and zero-based table-cell row/column when authored, so callers can build a transparent selection overlay or stable shape mapping; `dim` (a DimOptions) paints a translucent wash over the finished slide (hidden-slide dimming). Equations render when a `math` engine was passed to `load`. Unavailable in `mode: "worker"` — use renderSlideToBitmap.' },
        { sig: 'renderSlideToBitmap(index, opts?: { width?, dpr?, dim? }): Promise<ImageBitmap>', desc: 'Render one slide and return it as an ImageBitmap (both modes; in worker mode slide paint, equations, ChartEx, 3-D charts and Region Maps run off the main thread). `dim` paints a translucent overlay over the slide (hidden-slide dimming). The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.', emphasis: 'The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.' },
        { sig: 'presentSlide(canvas, index, opts?: PresentSlideOptions): Promise<PresentationHandle>', desc: 'Render a slide and attach canvas-native audio/video playback, returning a handle with play() / pause() / destroy(). Initial render and media acquisition failures reject this Promise. PresentSlideOptions.onError observes decode or playback failures that occur only after the handle has been returned. Works in both modes — in `mode: "worker"` the base slide and text-run geometry are produced off-thread and the video overlay is composited on the main thread.' },
        { sig: 'getNotes(slideIndex: number): string | null', desc: 'Speaker-notes text for a slide (0-based; ECMA-376 §13.3.5). During progressive loading, null for an index at or beyond availableSlideCount means the slide is not ready, not necessarily that it has no notes. Await completion before a whole-deck scan.' },
        { sig: 'getComments(slideIndex: number): readonly Readonly<PptxComment>[]', desc: 'Detached comment threads for one slide. During progressive loading, results are authoritative only below availableSlideCount; await waitUntilLayoutComplete() before scanning the whole deck. Modern comments expose slide, drawing-element, or text-range anchors; classic comments retain their authored slide point.' },
        { sig: 'isHidden(slideIndex: number): boolean', desc: 'Whether a slide is authored as hidden. During progressive loading, results are authoritative only below availableSlideCount; await completion before scanning every slide.' },
        { sig: 'get slideWidth(): number', desc: 'Slide width in EMU (0 until loaded).' },
        { sig: 'get slideHeight(): number', desc: 'Slide height in EMU (0 until loaded).' },
        { sig: 'get mode(): "main" | "worker"', desc: 'The render mode this engine was loaded with. A borrowed engine’s mode decides whether slides render via renderSlide (main) or renderSlideToBitmap (worker).' },
        { sig: 'getElementContextAt(slideIndex, point, options?): Promise<PptxElementContext | null>', desc: 'Return compact context for the topmost transformed element frame at a slide-EMU point in either mode (line segments use tolerance). Includes master/layout/slide provenance, never editor tree indexes or mutable elements.' },
        { sig: 'getElementBoundsByIds(slideIndex, elementIds): Promise<readonly PptxElementBounds[]>', desc: 'Resolve authored DrawingML element ids to immutable slide geometry in one lazy slide read. Use this with modern-comment drawing or text anchors; it works in main and worker modes.' },
        RESOURCE_METRICS_METHOD,
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
        { name: 'zoomMin / zoomMax', type: 'number', def: '0.1 / 4', desc: 'Absolute zoom scale bounds (10%–400%). When width fit needs a smaller scale, that fitted scale remains reachable as the effective minimum.' },
        { name: 'refitOnResize', type: 'boolean', def: 'true', desc: 'Re-fit to the container width when it resizes. Set false to preserve an absolute scale independently of viewport width; explicit fitWidth() / fitPage() still work.' },
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent, selectable text layer per slide for native copy in both render modes.' },
        { name: 'comments', type: 'boolean | PptxCommentsOptions', def: 'false', desc: 'Show read-only slide comment targets, message icons, and built-in margin cards. Pass `cards: false` for an application-owned list that retains Viewer-owned target highlighting, or `markers: false` to hide idle message icons. The options object also controls resolved-thread visibility, side, and optional connectors. Theme cards and markers with CSS custom properties or documented classes on the Viewer container.', detailsHref: '/review-ui', detailsLabel: 'Comment UI guide' },
        { name: 'enableElementSelection', type: 'boolean', def: 'false', desc: 'Enable read-only element selection on mounted slide canvases with a non-editable outline and element context.' },
        { name: 'elementHitTolerance', type: 'number', def: '6', desc: 'Straight-line hit tolerance in CSS pixels.' },
        { name: 'onSelectionContextChange', type: '(context: PptxSelectionContext | null) => void', desc: 'Receive bounded detached text, selected-comment, or element context for external AI/MCP integrations. This callback does not enable element hit-testing.' },
        CONTEXT_MENU('PptxSelectionContext'),
        FIND_HIGHLIGHT_COLORS,
        { name: 'enableMediaPlayback', type: 'boolean', def: 'false', desc: 'Make embedded audio/video interactive inside the real viewport plus mediaOverscan. Other mounted slides remain static and selectable without allocating media blobs or RAF loops.' },
        { name: 'mediaOverscan', type: 'number', def: '1', desc: 'Slides beyond the real viewport that may keep interactive media handles. Independent from the general overscan used for mounted canvases/text overlays.' },
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        GFONTS,
        PASSWORD,
        ZIP,
        RESOURCE_LIMITS,
        IMAGE_RESOURCES,
        RESOURCE_METRICS,
        DEBUG,
        MATH,
        THREE_D,
        REGION_MAP,
        CHART_EX,
        TIFF,
        DPR,
        MODE,
        ...PPTX_PROGRESSIVE_OPTIONS,
        { name: 'onVisibleSlideChange', type: '(topIndex: number, total: number, layoutComplete: boolean) => void', desc: 'Fires when the top-most visible slide changes or progressive completion changes while the same slide remains visible.' },
        VIEWER_ON_ERROR,
      ],
      methods: [
        { sig: 'static fromPresentation(container, presentation, options?): Omit<PptxScrollViewer, "load">', desc: 'Synchronously create a Scroll Viewer that borrows one loaded presentation and lays out its initial virtual window.' },
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load a Viewer-owned deck. With progressiveLayout, render the opening paintable window while reserving the final scroll extent.' },
        { sig: 'scrollToSlide(index: number, opts?: { behavior?: "auto" | "smooth" }): void', desc: 'Scroll so slide index’s top edge sits at the viewport top (index clamped).' },
        { sig: 'goToComment(slideIndex: number, commentIndex: number, opts?: { behavior?: "auto" | "smooth" }): Promise<boolean>', desc: 'Reveal and highlight one entry from presentation.getComments(slideIndex). Resolves after modern element bounds or the authored slide point has been selected; returns false for an invalid locator or unresolved target.' },
        ...findMethods('PptxMatchLocation'),
        { sig: 'setScale(scale: number): void', desc: 'Set the absolute zoom scale at runtime (clamped to the effective zoom range, which includes a width fit below zoomMin). Flicker-free. View-only.' },
        { sig: 'relayout(): void', desc: 'Force a re-fit + re-mount of the visible window. Called automatically after load / resize / zoom; use it when the container resizes in a way a ResizeObserver cannot observe (e.g. a late web-font load). Idempotent.' },
        { sig: 'get slideCount(): number', desc: 'Total slides (0 until loaded).' },
        { sig: 'get availableSlideCount(): number', desc: 'Paintable opening-slide prefix; the full scroll extent already uses slideCount.' },
        { sig: 'get layoutComplete(): boolean', desc: 'True only when every slide is paintable. It remains false if background preparation fails; waitUntilLayoutComplete() reports that failure.' },
        { sig: 'waitUntilLayoutComplete(): Promise<void>', desc: 'Wait until every slide is paintable; rejects if background preflight fails.' },
        { sig: 'get topVisibleSlide(): number', desc: 'Index of the top-most visible slide.' },
        { sig: 'getSelectionContext(options?: PptxSelectionContextOptions): PptxSelectionContext | null', desc: 'Return the current mounted text selection, selected comment thread, or clicked-element context.' },
        RESOURCE_METRICS_METHOD,
        { sig: 'destroy(): void', desc: 'Tear down the DOM subtree. Destroys a self-loaded engine; a borrowed one is left intact.' },
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
        PASSWORD,
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent text layer for native selection & copy.' },
        ...DOCX_LAYOUT_VIEW_OPTIONS,
        { name: 'enableElementSelection', type: 'boolean', def: 'false', desc: 'Enable read-only picture, chart, and shape selection with a non-editable outline and element context. No editor model is added.' },
        { name: 'onSelectionContextChange', type: '(context: DocxSelectionContext | null) => void', desc: 'Receive bounded detached text or element context. This callback does not enable element hit-testing by itself.' },
        CONTEXT_MENU('DocxSelectionContext'),
        FIND_HIGHLIGHT_COLORS,
        ZIP,
        RESOURCE_LIMITS,
        IMAGE_RESOURCES,
        RESOURCE_METRICS,
        DEBUG,
        MATH,
        THREE_D,
        REGION_MAP,
        CHART_EX,
        TIFF,
        VIEWER_MODE,
        DOCX_PROGRESSIVE_LAYOUT,
        DOCX_SLICE_LAYOUT,
        DOCX_LAYOUT_PROGRESS,
        DOCX_LAYOUT_PARTIAL,
        DOCX_LAYOUT_COMPLETE,
        ZOOM_MIN_MAX,
        ON_SCALE_CHANGE,
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        { name: 'onPageChange', type: '(index: number, total: number, layoutComplete: boolean) => void', desc: 'Called after a page finishes rendering and again when a progressive page-count publication changes total. While layoutComplete is false, total is the pages available so far rather than the final count.', emphasis: 'While layoutComplete is false, total is the pages available so far rather than the final count.', detailsHref: '/docx#progressive-layout', detailsLabel: 'Progressive layout guide' },
        VIEWER_ON_ERROR,
      ],
      methods: [
        { sig: 'static fromDocument(canvas, document, options?): Omit<DocxViewer, "load">', desc: 'Synchronously create a Viewer that borrows an already-loaded document. Render with goToPage(); destroy() leaves the document open.' },
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load a Viewer-owned URL or ArrayBuffer. With progressiveLayout, resolve when the opening page is paintable while pagination continues in the background.' },
        { sig: 'goToPage(index: number): Promise<void>', desc: 'Render a specific page (0-indexed, clamped). During progressive layout, a requested page beyond the published prefix waits with the loading indicator until it becomes available.' },
        { sig: 'nextPage(): Promise<void>', desc: 'Advance one page.' },
        { sig: 'prevPage(): Promise<void>', desc: 'Go back one page.' },
        { sig: 'setShowTrackedChanges(value: boolean): Promise<void>', desc: 'Switch between the final view (false, default) and the tracked-change markup view (true) at runtime, re-rendering the current page against the selected layout variant.' },
        ...zoomMethods(true),
        ...findMethods('DocxMatchLocation'),
        { sig: 'get pageCount(): number', desc: 'Pages available so far (0 until loaded); authoritative only when layoutComplete is true.' },
        { sig: 'get currentPage(): number', desc: 'Current page index.' },
        { sig: 'get layoutComplete(): boolean', desc: 'True only after the authoritative document layout succeeds. It is false while progressive layout is publishing pages and remains false if background pagination fails; waitUntilLayoutComplete() reports that failure.' },
        { sig: 'waitUntilLayoutComplete(): Promise<void>', desc: 'Wait for the authoritative full layout before operations that require the final page count. Rejects if background pagination fails after load() resolved.' },
        { sig: 'get canvasElement(): HTMLCanvasElement', desc: 'The underlying canvas.' },
        { sig: 'getSelectionContext(options?: DocxSelectionContextOptions): DocxSelectionContext | null', desc: 'Return the current bounded native-text or clicked-element snapshot. Throws after destroy().' },
        RESOURCE_METRICS_METHOD,
        { sig: 'destroy(): void', desc: 'Tear down the worker and release resources.' },
      ],
    },
    {
      name: 'DocxDocument',
      ctor: 'await DocxDocument.load(source, options?)',
      note: 'Headless engine — render any page into any canvas you supply.',
      options: [GFONTS, PASSWORD, WASM_URL, ZIP, RESOURCE_LIMITS, RESOURCE_METRICS, DEBUG, WORKER_TIMEOUT, MATH, THREE_D, REGION_MAP, CHART_EX, TIFF, MODE, ...DOCX_LAYOUT_VIEW_OPTIONS, DOCX_PROGRESSIVE_LAYOUT, DOCX_SLICE_LAYOUT, DOCX_LAYOUT_PROGRESS, DOCX_LAYOUT_PARTIAL, DOCX_LAYOUT_COMPLETE],
      methods: [
        { sig: 'static load(source, options?): Promise<DocxDocument>', desc: 'Parse a document from a URL or ArrayBuffer. With progressiveLayout, resolve when the opening pages are paintable while pagination continues in the background.' },
        { sig: 'get comments(): readonly Readonly<DocComment>[]', desc: 'Immutable detached comments and replies stored in the document.' },
        { sig: 'get revisions(): readonly Readonly<DocRevision>[]', desc: 'Immutable detached WordprocessingML body-story insertion, deletion, and move records. This is the current DOCX change-history API, not a cross-format revision contract.' },
        { sig: 'commentAnchorRanges(): readonly CommentAnchorRange[]', desc: 'Logical comment ranges for the currently available page prefix. Await waitUntilLayoutComplete() before treating this as a full-document projection.' },
        { sig: 'getCommentThreads(pageIndex: number, options?: DocxPageCommentThreadsOptions): Promise<readonly Readonly<ResolvedDocxCommentThread>[]>', desc: 'Resolve top-level threads with rendered anchor geometry on one page. Cross-page ranges and repeating stories can occur in more than one page result; each result contains only that page’s rectangles.' },
        { sig: 'revisionAnchorRanges(): readonly RevisionAnchorRange[]', desc: 'Logical tracked-change ranges for the currently available page prefix. Await waitUntilLayoutComplete() before treating this as a full-document projection.' },
        { sig: 'getBookmarkPage(name: string): number | undefined', desc: 'Resolve a bookmark in the currently available page prefix. Await waitUntilLayoutComplete() before concluding that an unresolved name is absent from the document.' },
        { sig: 'collectPageRuns(index, options?): Promise<DocxTextRunInfo[]>', desc: 'Collect the same immutable text-run geometry emitted while rendering one page.' },
        { sig: 'get pageCount(): number', desc: 'Pages available so far; authoritative only when layoutComplete is true.' },
        { sig: 'get layoutComplete(): boolean', desc: 'True only after the authoritative document layout succeeds. It is false while progressive layout is publishing pages and remains false if background pagination fails; waitUntilLayoutComplete() reports that failure.' },
        { sig: 'waitUntilLayoutComplete(): Promise<void>', desc: 'Wait for the authoritative full layout. Rejects if background pagination fails after load() resolved.' },
        { sig: 'pageSize(pageIndex: number): { widthPt, heightPt }', desc: 'Page size in pt for a page (ECMA-376 §17.6.13 / §17.6.11 — per section, so a mixed portrait/landscape document returns different sizes per page). Available in both modes; index is clamped. `{ 0, 0 }` means "not loaded". Returns a fresh object per call.', emphasis: '`{ 0, 0 }` means "not loaded".' },
        { sig: 'get mode(): "main" | "worker"', desc: 'The render mode this engine was loaded with. A borrowed engine’s mode decides whether pages render via renderPage (main) or renderPageToBitmap (worker).' },
        { sig: 'setLayoutView(view?: { showTrackedChanges?, currentDate? }): Promise<void>', desc: 'Select the layout variant used by geometry and paint. In worker mode, resolves after matching page metadata is ready and installed atomically.' },
        { sig: 'renderPage(canvas, index, opts?: { width?, dpr?, onTextRun? }): Promise<void>', desc: 'Render one page into the given canvas. `onTextRun` receives each segment as `DocxTextRunInfo`, including the authored `w14:paraId` as `paragraphId` when present. Unavailable in `mode: "worker"` — use renderPageToBitmap.' },
        { sig: 'renderPageToBitmap(index, opts?: { width?, dpr?, onTextRun? }): Promise<ImageBitmap>', desc: 'Render one page and return it as an ImageBitmap (both modes; in worker mode page paint, equations, ChartEx, 3-D charts and Region Maps run off the main thread and return the same text-run stream beside the bitmap). The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.', emphasis: 'The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.' },
        { sig: 'getElementContextAt(pageIndex, point, options?): Promise<DocxElementContext | null>', desc: 'Return compact context for the topmost rendered picture, chart, or shape—including inline content—at a physical-page-point coordinate in either mode.' },
        RESOURCE_METRICS_METHOD,
        { sig: 'destroy(): void', desc: 'Release the worker.' },
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
        { name: 'zoomMin / zoomMax', type: 'number', def: '0.1 / 4', desc: 'Absolute zoom scale bounds (10%–400%). When width fit needs a smaller scale, that fitted scale remains reachable as the effective minimum.' },
        { name: 'refitOnResize', type: 'boolean', def: 'true', desc: 'Re-fit to the container width when it resizes. Set false to preserve an absolute scale independently of viewport width; explicit fitWidth() / fitPage() still work.' },
        { name: 'enableTextSelection', type: 'boolean', def: 'false', desc: 'Overlay a transparent, selectable text layer per page for native copy in both render modes.' },
        ...DOCX_LAYOUT_VIEW_OPTIONS,
        { name: 'comments', type: 'boolean | DocxCommentsOptions', def: 'false', desc: 'Show read-only document comment highlights, message icons, and built-in margin cards. Pass `cards: false` for an application-owned list that retains Viewer-owned range highlighting, or `markers: false` to hide only the icons. The options object also controls resolved-thread visibility, side, and optional connectors. Theme cards, highlights, and markers with CSS custom properties or documented classes on the Viewer container.', detailsHref: '/review-ui', detailsLabel: 'Comment UI guide' },
        { name: 'enableElementSelection', type: 'boolean', def: 'false', desc: 'Enable read-only drawing selection on mounted pages with a non-editable outline and element context.' },
        { name: 'onSelectionContextChange', type: '(context: DocxSelectionContext | null) => void', desc: 'Receive bounded detached text, selected-comment, or element context for external AI/MCP integrations. This callback does not enable element hit-testing.' },
        CONTEXT_MENU('DocxSelectionContext'),
        FIND_HIGHLIGHT_COLORS,
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        GFONTS,
        PASSWORD,
        ZIP,
        RESOURCE_LIMITS,
        IMAGE_RESOURCES,
        RESOURCE_METRICS,
        DEBUG,
        MATH,
        THREE_D,
        REGION_MAP,
        CHART_EX,
        TIFF,
        DPR,
        MODE,
        DOCX_PROGRESSIVE_LAYOUT,
        DOCX_SLICE_LAYOUT,
        DOCX_LAYOUT_PROGRESS,
        DOCX_LAYOUT_PARTIAL,
        DOCX_LAYOUT_COMPLETE,
        { name: 'onVisiblePageChange', type: '(topIndex: number, total: number, layoutComplete: boolean) => void', desc: 'Fires when the top-most visible page changes and again when a progressive page-count publication changes total, even if the same page remains visible. While layoutComplete is false, total is the pages available so far rather than the final count.', emphasis: 'While layoutComplete is false, total is the pages available so far rather than the final count.', detailsHref: '/docx#progressive-layout', detailsLabel: 'Progressive layout guide' },
        VIEWER_ON_ERROR,
      ],
      methods: [
        { sig: 'static fromDocument(container, document, options?): Omit<DocxScrollViewer, "load">', desc: 'Synchronously create a Scroll Viewer that borrows one loaded document and lays out its initial virtual window.' },
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load a Viewer-owned document. With progressiveLayout, resolve when the opening window is paintable and grow the scroll surface as pagination continues.' },
        { sig: 'scrollToPage(index: number, opts?: { behavior?: "auto" | "smooth" }): void', desc: 'Scroll so page index’s top edge sits at the viewport top (index clamped).' },
        { sig: 'goToComment(commentId: string, opts?: { pageIndex?: number; behavior?: "auto" | "smooth" }): Promise<boolean>', desc: 'Reveal and highlight a top-level DOCX comment. Omit pageIndex for its first rendered occurrence, or pass the page returned by getCommentThreads() to select that specific occurrence. Returns false when the locator has no rendered anchor.' },
        { sig: 'setShowTrackedChanges(value: boolean): Promise<void>', desc: 'Switch between the final view (false, default) and the tracked-change markup view (true) at runtime. Resolves after matching layout geometry is ready and mounted pages have been refreshed.' },
        ...findMethods('DocxMatchLocation'),
        { sig: 'setScale(scale: number): void', desc: 'Set the absolute zoom scale at runtime (clamped to the effective zoom range, which includes a width fit below zoomMin). Flicker-free. View-only.' },
        { sig: 'relayout(): void', desc: 'Force a re-fit + re-mount of the visible window. Called automatically after load / resize / zoom; use it when the container resizes in a way a ResizeObserver cannot observe (e.g. a late web-font load). Idempotent.' },
        { sig: 'get pageCount(): number', desc: 'Pages available so far (0 until loaded); authoritative only when layoutComplete is true.' },
        { sig: 'get topVisiblePage(): number', desc: 'Index of the top-most visible page.' },
        { sig: 'get layoutComplete(): boolean', desc: 'True only after the authoritative document layout succeeds. It is false while progressive layout is publishing pages and remains false if background pagination fails; waitUntilLayoutComplete() reports that failure.' },
        { sig: 'waitUntilLayoutComplete(): Promise<void>', desc: 'Wait for the authoritative full layout before operations that require the final page count. Rejects if background pagination fails after load() resolved.' },
        { sig: 'getSelectionContext(options?: DocxSelectionContextOptions): DocxSelectionContext | null', desc: 'Return the current mounted text selection, selected comment thread, or clicked-element context.' },
        RESOURCE_METRICS_METHOD,
        { sig: 'destroy(): void', desc: 'Tear down the DOM subtree. Destroys a self-loaded engine; a borrowed one is left intact.' },
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
        { name: 'resizable', type: 'boolean', def: 'true', desc: 'Allow resizing columns/rows by dragging header borders. View-only — it changes the on-screen view only and never modifies the loaded file. Set false to disable.', emphasis: 'View-only — it changes the on-screen view only and never modifies the loaded file.' },
        { name: 'showScrollbars', type: 'boolean', def: 'true', desc: 'Show native worksheet scrollbars. Set false only when the host supplies another viewport navigation UI.' },
        { name: 'selectionColor', type: 'string', def: "'#1a73e8'", desc: 'Accent color for the cell-selection rectangle (any CSS color). The fill is the same color at 8% opacity.' },
        { name: 'enableElementSelection', type: 'boolean', def: 'false', desc: 'Enable read-only chart, picture, and shape selection with a non-editable outline and element context, without changing the underlying cell selection.' },
        { name: 'comments', type: 'boolean | XlsxCommentsOptions', def: 'true', desc: 'Show authored cell note or threaded-comment markers and their anchored read-only popup. Pass an options object to control resolved-thread visibility. Theme the popup with documented CSS custom properties or classes on the Viewer container.', detailsHref: '/review-ui', detailsLabel: 'Comment UI guide' },
        FIND_HIGHLIGHT_COLORS,
        { name: 'hiddenSheetMode', type: "'show' | 'skip' | 'dim'", def: "'show'", desc: 'How hidden / very-hidden sheets (`<sheet state>`, §18.2.19) appear in the tab bar. `show` renders a tab like any other; `skip` hides the tab (`display:none`) and makes sequential navigation jump over it; `dim` renders the tab at reduced opacity. Mirrors pptx `hiddenSlideMode`.' },
        GFONTS,
        PASSWORD,
        ZIP,
        RESOURCE_LIMITS,
        IMAGE_RESOURCES,
        RESOURCE_METRICS,
        DEBUG,
        MATH,
        THREE_D,
        REGION_MAP,
        CHART_EX,
        TIFF,
        VIEWER_MODE,
        ON_SCALE_CHANGE,
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        { name: 'onReady', type: '(sheetNames: string[]) => void', desc: 'Called once the workbook is parsed.' },
        { name: 'onSheetChange', type: '(index: number, total: number) => void', desc: 'Called when the active sheet changes; `total` is the sheet count. Read the name via `sheetNames[index]`.' },
        { name: 'onSelectionStateChange', type: '(sel: XlsxSelectionState | null) => void', desc: 'Called only when canonical selection state changes; geometry, ActiveCell, extension anchor, and multiple areas remain distinct.' },
        { name: 'onSelectionContextChange', type: '(context: XlsxSelectionContext | null) => void', desc: 'Receive bounded detached selected-cell content, including attached comments, or element context for read-only AI/MCP handoff. This callback does not enable element hit-testing; rapid changes are coalesced per animation frame.' },
        CONTEXT_MENU('XlsxSelectionContext'),
        { name: 'onViewportChange', type: '(offset: XlsxViewportOffset) => void', desc: 'Called with the clamped logical CSS-pixel offset after the active viewport moves. Horizontal x is measured from column A in both LTR and RTL sheets.' },
        VIEWER_ON_ERROR,
      ],
      methods: [
        { sig: 'static fromWorkbook(container, workbook, options?): Omit<XlsxViewer, "load">', desc: 'Synchronously create a full Workbook Viewer that borrows one loaded workbook and starts its initial sheet display.' },
        { sig: 'load(source: string | ArrayBuffer): Promise<void>', desc: 'Load a Viewer-owned workbook and render the first sheet.' },
        { sig: 'goToSheet(index: number): Promise<void>', desc: 'Show a specific sheet (0-indexed, clamped).' },
        { sig: 'nextSheet(): Promise<void>', desc: 'Advance one sheet.' },
        { sig: 'prevSheet(): Promise<void>', desc: 'Go back one sheet.' },
        { sig: 'get sheetIndex(): number', desc: 'Current sheet index.' },
        { sig: 'get sheetCount(): number', desc: 'Total sheets (0 until loaded).' },
        { sig: 'get sheetNames(): string[]', desc: 'Names of all sheets.' },
        { sig: 'getComments(): readonly Readonly<XlsxComment>[]', desc: 'Detached notes and threaded comments for the current sheet, in authored order.' },
        { sig: 'goToComment(sheetIndex: number, cellRef: string, options?: XlsxScrollToCellOptions): Promise<boolean>', desc: 'Switch to the explicit sheet, reveal the commented cell, and select it. Returns false when the sheet or cell comment locator is invalid.' },
        { sig: 'get selectionState(): XlsxSelectionState | null', desc: 'Detached canonical state for the current selection.' },
        { sig: 'setSelection(input: string | XlsxSelectionState | null): void', desc: 'Set one A1 area, a complete canonical selection state, or clear the selection. A1 endpoint order does not encode ActiveCell.' },
        { sig: 'getSelectionContext(options?: { maxCells?: number; maxTextCharacters?: number }): XlsxSelectionContext | null', desc: 'Return bounded `{ kind: "range" }` cell content, including attached comments, or `{ kind: "element" }` clicked-object context for read-only AI/MCP handoff.' },
        { sig: 'copySelection(): Promise<XlsxCopyResult>', desc: 'Copy a bounded TSV and report copied, resource-limit, unsupported-multiple-area, or Clipboard API outcomes.' },
        { sig: 'getViewportOffset(): XlsxViewportOffset', desc: 'Return the active sheet viewport offset in logical CSS pixels.' },
        { sig: 'setViewportOffset(offset: XlsxViewportOffset): Promise<void>', desc: 'Move the active viewport to a finite offset, clamped to the used scroll extent.' },
        { sig: 'scrollToCell(ref: string, options?: XlsxScrollToCellOptions): Promise<void>', desc: 'Scroll a cell reference into view, optionally aligning it to the start, center, end, or nearest edge.' },
        { sig: 'getCellViewportRect(cell: CellAddress | string): XlsxCellViewportRect | null', desc: 'Return one cell’s CSS-pixel bounds relative to the worksheet viewport. Use it to anchor application-owned comments or annotations.' },
        { sig: 'relayout(): Promise<void>', desc: 'Re-read the viewport box, clamp the current offset, and render again after an external layout change.' },
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
        RESOURCE_METRICS_METHOD,
        { sig: 'destroy(): void', desc: 'Tear down the DOM subtree. Destroys a self-loaded workbook; a borrowed one remains caller-owned.' },
      ],
    },
    {
      name: 'XlsxSheetViewer',
      ctor: 'new XlsxSheetViewer(canvas: HTMLCanvasElement, options?: XlsxSheetViewerOptions)',
      note: 'Canvas-mounted active-sheet viewport. It uses the caller canvas and the same sheet rendering, selection, find and navigation mechanics as XlsxViewer, but creates no sheet-tab/footer chrome. As a secondary convenience, its load method can explicitly reuse this Excel-style surface for one CSV, TSV, or generic delimited-text source; this does not add delimited-text loading to XlsxWorkbook or XlsxViewer, and every field remains text. Native worksheet scrollbars are visible by default. DOM chrome, styles and listeners follow canvas.ownerDocument, so a parent page can mount borrowed workbook sheets into same-origin popup canvases.',
      options: [
        { name: 'cellScale', type: 'number', def: '1', desc: 'Scale factor for cell/header dimensions (0.5 = half size).' },
        { name: 'zoomMin / zoomMax', type: 'number', def: '0.1 / 4', desc: 'Zoom bounds as scale factors (10%–400%).' },
        { name: 'resizable', type: 'boolean', def: 'true', desc: 'Allow resizing columns/rows by dragging header borders. View-only.' },
        { name: 'showScrollbars', type: 'boolean', def: 'true', desc: 'Show native worksheet scrollbars. Set false only when the host supplies another viewport navigation UI.' },
        { name: 'selectionColor', type: 'string', def: "'#1a73e8'", desc: 'Accent color for the cell-selection rectangle.' },
        { name: 'enableElementSelection', type: 'boolean', def: 'false', desc: 'Enable read-only chart, picture, and shape selection with a non-editable outline and element context.' },
        { name: 'comments', type: 'boolean | XlsxCommentsOptions', def: 'true', desc: 'Show authored cell note or threaded-comment markers and their anchored read-only popup. Pass an options object to control resolved-thread visibility. Theme the popup with documented CSS custom properties or classes on the Viewer container.', detailsHref: '/review-ui', detailsLabel: 'Comment UI guide' },
        FIND_HIGHLIGHT_COLORS,
        { name: 'hiddenSheetMode', type: "'show' | 'skip' | 'dim'", def: "'show'", desc: 'Controls sequential navigation and hidden-sheet visibility without adding tab chrome.' },
        { name: 'onViewportChange', type: '(offset: XlsxViewportOffset) => void', desc: 'Called with the clamped logical CSS-pixel offset after the active viewport moves. Horizontal x is measured from column A independently of browser RTL scrollLeft conventions.' },
        GFONTS,
        PASSWORD,
        WASM_URL,
        ZIP,
        RESOURCE_LIMITS,
        IMAGE_RESOURCES,
        RESOURCE_METRICS,
        DEBUG,
        WORKER_TIMEOUT,
        MATH,
        THREE_D,
        REGION_MAP,
        CHART_EX,
        TIFF,
        VIEWER_MODE,
        ON_SCALE_CHANGE,
        ON_HYPERLINK_CLICK,
        ENABLE_HYPERLINKS,
        { name: 'onReady', type: '(sheetNames: string[]) => void', desc: 'Called once the workbook is parsed.' },
        { name: 'onSheetChange', type: '(index: number, total: number) => void', desc: 'Called when the active sheet changes.' },
        { name: 'onSelectionStateChange', type: '(sel: XlsxSelectionState | null) => void', desc: 'Called only when canonical selection state changes.' },
        { name: 'onSelectionContextChange', type: '(context: XlsxSelectionContext | null) => void', desc: 'Receive bounded detached selected-cell content, including attached comments, or element context, coalesced per animation frame. This callback does not enable element hit-testing.' },
        CONTEXT_MENU('XlsxSelectionContext'),
        VIEWER_ON_ERROR,
      ],
      methods: [
        { sig: 'static fromWorkbook(canvas, workbook, options?): Omit<XlsxSheetViewer, "load">', desc: 'Synchronously attach a borrowed workbook without materializing a sheet. Await goToSheet(index) to render only the requested sheet.' },
        { sig: 'load(source: string | ArrayBuffer, options?: XlsxSheetLoadOptions): Promise<void>', desc: 'Load a Viewer-owned XLSX source by default, or explicitly preview CSV, TSV, or generic delimited text. Use { format: "delimited-text", delimiter } for formats such as .txt, .dat, and .psv. String sources are fetched URLs rather than raw text or filesystem paths; use File.arrayBuffer() for a browser File. Delimited fields stay text with no format or value inference. Reload replacement, callbacks, worker rendering, and destroy ownership are the same as XLSX loading.' },
        { sig: 'goToSheet(index: number): Promise<void>', desc: 'Show a specific sheet (0-indexed, clamped).' },
        { sig: 'nextSheet(): Promise<void>', desc: 'Advance one sheet.' },
        { sig: 'prevSheet(): Promise<void>', desc: 'Go back one sheet.' },
        { sig: 'get sheetIndex(): number', desc: 'Current sheet index.' },
        { sig: 'get sheetCount(): number', desc: 'Total sheets (0 until loaded).' },
        { sig: 'get sheetNames(): string[]', desc: 'Names of all sheets.' },
        { sig: 'getComments(): readonly Readonly<XlsxComment>[]', desc: 'Detached notes and threaded comments for the current sheet, in authored order.' },
        { sig: 'goToComment(sheetIndex: number, cellRef: string, options?: XlsxScrollToCellOptions): Promise<boolean>', desc: 'Switch to the explicit sheet, reveal the commented cell, and select it. Returns false when the sheet or cell comment locator is invalid.' },
        { sig: 'getViewportOffset(): XlsxViewportOffset', desc: 'Read the logical start-anchored viewport offset in CSS pixels at the current scale.' },
        { sig: 'setViewportOffset(offset: XlsxViewportOffset): Promise<void>', desc: 'Move to a finite logical offset, clamped to the used scroll extent.' },
        { sig: 'scrollToCell(ref: string, options?: { align?: "nearest" | "start" | "center" | "end" }): Promise<void>', desc: 'Move the viewport to an A1 cell reference with the requested alignment.' },
        { sig: 'getCellViewportRect(cell: CellAddress | string): XlsxCellViewportRect | null', desc: 'Return one cell’s CSS-pixel bounds relative to the worksheet viewport. Use it to anchor application-owned comments or annotations.' },
        { sig: 'relayout(): Promise<void>', desc: 'Re-read the canvas CSS box and repaint the current viewport.' },
        ...zoomMethods(false),
        ...findMethods('XlsxMatchLocation'),
        { sig: 'get selectionState(): XlsxSelectionState | null', desc: 'Detached canonical state for the current selection.' },
        { sig: 'setSelection(input: string | XlsxSelectionState | null): void', desc: 'Set an A1 area, a complete canonical state, or clear the selection.' },
        { sig: 'getSelectionContext(options?: { maxCells?: number; maxTextCharacters?: number }): XlsxSelectionContext | null', desc: 'Return bounded selected-cell content, including attached comments, or clicked chart, picture, or shape context for AI/MCP use.' },
        { sig: 'copySelection(): Promise<XlsxCopyResult>', desc: 'Copy bounded TSV and return an observable result.' },
        { sig: 'getCellAt(clientX: number, clientY: number): CellAddress | null', desc: 'Hit-test a viewport coordinate to a cell address.' },
        { sig: 'get canvasElement(): HTMLCanvasElement', desc: 'The caller-owned canvas used by the viewer.' },
        RESOURCE_METRICS_METHOD,
        { sig: 'destroy(): void', desc: 'Permanently close the viewer and restore the caller canvas. A workbook borrowed through fromWorkbook() remains caller-owned and is not destroyed.' },
      ],
    },
    {
      name: 'XlsxWorkbook',
      ctor: 'await XlsxWorkbook.load(source, options?)',
      note: 'Headless engine — parse once, render any sheet viewport into any canvas you supply.',
      options: [GFONTS, PASSWORD, WASM_URL, ZIP, RESOURCE_LIMITS, RESOURCE_METRICS, DEBUG, WORKER_TIMEOUT, MATH, THREE_D, REGION_MAP, CHART_EX, TIFF, MODE],
      methods: [
        { sig: 'static load(source, options?): Promise<XlsxWorkbook>', desc: 'Parse a workbook from a URL or ArrayBuffer.' },
        { sig: 'get sheetNames(): string[]', desc: 'Names of all sheets.' },
        { sig: 'get sheetCount(): number', desc: 'Total sheets.' },
        { sig: 'get mode(): "main" | "worker"', desc: 'The render mode owned by this loaded workbook.' },
        { sig: 'getWorksheet(sheetIndex): Promise<Worksheet>', desc: 'Parse and return one worksheet model. Saved pivot-table facts are exposed read-only via `Worksheet.pivotTables`, with skipped malformed parts reported through `Worksheet.pivotDiagnostics`; saved worksheet cells and styles remain authoritative.' },
        { sig: 'getComments(sheetIndex: number): Promise<readonly Readonly<XlsxComment>[]>', desc: 'Return a detached snapshot of comments for one lazily materialized sheet, in authored order.' },
        { sig: 'renderViewport(canvas, sheetIndex, viewport, opts?: XlsxRenderViewportOptions): Promise<void>', desc: 'Render a row/col window of a sheet into the given canvas. `onTextRun` receives each text cell as `XlsxTextRunInfo` with required `sheetName` and A1 `cellRef` identity. Image bytes and decoded-image caches stay owned by the workbook. Equations in shapes render when a `math` engine was passed to `load`. Unavailable in `mode: "worker"` — use renderViewportToBitmap.' },
        { sig: 'renderViewportToBitmap(sheetIndex, viewport, opts: RenderViewportToBitmapOptions): Promise<ImageBitmap>', desc: 'Render a sheet viewport and return it as an ImageBitmap (both modes; in worker mode viewport paint, equations, ChartEx, 3-D charts and Region Maps run off the main thread). `width` and `height` are required — a worker has no DOM element to measure. The bitmap is caller-owned: pass it to `transferFromImageBitmap` (which consumes it) or call `bitmap.close()`.', emphasis: '`width` and `height` are required — a worker has no DOM element to measure.' },
        { sig: 'resolveValidationList(sheetIndex, formula1): Promise<ResolvedList>', desc: 'Resolve a list-type data-validation `formula1` (ECMA-376 §18.3.1.32) into the allowed values to display — inline quoted list, a range reference (each cell’s display string), or `{ kind: \'formula\' }` for named ranges. Read-only.' },
        RESOURCE_METRICS_METHOD,
        { sig: 'destroy(): void', desc: 'Release the worker.' },
      ],
    },
  ],
};
