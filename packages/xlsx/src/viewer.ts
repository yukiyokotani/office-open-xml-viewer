import {
  XlsxWorkbook,
  acquireXlsxWorksheetPreview,
  retainXlsxWorksheetReference,
  loadXlsxSheetSource,
  prepareXlsxViewerRowHeights,
  releaseXlsxViewerProjection,
  retainXlsxViewerFonts,
} from './workbook.js';
import type { LoadOptions } from './workbook.js';
import type { ViewportRange, Worksheet, XlsxChromeColors, XlsxComment } from './types.js';
import type { FindHighlightColors, HyperlinkTarget, FindMatch, FindMatchesOptions, OoxmlResourceMetrics, ViewerContextMenuEvent, ZoomableViewer } from '@silurus/ooxml-core';
import { nextVisibleIndex, resolveVisibleIndex, countVisible, anchoredZoomOffset, openExternalHyperlink, nextZoomStep, prevZoomStep, fitScale } from '@silurus/ooxml-core';
import {
  CallerCanvasMount,
  resolveCanvasViewerMode,
  type CanvasViewerRenderMode,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import {
  HEADER_W,
  HEADER_H,
  invalidateAutoRowHeights,
  invalidateSheetRenderCache,
  getGridGeometryForWorksheet,
  rtlMirrorX,
} from './renderer.js';
import { parseA1 } from './a1.js';
import { inheritWorksheetPreviewBounds } from './internal/worksheet-content-bounds.js';
import { inheritWorksheetPolicy } from './worksheet-policy-context.js';
import {
  anchorBaselinePreviewBlocker,
  viewportPreviewBlocker,
  type ViewportPreviewBlocker,
} from './internal/worksheet-preview-eligibility.js';
import {
  bindInitialAnchorSizes,
  captureInitialAnchorSizes,
  initialAnchorRowCoverage,
  releaseInitialAnchorSizeReference,
  type InitialAnchorSizeReference,
} from './internal/initial-anchor-sizes.js';
import type {
  CellAddress,
  XlsxSelectionContext,
  XlsxSelectionContextOptions,
  XlsxElementContext,
  XlsxSelectionInput,
  XlsxSelectionState,
} from './selection.js';
import {
  hitTestXlsxElementContext,
  limitXlsxElementContext,
  type XlsxElementHitViewport,
} from './element-context.js';
import {
  normalizeSelectionState,
  selectionStateFromReference,
  selectionStatesEqual,
} from './selection.js';
export type { CellAddress } from './selection.js';
import type { XlsxMatchLocation } from './find.js';
import type { XlsxCommentsOptions } from './comment-card.js';
import { withViewerRenderContext } from './worker-protocol.js';
import { SheetViewEdits } from './internal/viewer/sheet-view-edits.js';
import { OutlineGutter } from './internal/viewer/outline-gutter.js';
import { SheetTabBar } from './internal/viewer/sheet-tab-bar.js';
import { ZoomControl } from './internal/viewer/zoom-control.js';
import { ValidationPanel } from './internal/viewer/validation-panel.js';
import { HyperlinkDispatcher } from './internal/viewer/hyperlink-dispatcher.js';
import { FindAdapter } from './internal/viewer/find-adapter.js';
import { CopyController } from './internal/viewer/copy-controller.js';
import { SelectionInput } from './internal/viewer/selection-input.js';
import { SelectionOverlay } from './internal/viewer/selection-overlay.js';
import { SelectionNotifier } from './internal/viewer/selection-notifier.js';
import { SelectionContextReader } from './internal/viewer/selection-context.js';
import { ChromeTheme } from './internal/viewer/chrome-theme.js';
import {
  COMMENT_POPUP_MAX_H,
  COMMENT_POPUP_MAX_W,
  CommentPopup,
  createCommentMap,
} from './internal/viewer/comment-popup.js';
import type { OutlineAxis } from './outline.js';
import { GridGeometry } from './internal/grid-geometry.js';
import {
  SheetAcquisition,
  SheetRenderDispatcher,
  SelectionController,
  ViewportState,
  createSheetViewModel,
  type SheetSelectionMode,
} from './internal/sheet-viewer-runtime.js';
import { CanvasSurface, SheetOverlayHost } from './internal/sheet-surface.js';
import { withXlsxRenderCommitGuard } from './render-orchestrator.js';
import {
  DEFAULT_WORKSHEET_CONTENT_MINIMUMS,
  worksheetContentBounds,
} from './internal/worksheet-content-bounds.js';
import type { XlsxSheetLoadOptions } from './delimited-text.js';

export type { XlsxSheetLoadOptions } from './delimited-text.js';

const borrowedWorkbookOption = Symbol('XlsxViewer.borrowedWorkbook');
/** @internal Shared source-loading hook for the two XLSX viewer facades. */
const loadXlsxViewerSource = Symbol('XlsxViewer.loadSource');
// Re-exported for the existing xlsx zoom tests (resize-zoom.test.ts imports it
// from this module) and any consumer that referenced it here before it moved to
// @silurus/ooxml-core. The single source of truth is core (design §5.2).
export { zoomStepScale } from '@silurus/ooxml-core';

/** Max width of the list-validation dropdown panel (CSS px). */
const VALIDATION_PANEL_MAX_W = 240;
/** Max height before the value list scrolls (CSS px). */
const VALIDATION_PANEL_MAX_H = 200;

let nextViewerProjectionId = 1;

/** How {@link XlsxViewer} presents hidden sheets (`<sheet state>`, §18.2.19). */
export type HiddenSheetMode = 'show' | 'skip' | 'dim';

/** Marker attribute on the single injected viewer stylesheet, so the module-
 *  level injector is idempotent and destroy() can leave it in place. */
const VIEWER_STYLE_ATTR = 'data-xlsx-viewer-styles';

/** Class-constant CSS shared by every XlsxViewer: it styles pseudo-elements
 *  (scrollbar, slider track/thumb) that inline `element.style` cannot reach, so
 *  it must live in a stylesheet rather than on the elements. */
const VIEWER_STYLE_CSS =
  `.xlsx-tab-strip::-webkit-scrollbar{display:none}` +
  // The viewport remains focusable so copy shortcuts belong to the active
  // Viewer. Focus stays quiet by default; consumers can opt into a keyboard
  // focus ring without conflating it with the selected-cell border.
  `[data-xlsx-viewport-input]:focus{outline:none}` +
  `[data-xlsx-viewport-input]:focus-visible{outline:2px solid var(--ooxml-xlsx-focus-ring,transparent);outline-offset:-2px}` +
  `.xlsx-tab-nav{background:transparent;transition:background 0.1s;}` +
  `.xlsx-tab-nav:hover{background:color-mix(in srgb,var(--ooxml-xlsx-chrome-text,#444) 8%,transparent);}` +
  // Excel-status-bar zoom slider: a thin uniform gray track (no colored
  // fill on either side of the thumb) with a small round gray handle.
  `.xlsx-zoom-slider{-webkit-appearance:none;appearance:none;background:transparent;height:15px;margin:0;}` +
  `.xlsx-zoom-slider::-webkit-slider-runnable-track{height:4px;background:var(--ooxml-xlsx-chrome-border,#c4c4c4);border-radius:2px;}` +
  `.xlsx-zoom-slider::-webkit-slider-thumb{-webkit-appearance:none;appearance:none;width:12px;height:12px;margin-top:-4px;border-radius:50%;background:var(--ooxml-xlsx-chrome-text-muted,#808080);cursor:pointer;}` +
  `.xlsx-zoom-slider:hover::-webkit-slider-thumb{background:var(--ooxml-xlsx-chrome-text,#5f5f5f);}` +
  `.xlsx-zoom-slider::-moz-range-track{height:4px;background:var(--ooxml-xlsx-chrome-border,#c4c4c4);border-radius:2px;}` +
  `.xlsx-zoom-slider::-moz-range-thumb{width:12px;height:12px;border:none;border-radius:50%;background:var(--ooxml-xlsx-chrome-text-muted,#808080);cursor:pointer;}`;

/**
 * Inject the shared viewer stylesheet into one owning document exactly once,
 * keyed by the {@link VIEWER_STYLE_ATTR} marker. Earlier this ran
 * per-instance, so every mount/unmount cycle leaked another `<style>` into the
 * head (unbounded growth). It is deliberately NEVER removed on destroy: the CSS
 * is a class constant that any still-live viewer may depend on, and a single
 * leftover `<style>` after the last teardown is harmless (a fixed, bounded cost,
 * not a per-instance leak).
 */
function ensureViewerStyleInjected(ownerDocument: Document): void {
  if (!ownerDocument.head) return;
  if (ownerDocument.head.querySelector(`style[${VIEWER_STYLE_ATTR}]`)) return;
  const style = ownerDocument.createElement('style');
  style.setAttribute(VIEWER_STYLE_ATTR, '');
  style.textContent = VIEWER_STYLE_CSS;
  ownerDocument.head.appendChild(style);
}

export interface XlsxSheetViewerOptions extends LoadOptions {
  /** Adaptive decoded-raster memory policy for visible worksheet paints. */
  imageResources?: import('@silurus/ooxml-core').ImageResourceOptions;
  /** Scale factor for cell/header dimensions (default 1). 0.5 = half size. */
  cellScale?: number;
  /**
   * Enable drag-to-resize of column widths / row heights by dragging header
   * borders. Resizing only changes the on-screen view — it never modifies the
   * loaded file. Default: true.
   */
  resizable?: boolean;
  /**
   * Show native horizontal and vertical scrollbars for the worksheet viewport.
   * Default: true. Wheel/trackpad panning remains available when explicitly
   * disabled.
   */
  showScrollbars?: boolean;
  /**
   * Minimum number of worksheet rows included in the grid and fit calculations.
   * Content, drawings, and frozen panes can extend it further. Default: 50.
   */
  minRows?: number;
  /**
   * Minimum number of worksheet columns included in the grid and fit
   * calculations. Content, drawings, and frozen panes can extend it further.
   * Default: 26.
   */
  minCols?: number;
  /** Additional blank rows after the resolved grid extent. Default: 30. */
  marginRows?: number;
  /** Additional blank columns after the resolved grid extent. Default: 10. */
  marginCols?: number;
  /** Lower/upper bounds for the zoom slider as scale factors. Default 0.1–4
   *  (10%–400%, matching Excel's zoom range). Also the clamp range for the IX9
   *  {@link ZoomableViewer} zoom contract ({@link XlsxViewer.setScale} etc.). */
  zoomMin?: number;
  zoomMax?: number;
  /**
   * IX9 — fires whenever the zoom factor actually changes (`1` = 100%), whatever
   * the source: {@link XlsxViewer.setScale}, {@link XlsxViewer.zoomIn} /
   * {@link XlsxViewer.zoomOut}, {@link XlsxViewer.fitWidth} /
   * {@link XlsxViewer.fitPage}, the built-in zoom slider, the +/- buttons, or a
   * Ctrl/⌘+wheel gesture. Named `onScaleChange` to match the docx/pptx viewers so
   * all five share one notification shape. Not fired when a call resolves to the
   * same (clamped/snapped) scale.
   */
  onScaleChange?: (scale: number) => void;
  onReady?: (sheetNames: string[]) => void;
  /**
   * Called when the active sheet changes, with the new sheet's zero-based
   * `index` and the `total` number of sheets in the workbook. This mirrors the
   * docx `onPageChange` and pptx `onSlideChange` contracts so all three viewers
   * share one callback shape. To get the sheet *name*, look it up by index from
   * `viewer.sheetNames[index]` (or the `sheetNames` array delivered to
   * `onReady`).
   */
  onSheetChange?: (index: number, total: number) => void;
  /**
   * Receives asynchronous Viewer-managed failures that cannot be observed by
   * awaiting the method that started them. Failures from `load()`, including
   * its initial render, always reject that Promise and are not also delivered
   * here. Later event-driven render failures invoke this callback, or fall back
   * to `console.error` when omitted.
   *
   * Stable cases can be narrowed with `OoxmlError`,
   * `OoxmlResourceLimitError`, or `OoxmlDecodedImageLimitError` re-exported by
   * this package. Other failures remain `Error` values; do not parse message
   * text as an API. A `code` of `parser-crashed` identifies a recognized WASM
   * trap, not a reliably classified OOM.
   */
  onError?: (err: Error) => void;
  /** Called with the canonical selection state whenever it actually changes. */
  onSelectionStateChange?: (selection: XlsxSelectionState | null) => void;
  /**
   * Called with a bounded, detached read-only context after selection changes.
   * Rapid changes are coalesced to one notification per animation frame. Use
   * `onSelectionStateChange` instead when canonical UI geometry is required.
   */
  onSelectionContextChange?: (context: XlsxSelectionContext | null) => void;
  /**
   * Called synchronously for a browser `contextmenu` event. The original event
   * can suppress the native menu; `getContext()` resolves the range or element
   * context established at the event target.
   */
  onContextMenu?: (event: ViewerContextMenuEvent<XlsxSelectionContext>) => void;
  /**
   * Enable read-only selection of rendered charts, pictures, and shapes. The
   * selected object exposes element context and receives a non-editable outline.
   * Default false; hit-testing runs only for pointer clicks when enabled.
   */
  enableElementSelection?: boolean;
  /**
   * IX1 (design decision — NOT user-confirmed, integrator may veto). Fires when a
   * cell carrying a hyperlink (ECMA-376 §18.3.1.47) is clicked. Default when
   * omitted: external → {@link openExternalHyperlink} (new tab, sanitised,
   * noopener); internal (`location`) → navigate to the referenced sheet/cell
   * when resolvable. When supplied, this callback fully owns the behaviour and
   * receives the raw {@link HyperlinkTarget} verbatim (URL sanitisation is the
   * default handler's job, so a blocked scheme still reaches a custom callback).
   */
  onHyperlinkClick?: (target: HyperlinkTarget) => void;
  /** IX1 — master switch for hyperlink interactivity. Default `true`. When
   *  `false`, the cell hit-test reports no hyperlink under any cell, so hyperlink
   *  interactivity is disabled entirely: no pointer cursor over a link, no default
   *  navigation (external new-tab / internal sheet jump), and `onHyperlinkClick`
   *  is never called. Hyperlinked cells still render exactly as authored but are
   *  inert. */
  enableHyperlinks?: boolean;
  /**
   * Color of the cell-selection highlight. A single CSS color drives both the
   * selection rectangle's border (drawn in this color) and its fill (the same
   * color made translucent — see {@link selectionOverlayStyle}), so callers pick
   * one accent color instead of a separate border + background. Any CSS color
   * string works (`#1a73e8`, `rgb(...)`, `tomato`, …). Default `#1a73e8`
   * (Google blue), matching the historical look. Can also be changed at runtime
   * via {@link XlsxViewer.setSelectionColor}.
   */
  selectionColor?: string;
  /** CSS backgrounds for ordinary and active in-document search matches. */
  findHighlightColors?: FindHighlightColors;
  /**
   * Show authored cell notes and threaded comments. Pass options to configure
   * resolved-thread visibility. Default true.
   */
  comments?: boolean | XlsxCommentsOptions;
  /**
   * `'main'` (default): parse in a worker, render on the main thread. `'worker'`:
   * parse AND render entirely inside the worker and paint the returned
   * ImageBitmap onto the viewer's canvas, so document rendering never blocks the
   * UI thread. All interaction (scroll, sheet tabs, frozen panes, zoom, cell
   * selection) is unchanged. Requires `Worker` + `OffscreenCanvas`. Built-in
   * math and chart renderers are reconstructed inside the worker from their
   * serializable stable identities.
   */
  /**
   * How hidden / veryHidden sheets (`<sheet state>`, ECMA-376 §18.2.19) are
   * presented:
   * - `'show'` (default): every sheet gets a tab — current behavior.
   * - `'skip'`: hidden/veryHidden sheets get no tab and are jumped over by
   *   `nextSheet`/`prevSheet` and initial load; absolute indices are unchanged,
   *   and an explicit `goToSheet(i)` to a hidden sheet is still honored.
   * - `'dim'`: hidden/veryHidden tabs are shown greyed but stay selectable.
   *
   * Named to match the {@link XlsxViewer.hiddenSheetMode} getter and
   * {@link XlsxViewer.setHiddenSheetMode} setter. Mirrors pptx `hiddenSlideMode`.
   */
  hiddenSheetMode?: HiddenSheetMode;
  /** Called after viewport movement with logical CSS-pixel offsets. */
  onViewportChange?: (offset: XlsxViewportOffset) => void;
}

export interface XlsxViewerOptions extends XlsxSheetViewerOptions {
  /** Show the Excel-style zoom slider at the right end of the sheet-tab bar.
   *  Default `true`. Set `false` to hide it (e.g. when the host supplies its
   *  own zoom control). */
  showZoomSlider?: boolean;
}

type InternalXlsxViewerOptions = (XlsxViewerOptions | XlsxSheetViewerOptions) & {
  [borrowedWorkbookOption]?: XlsxWorkbook;
};

interface GridExtentOptions {
  readonly minRows: number;
  readonly minCols: number;
  readonly marginRows: number;
  readonly marginCols: number;
}

function gridExtentOption(value: number | undefined, fallback: number, name: string): number {
  const resolved = value ?? fallback;
  if (!Number.isSafeInteger(resolved) || resolved < 0) {
    throw new TypeError(`${name} must be a non-negative safe integer`);
  }
  return resolved;
}

function resolveGridExtentOptions(options: XlsxSheetViewerOptions): GridExtentOptions {
  return {
    minRows: gridExtentOption(
      options.minRows,
      DEFAULT_WORKSHEET_CONTENT_MINIMUMS.minRows,
      'minRows',
    ),
    minCols: gridExtentOption(
      options.minCols,
      DEFAULT_WORKSHEET_CONTENT_MINIMUMS.minCols,
      'minCols',
    ),
    marginRows: gridExtentOption(options.marginRows, 30, 'marginRows'),
    marginCols: gridExtentOption(options.marginCols, 10, 'marginCols'),
  };
}

export interface XlsxViewportOffset {
  /** Horizontal CSS-pixel offset from the logical start edge (column A side). */
  readonly x: number;
  /** Vertical CSS-pixel offset from the top of the sheet. */
  readonly y: number;
}

/** Cell bounds in CSS pixels relative to the worksheet viewport's top-left.
 * Values may extend outside the visible viewport for an off-screen cell. */
export interface XlsxCellViewportRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface XlsxScrollToCellOptions {
  readonly align?: 'nearest' | 'start' | 'center' | 'end';
}

export type XlsxCopyResult =
  | Readonly<{ status: 'copied'; cellCount: number; utf16CodeUnits: number }>
  | Readonly<{ status: 'empty-selection' }>
  | Readonly<{ status: 'unsupported-multiple-areas' }>
  | Readonly<{ status: 'too-large'; limit: 'cells' | 'text' }>
  | Readonly<{ status: 'clipboard-unavailable' }>
  | Readonly<{ status: 'clipboard-denied' }>;

export { resizeHitIndex } from './internal/viewer/selection-input.js';

export { selectionOverlayStyle } from './internal/viewer/selection-overlay.js';
export { findHighlightOverlayStyle } from './internal/viewer/find-adapter.js';

type XlsxViewerMount =
  | { readonly kind: 'composite' }
  | {
      readonly kind: 'sheet';
      readonly canvas: HTMLCanvasElement;
      /** Resolved before the caller-owned canvas is reparented. */
      readonly mode: CanvasViewerRenderMode;
    };

class XlsxViewerEngine implements ZoomableViewer {
  private readonly container: HTMLElement;
  /** DOM realm of the mount target. Sheet canvases may belong to a same-origin
   * popup rather than the Window that created the viewer instance. */
  private readonly hostDocument: Document;
  private readonly hostWindow: Window & typeof globalThis;
  private readonly acquisition = new SheetAcquisition();
  private readonly viewport: ViewportState;
  private readonly renderDispatcher: SheetRenderDispatcher;
  /** The single subtree root the constructor appended to the caller's
   *  container. destroy() removes it to return the container to its original
   *  (empty) state. */
  private wrapper!: HTMLDivElement;
  private canvas: HTMLCanvasElement;
  /** Region holding the outline gutters (top/left) and the inset {@link canvasArea}.
   *  When the active sheet has no outlining the gutters collapse to 0 px and this
   *  is a transparent pass-through, so an outline-free sheet lays out identically. */
  private gridRegion!: HTMLDivElement;
  /** Row/column grouping gutters (XL4) beside the grid. */
  private readonly outlineGutter: OutlineGutter;
  /** View-only outline/resize edits, replayed onto every sheet projection. */
  private readonly viewEdits = new SheetViewEdits();
  private readonly projectionId = nextViewerProjectionId++;
  private canvasArea: HTMLDivElement;
  private scrollHost: HTMLDivElement;
  private spacer: HTMLDivElement;
  private readonly surface: CanvasSurface;
  private readonly overlayHost: SheetOverlayHost;
  /** Composite-viewer footer chrome; `null` for sheet mounts, which create no
   *  footer DOM. */
  private readonly sheetTabs: SheetTabBar | null = null;
  private readonly zoomControl: ZoomControl | null = null;
  private currentSheet = 0;
  /** Atomically commits an asynchronously acquired worksheet with its index.
   * Incremented by every navigation and teardown so late acquisitions are no-ops. */
  private sheetRequestGeneration = 0;
  private fontBindingGeneration = 0;
  private fontBinding: Readonly<{ workbook: XlsxWorkbook; release: () => void }> | null = null;
  private _hiddenSheetMode: HiddenSheetMode;
  /** During navigation the outgoing graph stays live for interaction while
   * its lease is released. It can briefly coexist with the incoming graph,
   * so viewer memory can peak at two worksheet models until the swap. */
  private currentWorksheet: Worksheet | null = null;
  private previewCompletion: Promise<Worksheet> | null = null;
  private firstPreviewRender = false;
  /** Counts frames that actually reached the canvas. Completing a pull can
   * supersede the first render before it commits, even when the preview flag
   * has already been cleared by the completion callback. */
  private committedFrameCount = 0;
  private previewPreparedViewport: { width: number; height: number; scale: number } | null = null;
  private previewFallbackReason: ViewportPreviewBlocker | null = null;
  private releaseCurrentWorksheet: (() => void) | null = null;
  /** Authored comments for the selected sheet. Presentation filtering must not
   * erase the application-owned data and selection-context contracts. */
  private currentSourceComments: readonly XlsxComment[] = [];
  /** Latest application-owned comment-list navigation. `scrollToCell()` awaits
   * a render, so an older click must not restore its selection after a newer
   * click or after the current sheet has changed. */
  private commentNavigationGeneration = 0;
  private sourceCommentMap = new Map<string, XlsxComment>();
  /** Viewer-owned projections of workbook-cached worksheets. Only view-mutable
   * size/outline state is copied; immutable cell/content graphs stay shared. */
  private sheetViews = new Map<number, Worksheet>();
  /**
   * #1713 prepared-initial sizes of tagged `twoCellAnchor editAs="oneCell"`
   * pictures/groups by sheet index, valid only for {@link initialAnchorWorkbook}.
   * "Initial" is library policy: the rect resolved at 96 dpi and the viewer's
   * scale on its first prepared projection of the sheet (host fonts/MDW
   * bound, automatic row heights measured, no view edit replayed). Grids round
   * each band's scaled pixels, so that first prepared scale influences the
   * reference; capture divides by it and stores EMU, and later zoom only
   * rescales the stored EMU, never recapturing. Band edits then
   * move the anchor without resizing it (ECMA-376 Part 1 §20.5.2.33/§20.5.3.2).
   * Entries are compact primitive references, never a Worksheet, geometry,
   * cells or anchor objects, so an inactive sheet cannot pin its evicted
   * model. Captured once per workbook/sheet, never replaced nor recaptured
   * from an edited projection; `null` records "nothing eligible". Absolute
   * position freezing for editAs is outside #1713 and not implemented.
   */
  private initialAnchorWorkbook: XlsxWorkbook | null = null;
  private readonly initialAnchorReferences = new Map<number, InitialAnchorSizeReference | null>();
  /** Whether the displayed projection's automatic row heights were measured
   * on this thread, so the worker must not derive them again. */
  private currentAutoRowHeightsPrepared = false;
  private opts: XlsxViewerOptions;
  private readonly _mountKind: XlsxViewerMount['kind'];
  /** Whether this mount delegates viewport movement to a native scroll host. */
  private readonly _nativeScrollbars: boolean;
  /** 'main' renders on this thread; 'worker' paints worker-produced bitmaps. */
  private readonly _mode: 'main' | 'worker';
  private _borrowed = false;
  /** Workbook for which viewer-local state has been initialized. A borrowed
   * sheet mount defers this work until the caller's first goToSheet(), so it
   * never materializes an unrelated first sheet as a constructor side effect. */
  private preparedWorkbook: XlsxWorkbook | null = null;
  /** Set by {@link destroy} (first line). Guards {@link _reportRenderError} so a
   *  render rejection that lands AFTER teardown is swallowed rather than surfaced
   *  to an `onError` / `console.error` on a dead viewer — parity with the scroll
   *  viewers' `_destroyed` flag. */
  private _destroyed = false;
  private resizeObserver: ResizeObserver | null = null;
  /** Inherited chrome CSS variables for Canvas-painted headers and gutters. */
  private chromeTheme!: ChromeTheme;
  private get chromeColors(): XlsxChromeColors {
    return this.chromeTheme.colors;
  }
  /** Last offset delivered to onViewportChange. Keeping this in the shared
   *  engine prevents a programmatic scroll followed by the browser's native
   *  scroll event from producing duplicate notifications. */
  private _lastViewportNotification: XlsxViewportOffset | null = null;
  private get activeCell(): CellAddress | null {
    return this.selectionController.active;
  }

  private get selectionMode(): SheetSelectionMode {
    return this.selectionController.mode;
  }

  /** Gesture-only pointer anchor for the NEXT `setScale`, in canvasArea-viewport
   *  px (`{ x, y }` from the wheel event, relative to the grid's top-left). Set by
   *  the Ctrl/⌘+wheel handler right before it calls `setScale` so the zoom pivots
   *  on the cursor ("zoom toward the pointer") in BOTH axes, past the fixed
   *  header + frozen-pane lead-in; consumed and cleared by `setScale`. `null` for
   *  every non-gesture source (the public `setScale`, the +/- steppers, the zoom
   *  slider, `fitWidth`/`fitPage`), which keep the historical START-anchored
   *  (top-left) preservation so their behaviour is unchanged. */
  private _pendingZoomAnchor: { x: number; y: number } | null = null;

  // Selection state
  private readonly selectionController = new SelectionController();
  /** onSelectionStateChange / onSelectionContextChange delivery. */
  private readonly notifier: SelectionNotifier;
  /** Bounded range-context extraction for getSelectionContext(). */
  private readonly contextReader = new SelectionContextReader();
  private readonly gridExtent: GridExtentOptions;
  private elementContext: XlsxElementContext | null = null;
  /** Selection / object-context overlay painter. */
  private readonly selectionPaint: SelectionOverlay;
  /** IX2 — whole-workbook find and its highlight overlay. */
  private readonly finder: FindAdapter;
  /** Bounded TSV copy of the selected area. */
  private readonly copier = new CopyController({
    worksheet: () => this.currentWorksheet,
    selection: () => this.selectionState,
    workbook: () => this.wb,
    clipboard: () => this.hostWindow.navigator.clipboard,
  });
  /** Viewport pointer/wheel/focus/keyboard input: selection, auto-scroll,
   *  drag-to-resize and the activation gestures. */
  private selectionInput!: SelectionInput;

  /** Excel-style hover note for the displayed sheet's comments. */
  private readonly comments: CommentPopup;
  /** IX1 — cell hyperlink index, enable gate and click dispatch. */
  private readonly hyperlinks: HyperlinkDispatcher;

  /** List data-validation dropdown arrow and display-only value panel. */
  private readonly validation: ValidationPanel;

  constructor(
    container: HTMLElement,
    opts: XlsxViewerOptions | XlsxSheetViewerOptions = {},
    mount: XlsxViewerMount,
  ) {
    this.container = container;
    this.hostDocument =
      (mount.kind === 'sheet' ? mount.canvas.ownerDocument : container.ownerDocument) ?? document;
    const hostWindow = this.hostDocument.defaultView;
    if (!hostWindow) throw new Error('XlsxViewer requires a document with an active Window');
    this.hostWindow = hostWindow;
    this.opts = opts;
    this.gridExtent = resolveGridExtentOptions(opts);
    this._mountKind = mount.kind;
    this._nativeScrollbars = opts.showScrollbars ?? true;
    const borrowedWorkbook = (opts as InternalXlsxViewerOptions)[borrowedWorkbookOption];
    this._borrowed = borrowedWorkbook !== undefined;
    this._mode = mount.kind === 'sheet'
      ? mount.mode
      : resolveCanvasViewerMode('XlsxViewer', opts.mode, borrowedWorkbook);
    this._hiddenSheetMode = opts.hiddenSheetMode ?? 'show';
    this.viewport = new ViewportState(opts.cellScale ?? 1);

    this.wrapper = this.hostDocument.createElement('div');
    this.wrapper.style.cssText =
      `position:relative;width:100%;height:100%;` +
      `background:${mount.kind === 'composite' ? 'var(--ooxml-xlsx-chrome-surface,#fff)' : 'transparent'};` +
      `box-sizing:border-box;font-family:sans-serif;display:flex;flex-direction:column;`;

    // The grid region fills the space above the tab bar. The outline gutters
    // (XL4) sit at its top / left edges and {@link canvasArea} is inset by the
    // gutter extents. With no outlining both extents are 0, so canvasArea covers
    // the whole region exactly as before (byte-identical layout).
    this.gridRegion = this.hostDocument.createElement('div');
    this.gridRegion.style.cssText = `position:relative;flex:1;min-height:0;overflow:hidden;`;

    this.canvasArea = this.hostDocument.createElement('div');
    this.canvasArea.style.cssText = `position:absolute;inset:0;overflow:hidden;`;

    this.canvas = mount.kind === 'sheet' ? mount.canvas : this.hostDocument.createElement('canvas');
    this.canvas.style.cssText = `position:absolute;top:0;left:0;z-index:0;display:block;`;
    this.renderDispatcher = new SheetRenderDispatcher(
      this.canvas,
      this._mode === 'worker',
      this.hostWindow,
    );

    this.scrollHost = this.hostDocument.createElement('div');
    this.scrollHost.setAttribute('data-xlsx-viewport-input', mount.kind);
    this.scrollHost.setAttribute('role', 'region');
    this.scrollHost.setAttribute(
      'aria-label',
      'Spreadsheet viewport. Use Arrow keys to move the selected cell. Press Enter to show its comment.',
    );
    this.scrollHost.tabIndex = 0;
    this.scrollHost.style.cssText =
      `position:absolute;inset:0;` +
      `overflow:${this._nativeScrollbars ? 'auto' : 'clip'};` +
      `z-index:2;background:transparent;` +
      `scrollbar-color:var(--ooxml-xlsx-chrome-scrollbar-color,auto);`;
    this.spacer = this.hostDocument.createElement('div');
    this.spacer.style.cssText = `position:absolute;top:0;left:0;pointer-events:none;`;
    if (this._nativeScrollbars) this.scrollHost.appendChild(this.spacer);
    this.surface = new CanvasSurface(this.canvas, this.canvasArea, this.scrollHost);
    this.overlayHost = new SheetOverlayHost(this.canvasArea, this.canvas, this.scrollHost, {
      commentMaxWidth: COMMENT_POPUP_MAX_W,
      commentMaxHeight: COMMENT_POPUP_MAX_H,
      validationMaxWidth: VALIDATION_PANEL_MAX_W,
      validationMaxHeight: VALIDATION_PANEL_MAX_H,
    });
    this.notifier = new SelectionNotifier({
      hostWindow: this.hostWindow,
      isDestroyed: () => this._destroyed,
      selectionState: () => this.selectionState,
      onSelectionStateChange: () => this.opts.onSelectionStateChange,
      onSelectionContextChange: () => this.opts.onSelectionContextChange,
      readContext: (maxTextCharacters) => this.getSelectionContext({ maxTextCharacters }),
      emitSelectionChange: () => this.emitSelectionChange(),
      scheduleSelectionContextNotification: () => this.scheduleSelectionContextNotification(),
    });
    this.selectionPaint = new SelectionOverlay({
      ownerDocument: this.hostDocument,
      canvasArea: this.canvasArea,
      overlayHost: this.overlayHost,
      worksheet: () => this.currentWorksheet,
      currentSheet: () => this.currentSheet,
      selectionState: () => this.selectionState,
      elementContext: () => this.elementContext,
      elementContextViewport: () => this.elementContextViewport(),
      selectionColor: () => this.opts.selectionColor,
      scale: () => this.viewport.scale,
      isRtl: () => this.isRtl,
      cellRect: (row, col) => this._cellRect(row, col),
      screenX: (x, w) => this.screenX(x, w),
      drawValidationDropdown: () => this.validation.drawDropdown(),
    });
    this.comments = new CommentPopup({
      ownerDocument: this.hostDocument,
      canvasArea: this.canvasArea,
      overlayHost: this.overlayHost,
      currentSheet: () => this.currentSheet,
      isRtl: () => this.isRtl,
      isDestroyed: () => this._destroyed,
      cellRect: (row, col) => this._cellRect(row, col),
      screenX: (x, w) => this.screenX(x, w),
      reportError: (error) => this._reportRenderError(error),
    });
    this.hyperlinks = new HyperlinkDispatcher({
      hostWindow: this.hostWindow,
      enabled: () => this.opts.enableHyperlinks !== false,
      onHyperlinkClick: () => this.opts.onHyperlinkClick,
      currentSheet: () => this.currentSheet,
      sheetNames: () => this.sheetNames,
      definedNames: () => this.currentWorksheet?.definedNames ?? [],
      goToSheet: (index) => this.goToSheet(index),
      scrollToCell: (ref) => this.scrollToCell(ref),
      reportError: (error) => this._reportRenderError(error),
    });
    this.validation = new ValidationPanel({
      ownerDocument: this.hostDocument,
      canvasArea: this.canvasArea,
      surface: this.surface,
      overlayHost: this.overlayHost,
      worksheet: () => this.currentWorksheet,
      workbook: () => this.wb,
      currentSheet: () => this.currentSheet,
      activeCell: () => this.activeCell,
      selectionMode: () => this.selectionMode,
      scale: () => this.viewport.scale,
      isRtl: () => this.isRtl,
      isDestroyed: () => this._destroyed,
      cellRect: (row, col) => this._cellRect(row, col),
      screenX: (x, w) => this.screenX(x, w),
    });
    this.outlineGutter = new OutlineGutter({
      gridRegion: this.gridRegion,
      canvasArea: this.canvasArea,
      surface: this.surface,
      worksheet: () => this.currentWorksheet,
      scale: () => this.viewport.scale,
      chromeColors: () => this.chromeColors,
      cellRect: (row, col) => this._cellRect(row, col),
      screenX: (x, w) => this.screenX(x, w),
      setBandHidden: (axis, index, hidden) => this.setBandHidden(axis, index, hidden),
      setBandCollapsed: (axis, index, collapsed) => this.setBandCollapsed(axis, index, collapsed),
      afterOutlineMutation: (ws, anchor) => this.afterOutlineMutation(ws, anchor),
    });
    // Inject the shared viewer stylesheet once per module (idempotent). Both
    // mounts use it; the composite footer also hides its tab-strip scrollbar.
    ensureViewerStyleInjected(this.hostDocument);

    if (mount.kind === 'composite') {
      this.sheetTabs = new SheetTabBar(this.hostDocument, {
        hiddenSheetMode: () => this._hiddenSheetMode,
        isHidden: (index) => Boolean(this.wb?.isHidden(index)),
        selectSheet: (index) => {
          void this.goToSheet(index).catch((error) => this._reportRenderError(error));
        },
      });
      if (this.opts.showZoomSlider !== false) {
        this.zoomControl = new ZoomControl(this.hostDocument, {
          setScale: (scale) => this.setScale(scale),
          zoomIn: () => this.zoomIn(),
          zoomOut: () => this.zoomOut(),
        }, this.viewport.scale, this.opts.zoomMin ?? 0.1, this.opts.zoomMax ?? 4);
        this.sheetTabs.append(this.zoomControl.element);
      }
    }

    // canvasArea only — the gutter canvases are attached lazily by
    // layoutGutters when (and only when) the shown sheet actually has an
    // outline, and detached again otherwise. Keeping them OUT of the DOM for
    // outline-free sheets preserves exact element parity with the pre-outline
    // viewer: consumers that count or index `<canvas>` elements (the layouts
    // smoke does `page.locator('canvas').count()`, which includes
    // `display:none` nodes) must see no difference.
    this.gridRegion.appendChild(this.canvasArea);
    this.wrapper.appendChild(this.gridRegion);
    if (this.sheetTabs) this.wrapper.appendChild(this.sheetTabs.tabBar);
    container.appendChild(this.wrapper);
    this.chromeTheme = new ChromeTheme({
      hostWindow: this.hostWindow,
      container: this.container,
      wrapper: this.wrapper,
      isDestroyed: () => this._destroyed,
      onChange: () => {
        this.renderGutters();
        this.scheduleRender();
      },
    });
    this.chromeTheme.install();

    // Gutter click handling (XL4): +/- toggles and the numbered level banks
    // (each in its own gutter's header strip; the corner is inert background).
    // Registered once; no-op when a sheet has no gutter (extents 0 ⇒ hidden).
    this.outlineGutter.installListeners();

    if (this._nativeScrollbars) this.surface.on('scroll', () => {
      // Any scroll cancels a deferred tap: the press that started it was a
      // scrollbar-thumb drag (overlay scrollbars) or a touch swipe, not a
      // cell click.
      this.selectionInput.cancelDeferredPress();
      // A comment popup is anchored to a cell's on-screen rect, which moves
      // under the cursor while scrolling — hide it (Excel does the same).
      this.hideCommentPopup();
      // The validation panel is anchored to the cell too; Excel closes its
      // dropdown on scroll, so do the same.
      this.hideValidationPanel();
      // Track the start-anchored position, but only while the host is laid
      // out: a hidden host reports clientWidth 0 and fires bogus scroll
      // events when the browser clamps scrollLeft, which must not overwrite
      // the last real position.
      if (this.scrollHost.clientWidth > 0) {
        const raw = this.scrollHost.scrollLeft;
        const logicalX = this.isRtl ? this.maxScrollLeft - raw : raw;
        this.viewport.setViewportSize(this.scrollHost.clientWidth, this.scrollHost.clientHeight);
        this.viewport.setOffset(logicalX, this.scrollHost.scrollTop);
      }
      this.emitViewportChange();
      // Coalesce into the next frame: a scroll gesture fires many events per
      // frame, and the previous synchronous redraw ran the full render on each
      // one. The overlay update is cheap DOM geometry (no canvas paint) and must
      // track the scroll immediately, so it stays synchronous.
      this.scheduleRender();
      this.updateSelectionOverlay();
      this.updateFindOverlay();
    });

    // Re-render whenever the canvas area changes size. Re-anchor first: a
    // size change shifts maxScrollLeft, and for RTL sheets the native
    // scrollLeft must be re-derived from the start-anchored position or the
    // view drifts (or, after a hidden mount, stays stranded at the far end).
    const resizeObserver = new this.hostWindow.ResizeObserver(() => {
      const offset = { x: this.viewport.x, y: this.viewport.y };
      this.viewport.setViewportSize(this.scrollHost.clientWidth, this.scrollHost.clientHeight);
      this.setViewportLeft(offset.x);
      this.viewportTop = offset.y;
      this.reanchorHorizontalScroll();
      // Re-place the outline gutter strips for the new region size (XL4). This
      // only rewrites styles (no canvasArea size change) so it can't feed back
      // into the observer.
      this.layoutGutters();
      // Container resizes can burst (a live window/pane drag); coalesce the
      // canvas paint into one frame. The re-anchor, overlay and nav updates are
      // cheap and must reflect the new size at once, so they stay synchronous.
      this.scheduleRender();
      this.updateSelectionOverlay();
      this.updateFindOverlay();
      this.sheetTabs?.updateNavButtons();
    });
    resizeObserver.observe(this.gridRegion);
    this.resizeObserver = resizeObserver;

    this.selectionInput = new SelectionInput({
      surface: this.surface,
      scrollHost: this.scrollHost,
      canvasArea: this.canvasArea,
      hostWindow: this.hostWindow,
      nativeScrollbars: this._nativeScrollbars,
      selection: this.selectionController,
      comments: this.comments,
      hyperlinks: this.hyperlinks,
      validation: this.validation,
      options: () => this.opts,
      worksheet: () => this.currentWorksheet,
      hasWorkbook: () => this.wb !== null,
      scale: () => this.viewport.scale,
      isRtl: () => this.isRtl,
      isDestroyed: () => this._destroyed,
      cellAt: (clientX, clientY) => this.getCellAt(clientX, clientY),
      cellRect: (row, col) => this._cellRect(row, col),
      screenX: (x, w) => this.screenX(x, w),
      scrollLeft: () => this.effectiveScrollLeft,
      scrollTop: () => this.viewportTop,
      setScrollLeft: (value) => this.setViewportLeft(value),
      setScrollTop: (value) => { this.viewportTop = value; },
      scrollCellIntoView: (row, col) => this._scrollCellIntoView(row, col),
      zoomAt: (anchor, scale) => {
        this._pendingZoomAnchor = anchor;
        this.setScale(scale);
      },
      elementContextAt: (clientX, clientY) => this.elementContextAt(clientX, clientY),
      setElementContext: (context) => this.setElementContext(context),
      selectionState: () => this.selectionState,
      setSelection: (ref) => this.setSelection(ref),
      getSelectionContext: () => this.getSelectionContext(),
      copySelection: () => this.copySelection(),
      updateSelectionOverlay: () => this.updateSelectionOverlay(),
      updateFindOverlay: () => this.updateFindOverlay(),
      scheduleRender: () => this.scheduleRender(),
      renderCurrentSheet: () => this.renderCurrentSheet(),
      emitSelectionChange: () => this.emitSelectionChange(),
      emitViewportChange: () => this.emitViewportChange(),
      hideCommentPopup: () => this.hideCommentPopup(),
      hideValidationPanel: () => this.hideValidationPanel(),
      recordSizeOverride: (axis, index, columnCssPx) =>
        this.recordSizeOverride(axis, index, columnCssPx),
      assertResizeBudget: (axis, indices, limit) =>
        this.viewEdits.assertResizeBudget(this.currentSheet, axis, indices, limit),
      beginRowResizePreview: targets => {
        const ws = this.currentWorksheet;
        if (!ws) throw new Error('Cannot resize an unavailable worksheet.');
        return this.viewEdits.beginRowResizePreview(ws, this.currentSheet, targets);
      },
      updateSpacerSize: (ws) => this.updateSpacerSize(ws),
      refitAutoRowsAfterColumnResize: () => this.refitAutoRowsAfterColumnResize(),
      reportError: (error) => this._reportRenderError(error),
    });
    this.selectionInput.install();

    this.finder = new FindAdapter({
      ownerDocument: this.hostDocument,
      overlayHost: this.overlayHost,
      workbook: () => this.wb,
      sheetCount: () => this.sheetCount,
      worksheet: () => this.currentWorksheet,
      currentSheet: () => this.currentSheet,
      scale: () => this.viewport.scale,
      highlightColors: () => this.opts.findHighlightColors,
      cellRect: (row, col) => this._cellRect(row, col),
      screenX: (x, w) => this.screenX(x, w),
      goToSheet: (index) => this.goToSheet(index),
      scrollCellIntoView: (row, col) => this._scrollCellIntoView(row, col),
    });

    if (borrowedWorkbook) {
      this.acquisition.install(borrowedWorkbook, false);
      if (this._mountKind === 'composite') {
        this.activateWorkbook(borrowedWorkbook).catch((error) => this._reportRenderError(error));
      }
    }

  }

  /**
   * Load an XLSX from URL or ArrayBuffer and render the first sheet.
   *
   * Parse, load, and initial-render failures always reject this Promise.
   * `onError` is reserved for later Viewer-managed work that has no directly
   * awaitable method result, so one failure is never delivered twice.
   */
  async [loadXlsxViewerSource](
    source: string | ArrayBuffer,
    sourceOptions: XlsxSheetLoadOptions = {},
  ): Promise<void> {
    this.assertOpen();
    if (this._borrowed) {
      throw new Error(
        `${this._mountKind === 'sheet' ? 'XlsxSheetViewer' : 'XlsxViewer'}.load() is unsupported ` +
          'on a Viewer created by fromWorkbook(); the borrowed workbook is already loaded.',
      );
    }
    // SC20 atomic swap: retain the previous workbook locally and only tear it down
    // AFTER the new one loads successfully. A re-load thus never orphans the old
    // workbook's worker + pinned WASM allocation (the leak this guards), yet a
    // FAILED re-load keeps the current workbook + its rendered sheet intact rather
    // than dropping to an empty viewer. The 2× memory window is bounded to the
    // load itself (the old workbook is freed the moment the new model arrives).
    try {
      const wb = await this.acquisition.replace(() => XlsxWorkbook[loadXlsxSheetSource](source, {
          password: this.opts.password,
          useGoogleFonts: this.opts.useGoogleFonts,
          cjkFallback: this.opts.cjkFallback,
          maxZipEntryBytes: this.opts.maxZipEntryBytes,
          resourceLimits: this.opts.resourceLimits,
          debug: this.opts.debug,
          onResourceMetrics: this.opts.onResourceMetrics,
          workerTimeoutMs: this.opts.workerTimeoutMs,
          wasmUrl: this.opts.wasmUrl,
          math: this.opts.math,
          threeD: this.opts.threeD,
          regionMap: this.opts.regionMap,
          chartEx: this.opts.chartEx,
          tiff: this.opts.tiff,
          mode: this._mode,
          ...(this.opts.modelSources === undefined ? undefined : { modelSources: this.opts.modelSources }),
        }, sourceOptions), () => {
          // Claim every async-operation generation before closing the old
          // workbook. Rejections caused by its worker termination are stale
          // completion, not errors belonging to the new workbook.
          this.sheetRequestGeneration++;
          this.renderDispatcher.begin();
          this.finder.invalidate();
          this.hideValidationPanel();
          this.releaseHostFonts();
        });
      if (!wb) return;
      if (this._destroyed) throw this.destroyedError();
      await this.activateWorkbook(wb);
    } catch (err) {
      if (this._destroyed) throw this.destroyedError();
      throw err instanceof Error ? err : new Error(String(err));
    }
  }

  /** Bind the current acquisition to its independent viewer state. Parsing,
   *  worksheet materialization, archive access, and caches remain workbook-owned. */
  private async activateWorkbook(workbook: XlsxWorkbook, sheetIndex?: number): Promise<void> {
    if (!this.prepareWorkbook(workbook)) return;
    await this.showSheet(sheetIndex ?? this._initialSheet());
  }

  private async ensureHostFonts(workbook: XlsxWorkbook): Promise<boolean> {
    if (this.fontBinding?.workbook === workbook) return true;
    const retain = workbook[retainXlsxViewerFonts];
    // Structural test doubles and pre-feature adapters have no font hook. A
    // real XlsxWorkbook always does; absence means there is nothing to retain.
    if (typeof retain !== 'function') return true;
    const generation = ++this.fontBindingGeneration;
    const release = await retain.call(workbook, this.hostDocument);
    if (
      this._destroyed ||
      generation !== this.fontBindingGeneration ||
      this.wb !== workbook
    ) {
      release();
      return false;
    }
    this.fontBinding?.release();
    this.fontBinding = { workbook, release };
    return true;
  }

  private releaseHostFonts(): void {
    this.fontBindingGeneration++;
    this.fontBinding?.release();
    this.fontBinding = null;
  }

  /** Initialize the viewer-local projection state without choosing a sheet.
   * This split lets a borrowed sheet viewer make goToSheet(index) its first
   * worksheet materialization, while the composite viewer can still open its
   * normal initial sheet automatically. */
  private prepareWorkbook(workbook: XlsxWorkbook): boolean {
    if (this._destroyed || this.wb !== workbook) return false;
    if (this.preparedWorkbook === workbook) return true;
    this.finder.invalidate();
    this.selectionInput.clearSheetGestures();
    this.viewEdits.clear();
    this.releaseInitialAnchorReferences();
    this.sheetViews.clear();
    this.buildTabs();
    this.preparedWorkbook = workbook;
    this.opts.onReady?.(workbook.sheetNames);
    return true;
  }

  /** The loaded workbook, or throws if {@link load} has not completed. */
  private get workbook(): XlsxWorkbook {
    const workbook = this.acquisition.current;
    if (!workbook) throw new Error('Workbook not loaded');
    return workbook;
  }

  private get wb(): XlsxWorkbook | null {
    return this.acquisition.current;
  }

  /** Internal assignment seam retained for focused viewer-mechanics tests. All
   *  ownership still flows through SheetAcquisition. */
  private set wb(workbook: XlsxWorkbook | null) {
    if (workbook) this.acquisition.install(workbook);
    else this.acquisition.destroy();
  }

  private async showSheet(index: number): Promise<void> {
    // End a live compact preview before reading/replaying the per-sheet ledger,
    // including same-sheet reloads. It must not become the new projection's
    // committed state merely because an asynchronous navigation interrupted it.
    this.selectionInput.clearSheetGestures();
    const generation = ++this.sheetRequestGeneration;
    this.previewFallbackReason = null;
    const workbook = this.workbook;
    let worksheet: Worksheet;
    let sourceWorksheet: Worksheet;
    let releaseNewWorksheet: (() => void) | undefined;
    let previewCompletion: Promise<Worksheet> | null = null;
    let previewPreparedViewport: { width: number; height: number; scale: number } | null = null;
    // #1713: reference captured by THIS request; installed only after the
    // generation check below, released if the request fails or is stale.
    let pendingReference: InitialAnchorSizeReference | null | undefined;
    let heightsPrepared = false;
    try {
      if (!await this.ensureHostFonts(workbook)) return;
      if (!this.isCurrentSheetRequest(generation, workbook)) return;
      if (index !== this.currentSheet && this.currentWorksheet) {
        // Permit cache eviction, but keep the displayed worksheet and every
        // interaction map intact until a replacement is ready to commit.
        this.releaseCurrentWorksheet?.();
        this.releaseCurrentWorksheet = null;
      }
      const lease = await acquireXlsxWorksheetPreview(workbook, index);
      sourceWorksheet = lease.worksheet;
      releaseNewWorksheet = lease.release;
      // A superseded request returns right after each await below, before it
      // reads or mutates a shared projection (sheetViews may already hold a
      // newer workbook's displayed view), releasing its lease exactly once.
      const abandoned = (): boolean => {
        if (this.isCurrentSheetRequest(generation, workbook)) return false;
        releaseNewWorksheet?.();
        releaseNewWorksheet = undefined;
        return true;
      };
      if (abandoned()) return;
      let eligiblePreview = lease.partial;
      const cachedView = this.sheetViews.get(index);
      // #1713: the first prepared projection of a sheet with tagged
      // twoCellAnchor editAs="oneCell" anchors defines their initial size.
      // A known reference (this viewer + workbook) is bound to every fresh
      // projection before manual state is replayed or a frame is painted;
      // otherwise this request captures it once from an unedited projection.
      // Ordinary sheets only pay one pass over their anchor arrays.
      const knownReference = this.initialAnchorReferenceFor(workbook, index);
      const anchorCoverage = knownReference === undefined
        ? initialAnchorRowCoverage(sourceWorksheet) : undefined;
      const capturing = anchorCoverage !== undefined;
      if (capturing && ((cachedView !== undefined && !cachedView.parseError) || this.viewEdits.hasViewEdits(index))) {
        // An edited projection's current geometry is not the prepared initial
        // geometry and authored sizes cannot reconstruct it: refuse instead.
        throw new Error(
          'XLSX viewer cannot establish the initial size of editAs="oneCell" anchors ' +
            `on sheet ${index} after view edits.`,
        );
      }
      const prepareView = (model: Worksheet): Worksheet => {
        // A parser placeholder replaces the content graph, and a later valid
        // reacquisition replaces that placeholder. Neither reuses the other.
        const reusableView = model.parseError || cachedView?.parseError ? undefined : cachedView;
        const view = reusableView ?? this.createVisibleSheetView(model);
        if (reusableView || lease.partial) view.rows = model.rows;
        if (!reusableView) {
          // A degraded parser placeholder is a different graph: never bind.
          if (knownReference && !model.parseError) bindInitialAnchorSizes(view, knownReference);
          // While capturing, hasViewEdits() is false, so this replay is empty
          // and the capture below still observes the unedited projection.
          this.viewEdits.restoreSheetViewState(index, view);
        }
        return view;
      };
      const prepareHeights = (view: Worksheet, refresh: boolean): void => {
        if (refresh) invalidateAutoRowHeights(view);
        const prepareRowHeights = workbook[prepareXlsxViewerRowHeights];
        if (typeof prepareRowHeights === 'function') {
          const measureCanvas = this.hostDocument.createElement('canvas');
          const measureCtx = measureCanvas.getContext('2d');
          if (measureCtx) {
            prepareRowHeights.call(workbook, view, measureCtx);
            heightsPrepared = true;
          }
        }
        // A superseded request must not write the current workbook's store.
        if (this.isCurrentSheetRequest(generation, workbook)) {
          this.viewEdits.syncAutomaticRowOverrides(index, view);
        }
      };
      // Capture from the fully prepared (host fonts/MDW, automatic heights)
      // unedited projection at the scale its first frame is painted with; a
      // superseded request or a degraded parser placeholder captures nothing.
      const captureInitial = (view: Worksheet): InitialAnchorSizeReference | null | undefined => {
        if (!capturing || view.parseError || !this.isCurrentSheetRequest(generation, workbook)) {
          return undefined;
        }
        const reference = captureInitialAnchorSizes(
          view, getGridGeometryForWorksheet(view), this.viewport.scale) ?? null;
        if (reference) bindInitialAnchorSizes(view, reference);
        return reference;
      };
      worksheet = prepareView(sourceWorksheet);
      if (lease.partial) {
        const preparedViewport = {
          width: this.canvasArea.clientWidth,
          height: this.canvasArea.clientHeight,
          scale: this.viewport.scale,
        };
        previewPreparedViewport = preparedViewport;
        const visibleRange = () => getGridGeometryForWorksheet(worksheet).visibleRange({
          width: preparedViewport.width,
          height: preparedViewport.height,
          scale: preparedViewport.scale,
          scrollX: 0, scrollY: 0,
          headerWidth: HEADER_W, headerHeight: HEADER_H, buffer: 2,
        });
        let visible = visibleRange();
        let coveringRow = 0;
        // #1713: while capturing, rows through every eligible anchor's marker
        // rows must be prepared as well (bounded; usually inside the viewport).
        const anchorRow = anchorCoverage ?? 0;
        for (;;) {
          const needed = Math.max(
            visible.range.row + visible.range.rows - 1, worksheet.freezeRows ?? 0, anchorRow);
          if (needed > coveringRow) {
            await (lease.waitForRows?.(needed) ?? Promise.resolve());
            if (abandoned()) return;
            coveringRow = needed;
          }
          // Rows that arrived while waiting can change display-derived height
          // and therefore bring additional rows into the first viewport.
          prepareHeights(worksheet, true);
          visible = visibleRange();
          if (Math.max(
            visible.range.row + visible.range.rows - 1, worksheet.freezeRows ?? 0, anchorRow,
          ) <= coveringRow) break;
        }
        // Paint includes the frozen corner, frozen row/column strips, and the
        // scrollable quadrant. Use their union for both row coverage and every
        // dependency check; the renderer can also spill text horizontally from
        // cells outside the visible column band in these rows.
        const painted = {
          row: (worksheet.freezeRows ?? 0) > 0 ? 1 : visible.range.row,
          col: (worksheet.freezeCols ?? 0) > 0 ? 1 : visible.range.col,
          rows: visible.range.row + visible.range.rows - ((worksheet.freezeRows ?? 0) > 0 ? 1 : visible.range.row),
          cols: visible.range.col + visible.range.cols - ((worksheet.freezeCols ?? 0) > 0 ? 1 : visible.range.col),
        };
        // The anchor baseline also needs final automatic heights for rows
        // outside the painted viewport. anchorBaselinePreviewBlocker reports
        // `drawing-dependency` when a statistical/formula conditional format
        // reaches rows not loaded yet (they can change those heights), so the
        // full model is awaited before capture (library policy).
        this.previewFallbackReason = viewportPreviewBlocker(worksheet, painted, coveringRow) ??
          (capturing ? anchorBaselinePreviewBlocker(worksheet, anchorRow, coveringRow) : null);
        if (this.previewFallbackReason) {
          sourceWorksheet = await lease.completion;
          if (abandoned()) return;
          eligiblePreview = false;
          previewPreparedViewport = null;
          worksheet = prepareView(sourceWorksheet);
          prepareHeights(worksheet, true);
        }
        // Also after the completion await: captureInitial re-checks generation.
        pendingReference = captureInitial(worksheet);
      } else {
        prepareHeights(worksheet, false);
        pendingReference = captureInitial(worksheet);
      }
      previewCompletion = eligiblePreview ? lease.completion : null;
    } catch (error) {
      releaseNewWorksheet?.();
      if (pendingReference) releaseInitialAnchorSizeReference(pendingReference);
      if (!this.isCurrentSheetRequest(generation, workbook)) return;
      await this.restoreDisplayedWorksheetLease(workbook, generation);
      throw error;
    }
    if (!this.isCurrentSheetRequest(generation, workbook)) {
      releaseNewWorksheet?.();
      if (pendingReference) releaseInitialAnchorSizeReference(pendingReference);
      return;
    }
    // Another row preview may have begun while acquisition awaited. Reconcile
    // after cancelling it, before the prepared snapshot becomes reachable.
    this.selectionInput.clearSheetGestures();
    this.viewEdits.restoreRowResizeRanges(index, worksheet);
    // #1713: install synchronously after the generation check, before the
    // projection is stored, painted or reachable by manual edit input.
    if (pendingReference !== undefined) {
      this.installInitialAnchorReference(workbook, index, pendingReference);
    }
    this.currentAutoRowHeightsPrepared = heightsPrepared;

    this.releaseCurrentWorksheet?.();
    this.releaseCurrentWorksheet = releaseNewWorksheet ?? null;
    // Viewer projections share the full cell graph. Keeping inactive entries
    // would defeat workbook eviction even after its cache drops the model.
    this.sheetViews.clear();
    this.sheetViews.set(index, worksheet);
    this.currentSheet = index;
    this.currentWorksheet = worksheet;
    this.previewCompletion = previewCompletion;
    this.firstPreviewRender = previewCompletion !== null;
    this.previewPreparedViewport = previewPreparedViewport;
    if (previewCompletion) {
      void previewCompletion.then((completed) => {
        if (!this.isCurrentSheetRequest(generation, workbook) || this.currentWorksheet !== worksheet) return;
        this.previewCompletion = null;
        this.firstPreviewRender = false;
        this.previewPreparedViewport = null;
        // Chart and sparkline references can resolve more completely once all
        // rows exist. Rebind the viewer to the committed model, preserving only
        // viewer-owned size and outline edits made during the pull.
        // A degraded parser placeholder is decided before any binding: it is
        // a different graph, so the known reference is never forced onto it.
        const degraded = Boolean(completed.parseError);
        this.selectionInput.clearSheetGestures();
        const finalized = this.createVisibleSheetView(completed);
        // #1713: a non-degraded terminal projection reuses this request's
        // reference, bound before manual state is replayed; never recaptured.
        const initialReference = degraded ? undefined : this.initialAnchorReferenceFor(workbook, index);
        if (initialReference) bindInitialAnchorSizes(finalized, initialReference);
        this.viewEdits.restoreSheetViewState(index, finalized);
        this.currentWorksheet = finalized;
        this.sheetViews.set(index, finalized);
        if (degraded) {
          // A later row may make the cursor produce the normal degraded-sheet
          // placeholder. Replace the provisional graph before the next frame.
          this.currentSourceComments = [];
          this.sourceCommentMap.clear();
          this.selectionController.reset();
          this.emitSelectionChange();
          this.updateSelectionOverlay();
          this.buildCommentMap(finalized);
          this.buildHyperlinkMap(finalized);
          this.buildOutline(finalized);
          this.layoutGutters();
          this.updateSpacerSize(finalized);
          this.scheduleRender();
          return;
        }
        invalidateSheetRenderCache(worksheet);
        invalidateAutoRowHeights(worksheet);
        const measureCtx = this.hostDocument.createElement('canvas').getContext('2d');
        if (measureCtx) {
          workbook[prepareXlsxViewerRowHeights](finalized, measureCtx);
          this.currentAutoRowHeightsPrepared = true;
        }
        this.viewEdits.syncAutomaticRowOverrides(index, finalized);
        this.currentSourceComments = completed.comments ?? [];
        this.sourceCommentMap = createCommentMap(this.currentSourceComments);
        this.buildCommentMap(finalized);
        this.buildHyperlinkMap(finalized);
        this.buildOutline(finalized);
        this.layoutGutters();
        this.updateSpacerSize(finalized);
        this.scheduleRender();
        this.scheduleSelectionContextNotification();
      }).catch((error: unknown) => {
        if (!this.isCurrentSheetRequest(generation, workbook) || this.currentWorksheet !== worksheet) return;
        this.previewCompletion = null;
        this.firstPreviewRender = false;
        this.previewPreparedViewport = null;
        this.selectionInput.clearSheetGestures();
        this.currentWorksheet = null;
        this.releaseCurrentWorksheet?.();
        this.releaseCurrentWorksheet = null;
        this.renderDispatcher.begin();
        if (this._mode === 'worker') {
          // A bitmaprenderer canvas has no 2D context, and resizing it can
          // retain the last transferred frame. Replace it with an empty bitmap.
          const surface = new OffscreenCanvas(1, 1);
          surface.getContext('2d');
          const blank = surface.transferToImageBitmap();
          this.canvas.getContext('bitmaprenderer')?.transferFromImageBitmap(blank);
          blank.close();
        } else {
          this.canvas.getContext('2d')?.clearRect?.(0, 0, this.canvas.width, this.canvas.height);
        }
        this._reportRenderError(error);
      });
    }
    this.currentSourceComments = sourceWorksheet.comments ?? [];
    if (this.opts.comments !== false && this.currentSourceComments.length > 0) {
      void this.comments.loadUi().catch((error) => this._reportRenderError(error));
    }
    this.sourceCommentMap = createCommentMap(this.currentSourceComments);
    this.setElementContext(null);
    this.selectionInput.clearSheetGestures();
    this.updateFooterDirection();
    this.viewportTop = 0;
    this.selectionController.reset();
    this.emitSelectionChange();
    this.hideCommentPopup();
    this.hideValidationPanel();
    this.updateSelectionOverlay();
    this.updateTabActive(index);
    this.buildCommentMap(this.currentWorksheet);
    this.buildHyperlinkMap(this.currentWorksheet);
    // XL4: build the outline layout for this sheet and size the gutters. Must run
    // before `updateSpacerSize` / render so the inset canvasArea has its final
    // size when the grid geometry is computed.
    this.buildOutline(this.currentWorksheet);
    this.layoutGutters();
    this.updateSpacerSize(this.currentWorksheet);
    // Reset the horizontal scroll origin to the natural START of the sheet.
    // For RTL sheets the start column (col A) lives at the RIGHT, which means
    // the native scrollbar thumb must sit at its right end (max scrollLeft);
    // for LTR sheets the start is scrollLeft=0. updateSpacerSize must run first
    // so scrollWidth reflects the new sheet before we read the max offset.
    this.resetHorizontalScroll();
    const frameBefore = this.committedFrameCount;
    await this.renderCurrentSheet();
    const paintedEarly = this.committedFrameCount > frameBefore;
    if (previewCompletion && !paintedEarly && this.isCurrentSheetRequest(generation, workbook)) {
      await previewCompletion;
      if (!this.isCurrentSheetRequest(generation, workbook)) return;
      await this.renderCurrentSheet();
    }
    if (!this.isCurrentSheetRequest(generation, workbook)) return;
    // Redraw find highlights for the newly shown sheet (the find state survives
    // a sheet switch; only the visible sheet's boxes are drawn).
    this.updateFindOverlay();
    this.emitViewportChange();
    this.opts.onSheetChange?.(index, this.workbook.sheetNames.length);
  }

  private isCurrentSheetRequest(generation: number, workbook: XlsxWorkbook): boolean {
    return !this._destroyed && generation === this.sheetRequestGeneration && this.wb === workbook;
  }

  /** #1713 entry of `index` for the workbook it was captured from: a
   * reference, `null` (captured; nothing eligible) or `undefined` (none). */
  private initialAnchorReferenceFor(
    workbook: XlsxWorkbook,
    index: number,
  ): InitialAnchorSizeReference | null | undefined {
    return this.initialAnchorWorkbook === workbook ? this.initialAnchorReferences.get(index) : undefined;
  }

  private installInitialAnchorReference(
    workbook: XlsxWorkbook,
    index: number,
    reference: InitialAnchorSizeReference | null,
  ): void {
    // References belong to one workbook; adopting another releases the old set.
    if (this.initialAnchorWorkbook !== workbook) {
      this.releaseInitialAnchorReferences();
      this.initialAnchorWorkbook = workbook;
    }
    // At most one capture per sheet: a request captures only when no entry
    // exists and installs only while it is the current generation.
    if (!this.initialAnchorReferences.has(index)) this.initialAnchorReferences.set(index, reference);
  }

  /** Release every reference (bound lookups stop applying), then forget them. */
  private releaseInitialAnchorReferences(): void {
    for (const reference of this.initialAnchorReferences.values()) {
      if (reference) releaseInitialAnchorSizeReference(reference);
    }
    this.initialAnchorReferences.clear();
    this.initialAnchorWorkbook = null;
  }

  private async restoreDisplayedWorksheetLease(workbook: XlsxWorkbook, generation: number): Promise<void> {
    if (!this.currentWorksheet || this.releaseCurrentWorksheet) return;
    const release = await retainXlsxWorksheetReference(workbook, this.currentSheet);
    if (this.isCurrentSheetRequest(generation, workbook) && this.currentWorksheet && !this.releaseCurrentWorksheet) {
      this.releaseCurrentWorksheet = release;
    } else {
      release();
    }
  }

  // ─── Outline gutter (XL4: row/column grouping) ────────────────────────────

  /** Recompute the per-axis outline layout for `ws` and bind the sheet's
   *  view-edit stashes. An outline-free sheet collapses both gutters to 0. */
  private buildOutline(ws: Worksheet): void {
    this.viewEdits.bindSheet(this.currentSheet);
    this.outlineGutter.rebuild(ws);
  }

  /** Place the gutters and inset canvasArea by their extents. */
  private layoutGutters(): void {
    this.outlineGutter.layout();
  }

  /** Repaint the gutters for the current scroll offset (after every frame). */
  private renderGutters(): void {
    this.outlineGutter.render();
  }

  /** Align an outline summary band to the scrollable viewport's start without
   * disturbing the perpendicular axis. */
  private scrollOutlineSummaryToStart(axis: OutlineAxis, summary: number): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const cs = this.viewport.scale;
    const offset = getGridGeometryForWorksheet(ws).scrollOffsetForCell(
      axis === 'row' ? summary : 1,
      axis === 'col' ? summary : 1,
      {
        scale: cs,
        viewportWidth: this.canvasArea.clientWidth,
        viewportHeight: this.canvasArea.clientHeight,
        currentX: this.effectiveScrollLeft,
        currentY: this.viewportTop,
        headerWidth: HEADER_W,
        headerHeight: HEADER_H,
        align: 'start',
      },
    );
    if (axis === 'row') this.viewportTop = offset.y;
    else this.setViewportLeft(offset.x);
  }

  private setBandHidden(axis: OutlineAxis, index: number, hidden: boolean): void {
    const ws = this.currentWorksheet;
    if (ws) this.viewEdits.setBandHidden(ws, this.currentSheet, axis, index, hidden);
  }

  private recordSizeOverride(axis: OutlineAxis, index: number, columnCssPx?: number): void {
    const ws = this.currentWorksheet;
    if (ws) this.viewEdits.recordSizeOverride(ws, this.currentSheet, axis, index, columnCssPx);
  }

  private wireSizeOverrides(): ReturnType<SheetViewEdits['wireSizeOverrides']> {
    return this.viewEdits.wireSizeOverrides(this.currentSheet);
  }

  private setBandCollapsed(axis: OutlineAxis, index: number, collapsed: boolean): void {
    const ws = this.currentWorksheet;
    if (ws) this.viewEdits.setBandCollapsed(ws, this.currentSheet, axis, index, collapsed);
  }

  /** Shared tail of a gutter interaction: invalidate the axis cache, rebuild the
   *  outline (collapsed flags changed), refresh dependent geometry, re-render. */
  private afterOutlineMutation(
    ws: Worksheet,
    anchor?: { axis: OutlineAxis; summary: number },
  ): void {
    GridGeometry.invalidate(ws);
    this.outlineGutter.rebuild(ws);
    this.updateSpacerSize(ws);
    if (anchor) this.scrollOutlineSummaryToStart(anchor.axis, anchor.summary);
    this.updateSelectionOverlay();
    this.updateFindOverlay();
    this.scheduleRender();
    if (anchor) this.emitViewportChange();
  }

  /** True when the current sheet's grid is laid out right-to-left. */
  private get isRtl(): boolean {
    return this.currentWorksheet?.rightToLeft === true;
  }

  /** Mirror the workbook footer for an RTL sheet (composite mounts only). */
  private updateFooterDirection(): void {
    this.sheetTabs?.setDirection(this.isRtl);
  }

  /** Maximum horizontal logical viewport offset (≥ 0). */
  private get maxScrollLeft(): number {
    this.syncNativeViewportExtent();
    return this.viewport.maxX;
  }

  private get maxScrollTop(): number {
    this.syncNativeViewportExtent();
    return this.viewport.maxY;
  }

  private syncNativeViewportExtent(): void {
    if (!this._nativeScrollbars) return;
    this.viewport.setViewportSize(this.scrollHost.clientWidth, this.scrollHost.clientHeight);
    this.viewport.ensureExtent(this.scrollHost.scrollWidth, this.scrollHost.scrollHeight);
  }

  private get viewportTop(): number {
    if (this._nativeScrollbars) {
      this.syncNativeViewportExtent();
      this.viewport.adoptNativeOffset(this.viewport.x, this.scrollHost.scrollTop);
    }
    return this.viewport.y;
  }

  private set viewportTop(value: number) {
    this.viewport.setOffset(this.viewport.x, value);
    if (this._nativeScrollbars) this.scrollHost.scrollTop = this.viewport.y;
  }

  /**
   * The logical horizontal scroll position used to find the start-of-sheet
   * (col A) edge, in *scaled* CSS pixels — the same unit as
   * `scrollHost.scrollLeft`. The renderer always lays the grid out LTR and then
   * mirrors it (ECMA-376 §18.3.1.87), so the viewer must hand it a position
   * where 0 = the START of the sheet (col A) and increasing values reveal later
   * columns.
   *
   * For LTR that is exactly the native `scrollLeft`. For RTL the sheet starts at
   * the RIGHT, so the native scrollbar runs the opposite way: thumb fully right
   * (`scrollLeft = maxScrollLeft`) is the start, thumb left is the far columns.
   * Inverting here makes wheel/trackpad follow the finger and aligns the
   * thumb↔page mapping with Excel, without depending on browser-specific RTL
   * `scrollLeft` sign conventions.
   */
  private get effectiveScrollLeft(): number {
    if (this._nativeScrollbars) {
      this.syncNativeViewportExtent();
      const raw = this.scrollHost.scrollLeft;
      this.viewport.adoptNativeOffset(this.isRtl ? this.maxScrollLeft - raw : raw, this.viewport.y);
    }
    return this.viewport.x;
  }

  private setViewportLeft(value: number): void {
    this.viewport.setOffset(value, this.viewport.y);
    if (this._nativeScrollbars) {
      this.scrollHost.scrollLeft = this.isRtl
        ? Math.max(0, this.maxScrollLeft - this.viewport.x)
        : this.viewport.x;
    }
  }

  /**
   * Map between the logical-LTR x used by all the cell-geometry math and the
   * on-screen (canvasArea CSS-pixel) x, applying the RTL mirror (ECMA-376
   * §18.3.1.87) via the same {@link rtlMirrorX} the renderer uses. For LTR this
   * is the identity. The mirror is an involution, so this one method serves
   * both cell→px (overlay draw, `w` = cell width) and px→cell (pointer
   * hit-testing, `w` = 0 for a point) — guaranteeing the overlay sits exactly
   * where the cell is drawn and a click resolves to that same cell at every
   * scroll offset. `canvasArea.clientWidth` equals the renderer's `canvasW`.
   */
  private screenX(logicalX: number, w: number): number {
    return this.isRtl ? rtlMirrorX(logicalX, w, this.canvasArea.clientWidth) : logicalX;
  }

  /** Park the scrollbar at the sheet's natural start: scrollLeft=0 for LTR,
   *  the right end for RTL (so col A shows first). */
  private resetHorizontalScroll(): void {
    this.viewport.setOffset(0, this.viewport.y);
    if (this._nativeScrollbars) {
      this.scrollHost.scrollLeft = this.isRtl ? this.maxScrollLeft : 0;
    }
  }

  /** Re-derive the native scrollLeft from the tracked start-anchored
   *  position after the scroll host's size changes. Only RTL needs this:
   *  for LTR the native scrollLeft *is* start-anchored and the browser
   *  already clamps it sensibly on resize. */
  private reanchorHorizontalScroll(): void {
    if (!this._nativeScrollbars) return;
    if (!this.isRtl || this.scrollHost.clientWidth === 0) return;
    const want = Math.max(0, this.maxScrollLeft - this.viewport.x);
    if (Math.abs(this.scrollHost.scrollLeft - want) > 1) {
      this.scrollHost.scrollLeft = want;
    }
  }

  /** 0-based index of the currently displayed sheet. */
  get sheetIndex(): number {
    return this.currentSheet;
  }

  /** Total number of sheets in the loaded workbook. */
  get sheetCount(): number {
    return this.wb?.sheetCount ?? 0;
  }

  /**
   * Navigate to a sheet by index, clamped to range. Canonical navigation verb
   * matching {@link PptxViewer.goToSlide} / {@link DocxViewer.goToPage}.
   */
  async goToSheet(index: number): Promise<void> {
    if (this.sheetCount === 0) return;
    const workbook = this.workbook;
    if (!this.prepareWorkbook(workbook)) return;
    await this.showSheet(Math.max(0, Math.min(index, this.sheetCount - 1)));
  }

  async nextSheet(): Promise<void> {
    await this.goToSheet(this._stepSheet(1));
  }

  async prevSheet(): Promise<void> {
    await this.goToSheet(this._stepSheet(-1));
  }

  /** Logical start-anchored viewport offset in CSS pixels at the current scale. */
  getViewportOffset(): XlsxViewportOffset {
    return {
      x: Math.max(0, this.effectiveScrollLeft),
      y: Math.max(0, this.viewportTop),
    };
  }

  private emitViewportChange(): void {
    const callback = this.opts.onViewportChange;
    if (!callback) return;
    const offset = this.getViewportOffset();
    const previous = this._lastViewportNotification;
    if (previous && previous.x === offset.x && previous.y === offset.y) return;
    this._lastViewportNotification = offset;
    callback(offset);
  }

  /** Move the active sheet viewport without exposing browser RTL scroll rules. */
  async setViewportOffset(offset: XlsxViewportOffset): Promise<void> {
    if (!Number.isFinite(offset.x) || !Number.isFinite(offset.y)) {
      throw new TypeError('XLSX viewport offsets must be finite numbers');
    }
    const x = Math.min(this.maxScrollLeft, Math.max(0, offset.x));
    const y = Math.min(this.maxScrollTop, Math.max(0, offset.y));
    this.setViewportLeft(x);
    this.viewportTop = y;
    await this.renderCurrentSheet();
    this.updateSelectionOverlay();
    this.updateFindOverlay();
    this.emitViewportChange();
  }

  /** Re-read the mount's CSS box and repaint the current viewport. */
  async relayout(): Promise<void> {
    this.reanchorHorizontalScroll();
    this.layoutGutters();
    if (this.currentWorksheet) this.updateSpacerSize(this.currentWorksheet);
    await this.renderCurrentSheet();
    this.updateSelectionOverlay();
    this.updateFindOverlay();
  }

  async scrollToCell(
    ref: string,
    options: XlsxScrollToCellOptions = {},
  ): Promise<void> {
    const cell = parseA1(ref);
    if (!cell || !this.currentWorksheet) return;
    if (this.previewCompletion) await this.previewCompletion;
    this._scrollCellIntoView(cell.row, cell.col, options.align ?? 'nearest');
    await this.renderCurrentSheet();
    this.updateSelectionOverlay();
    this.updateFindOverlay();
    this.emitViewportChange();
  }

  /** Next sheet index for sequential nav: skip mode jumps over hidden sheets. */
  private _stepSheet(dir: 1 | -1): number {
    if (this._hiddenSheetMode === 'skip' && this.wb) {
      return nextVisibleIndex(this.currentSheet, dir, (i) => this.wb!.isHidden(i), this.sheetCount);
    }
    return this.currentSheet + dir;
  }

  /** Initial sheet for load() / entering skip mode: land on a visible sheet. */
  private _initialSheet(): number {
    if (this._hiddenSheetMode === 'skip' && this.wb) {
      return resolveVisibleIndex(0, (i) => this.wb!.isHidden(i), this.sheetCount);
    }
    return 0;
  }

  /** Returns the cell at canvas-client coordinates, or null if outside the cell grid. */
  getCellAt(clientX: number, clientY: number): CellAddress | null {
    if (this._destroyed) return null;
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.viewport.scale;

    const rect = this.canvasArea.getBoundingClientRect();
    // Un-mirror the screen x into the logical-LTR layout the geometry below
    // assumes (header on the left). screenX is an involution, so applying it to
    // a screen point recovers the logical point; w = 0 for a point. Done in
    // scaled CSS px (canvasArea space) before converting to logical px.
    const lx = this.screenX(clientX - rect.left, 0);
    const ly = clientY - rect.top;

    const scaledHeaderW = Math.round(HEADER_W * cs);
    const scaledHeaderH = Math.round(HEADER_H * cs);
    if (lx < scaledHeaderW || ly < scaledHeaderH) return null;

    const innerX = lx - scaledHeaderW;
    const innerY = ly - scaledHeaderH;

    return getGridGeometryForWorksheet(ws).cellAt(innerX, innerY, {
      scrollX: this.effectiveScrollLeft,
      scrollY: this.viewportTop,
      scale: cs,
    });
  }

  /** Click-only DrawingML hit test. It walks just the sheet's anchored object
   * arrays and never scans worksheet cells or runs during render/scroll. */
  private elementContextViewport(): XlsxElementHitViewport | null {
    const worksheet = this.currentWorksheet;
    if (!worksheet) return null;
    const width = this.canvasArea.clientWidth;
    const height = this.canvasArea.clientHeight;
    if (width <= 0 || height <= 0) return null;
    const scale = this.viewport.scale;
    const geometry = getGridGeometryForWorksheet(worksheet);
    const visible = geometry.visibleRange({
      width,
      height,
      scale,
      scrollX: this.effectiveScrollLeft,
      scrollY: this.viewportTop,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
      buffer: 2,
    });
    return {
      width,
      height,
      cellScale: scale,
      viewport: visible.range,
      scrollOffsetX: visible.offsetX,
      scrollOffsetY: visible.offsetY,
      freezeRows: worksheet.freezeRows ?? 0,
      freezeCols: worksheet.freezeCols ?? 0,
    };
  }

  private elementContextAt(clientX: number, clientY: number): XlsxElementContext | null {
    if (!this.opts.enableElementSelection || this._destroyed) return null;
    const worksheet = this.currentWorksheet;
    const viewport = this.elementContextViewport();
    if (!worksheet || !viewport) return null;
    const rect = this.canvasArea.getBoundingClientRect();
    return hitTestXlsxElementContext(
      worksheet,
      this.currentSheet,
      { x: clientX - rect.left, y: clientY - rect.top },
      viewport,
    );
  }

  /** Returns the CSS-pixel rect of a cell within canvasArea, or null if not
   *  computable. Mirrors the renderer's per-cell rounding (Math.round(px * cs))
   *  so the selection overlay sits exactly on the canvas's drawn cell borders;
   *  multiplying logical accumulators by `cs` once at the end (the previous
   *  approach) drifted by up to 1 px per cell at non-integer scales.
   */
  private _cellRect(row: number, col: number): { x: number; y: number; w: number; h: number } | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.viewport.scale;
    return getGridGeometryForWorksheet(ws).cellRect(row, col, {
      scale: cs,
      scrollX: this.effectiveScrollLeft,
      scrollY: this.viewportTop,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
    });
  }

  /** Return one cell's viewport-relative CSS-pixel bounds. This is the forward
   * geometry primitive for application-owned comment or annotation overlays. */
  getCellViewportRect(cell: CellAddress | string): XlsxCellViewportRect | null {
    if (this._destroyed) return null;
    const address = typeof cell === 'string' ? parseA1(cell) : cell;
    if (!address || address.row < 1 || address.col < 1) return null;
    const rect = this._cellRect(address.row, address.col);
    return rect
      ? Object.freeze({
          x: this.screenX(rect.x, rect.w),
          y: rect.y,
          width: rect.w,
          height: rect.h,
        })
      : null;
  }

  /** Detached comments for the current sheet, in authored order. */
  getComments(): readonly Readonly<XlsxComment>[] {
    this.assertOpen();
    return structuredClone(this.currentSourceComments);
  }

  /**
   * Reveal and select the cell that owns a comment on an explicit sheet. This deliberately owns
   * no list UI: applications render detached records from `getComments()` and
   * call this navigation primitive from their own rows.
   *
   * Returns `false` when the sheet index is invalid or `cellRef` does not
   * identify a comment on that sheet.
   */
  async goToComment(
    sheetIndex: number,
    cellRef: string,
    options?: XlsxScrollToCellOptions,
  ): Promise<boolean> {
    const target = parseA1(cellRef);
    const workbook = this.wb;
    if (
      !target || !workbook || !Number.isInteger(sheetIndex) ||
      sheetIndex < 0 || sheetIndex >= workbook.sheetCount
    ) {
      return false;
    }
    const generation = ++this.commentNavigationGeneration;
    const comments = sheetIndex === this.currentSheet && this.currentWorksheet !== null
      ? this.currentSourceComments
      : await workbook.getComments(sheetIndex);
    if (this._destroyed) throw this.destroyedError();
    if (generation !== this.commentNavigationGeneration || workbook !== this.wb) return false;
    if (!comments.some((comment) => {
      const cell = parseA1(comment.cellRef);
      return cell?.row === target.row && cell.col === target.col;
    })) return false;

    if (sheetIndex !== this.currentSheet || this.currentWorksheet === null) {
      await this.goToSheet(sheetIndex);
      if (this._destroyed) throw this.destroyedError();
      if (
        generation !== this.commentNavigationGeneration || workbook !== this.wb ||
        sheetIndex !== this.currentSheet
      ) {
        return false;
      }
    }
    const sheetGeneration = this.sheetRequestGeneration;
    const sheet = this.currentSheet;
    const worksheet = this.currentWorksheet;
    await this.scrollToCell(cellRef, options);
    if (this._destroyed) throw this.destroyedError();
    if (
      generation !== this.commentNavigationGeneration ||
      workbook !== this.wb ||
      sheetGeneration !== this.sheetRequestGeneration ||
      sheet !== this.currentSheet ||
      worksheet !== this.currentWorksheet
    ) return false;
    this.setSelection(cellRef);
    return true;
  }

  /** Returns the full selection model, detached from viewer-owned state. */
  get selectionState(): XlsxSelectionState | null {
    return this.selectionController.snapshot();
  }

  /**
   * Set an A1 area (`B2:D5`, `2:4`, `B:D`), a complete canonical state, or
   * `null`. A string describes selection geometry only; its normalized
   * upper-left cell becomes ActiveCell and the Shift-extension anchor.
   */
  setSelection(input: XlsxSelectionInput): void {
    if (this._destroyed) throw new Error('XlsxViewer has been destroyed');
    let next: XlsxSelectionState | null;
    if (typeof input === 'string') {
      next = selectionStateFromReference(input);
      if (!next) throw new SyntaxError(`Invalid XLSX selection reference: ${input}`);
    } else {
      next = input ? normalizeSelectionState(input) : null;
    }
    this.commitSelection(next);
  }

  /**
   * Return a serializable, bounded snapshot of the current selection and the
   * populated cells it covers. Intended for read-only AI/MCP context handoff;
   * it exposes no mutable workbook objects and does not touch the Clipboard API.
   */
  getSelectionContext(options: XlsxSelectionContextOptions = {}): XlsxSelectionContext | null {
    this.assertOpen();
    // This synchronous data API cannot await an unloaded selection. Geometry
    // remains selectable; a fresh context notification follows completion.
    if (this.previewCompletion) return null;
    if (this.elementContext) {
      return limitXlsxElementContext(this.elementContext, options.maxTextCharacters);
    }
    const worksheet = this.currentWorksheet;
    const selection = this.selectionState;
    if (!worksheet || !selection) return null;
    return this.contextReader.read(
      worksheet,
      this.currentSheet,
      selection,
      this.sourceCommentMap,
      this.wb,
      options,
    );
  }

  private commitSelection(next: XlsxSelectionState | null): void {
    this.setElementContext(null);
    const current = this.selectionState;
    if (selectionStatesEqual(current, next)) return;
    this.hideValidationPanel();
    this.selectionController.setState(next);
    this.updateSelectionOverlay();
    if (this.wb) this.scheduleRender();
    this.emitSelectionChange();
  }

  private setElementContext(context: XlsxElementContext | null): boolean {
    if (JSON.stringify(this.elementContext) === JSON.stringify(context)) return false;
    this.elementContext = context ? structuredClone(context) : null;
    this.updateSelectionOverlay();
    this.scheduleSelectionContextNotification();
    return true;
  }

  private scheduleSelectionContextNotification(): void {
    this.notifier.scheduleContextNotification();
  }

  private emitSelectionChange(): void {
    this.notifier.emit();
  }

  /** Refit automatic rows once after a column-resize gesture. Doing this on
   * every pointermove would turn a drag into O(sheet cells × pointer events),
   * while Excel's observable result only needs to be committed at release. */
  private refitAutoRowsAfterColumnResize(): void {
    const ws = this.currentWorksheet;
    const workbook = this.preparedWorkbook;
    if (!ws || !workbook) return;
    const manualRows = this.viewEdits.manualRows(this.currentSheet);
    invalidateAutoRowHeights(ws, manualRows);
    const prepareRowHeights = workbook[prepareXlsxViewerRowHeights];
    if (typeof prepareRowHeights !== 'function') return;
    const measureCanvas = this.hostDocument.createElement('canvas');
    const measureCtx = measureCanvas.getContext('2d');
    if (!measureCtx) return;
    prepareRowHeights.call(workbook, ws, measureCtx);
    this.viewEdits.syncAutomaticRowOverrides(this.currentSheet, ws);
    this.updateSpacerSize(ws);
    this.updateSelectionOverlay();
    this.scheduleRender();
  }

  /**
   * Change the cell-selection highlight color at runtime (see {@link
   * XlsxViewerOptions.selectionColor}). The border takes the color as-is and the
   * fill becomes a translucent shade of it; the current selection repaints
   * immediately.
   */
  setSelectionColor(color: string): void {
    this.opts.selectionColor = color;
    this.updateSelectionOverlay();
  }

  /**
   * Switch the hidden-sheet mode at runtime: restyle the tabs and re-render.
   * Entering `'skip'` while on a hidden sheet advances to the nearest visible.
   */
  async setHiddenSheetMode(mode: HiddenSheetMode): Promise<void> {
    this._hiddenSheetMode = mode;
    this.buildTabs();
    if (mode === 'skip' && this.wb && this.wb.isHidden(this.currentSheet)) {
      await this.showSheet(
        resolveVisibleIndex(this.currentSheet, (i) => this.wb!.isHidden(i), this.sheetCount),
      );
    } else {
      this.updateTabActive(this.currentSheet);
    }
  }

  /** The current hidden-sheet mode. */
  get hiddenSheetMode(): HiddenSheetMode { return this._hiddenSheetMode; }

  /** Number of non-hidden sheets (absolute `sheetCount` is unchanged). */
  get visibleSheetCount(): number {
    if (!this.wb) return 0;
    const wb = this.wb;
    return countVisible((i) => wb.isHidden(i), this.sheetCount);
  }

  /**
   * Copy the selected area as bounded TSV. The same limits apply regardless of
   * whether pointer, keyboard, or API created the selection.
   */
  async copySelection(): Promise<XlsxCopyResult> {
    this.assertOpen();
    if (this.previewCompletion) await this.previewCompletion;
    return this.copier.copy();
  }

  /** Rebuild the selection / object-context overlay for the viewport. */
  private updateSelectionOverlay(): void {
    this.selectionPaint.update();
  }

  /** Redraw the find-highlight boxes for the displayed sheet. */
  private updateFindOverlay(): void {
    this.finder.updateOverlay();
  }

  /**
   * IX2 — find every occurrence of `query` across every sheet and highlight the
   * matched cells. Returns every match in document order (sheet ascending, then
   * row-major within a sheet), each tagged with its
   * `{ sheet, sheetName, ref, row, col }`. A cell is the search unit: search
   * runs over each cell's *rendered* display text (number formats, dates, rich
   * text flattened), so a query matches what the grid shows. Case-insensitive by
   * default; pass `{ caseSensitive: true }` for an exact match. An empty query
   * clears the find.
   */
  async findText(
    query: string,
    opts: FindMatchesOptions = {},
  ): Promise<FindMatch<XlsxMatchLocation>[]> {
    return this.finder.find(query, opts);
  }

  /**
   * IX2 — move to the next match (wrap-around), switching sheets and scrolling
   * the matched cell into view as needed, and highlight it as the active match.
   * Returns the now-active match, or `null` when there are none. Call
   * {@link findText} first.
   */
  async findNext(): Promise<FindMatch<XlsxMatchLocation> | null> {
    return this.finder.next();
  }

  /** IX2 — move to the previous match (wrap-around). */
  async findPrev(): Promise<FindMatch<XlsxMatchLocation> | null> {
    return this.finder.prev();
  }

  /** IX2 — clear all highlights and reset the find state. */
  clearFind(): void {
    this.finder.clear();
  }

  /**
   * Scroll the grid so cell (row, col) is comfortably in view. Computes the
   * cell's absolute logical offset from the axis metrics (the same the renderer
   * uses) and nudges the vertical / start-anchored horizontal viewport
   * only when the cell is outside the scrollable viewport — an in-view cell is
   * left where it is (Excel's find behaviour). Frozen cells are always visible,
   * so they need no scroll.
   */
  private _scrollCellIntoView(
    row: number,
    col: number,
    align: NonNullable<XlsxScrollToCellOptions['align']> = 'nearest',
  ): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const cs = this.viewport.scale;
    const offset = getGridGeometryForWorksheet(ws).scrollOffsetForCell(
      row,
      col,
      {
        scale: cs,
        viewportWidth: this.canvasArea.clientWidth,
        viewportHeight: this.canvasArea.clientHeight,
        currentX: this.effectiveScrollLeft,
        currentY: this.viewportTop,
        headerWidth: HEADER_W,
        headerHeight: HEADER_H,
        align,
      },
    );
    this.viewportTop = offset.y;
    this.setViewportLeft(offset.x);
  }

  /** Close the list-validation panel and cancel a pending resolution. */
  private hideValidationPanel(): void {
    this.validation.hide();
  }

  // ─── Comment hover popup ──────────────────────────────────────────────────

  /** Index the displayed sheet's comments for the hover popup. */
  private buildCommentMap(ws: Worksheet): void {
    this.comments.setComments(ws.comments ?? []);
  }

  private createVisibleSheetView(source: Worksheet): Worksheet {
    const worksheet = createSheetViewModel(source);
    if (this.opts.comments === false) {
      const hidden = { ...worksheet, commentRefs: [], comments: [] };
      inheritWorksheetPreviewBounds(worksheet, hidden);
      inheritWorksheetPolicy(worksheet, hidden);
      return hidden;
    }
    // Keep the pre-customization behavior: XLSX historically exposed resolved
    // threaded comments. Consumers may explicitly hide them.
    const commentOptions = typeof this.opts.comments === 'object'
      ? this.opts.comments
      : undefined;
    if (commentOptions?.includeResolved !== false) return worksheet;
    const resolved = new Set(
      (worksheet.comments ?? [])
        .filter((comment) => comment.resolved === true)
        .map((comment) => comment.cellRef),
    );
    if (resolved.size === 0) return worksheet;
    const unresolved = {
      ...worksheet,
      commentRefs: worksheet.commentRefs?.filter((ref) => !resolved.has(ref)),
      comments: worksheet.comments?.filter((comment) => !resolved.has(comment.cellRef)),
    };
    inheritWorksheetPreviewBounds(worksheet, unresolved);
    inheritWorksheetPolicy(worksheet, unresolved);
    return unresolved;
  }

  /** IX1 — index the displayed sheet's hyperlinks for hover and click. */
  private buildHyperlinkMap(ws: Worksheet): void {
    this.hyperlinks.build(ws);
  }

  /** Hide the comment popup and cancel any pending show. */
  private hideCommentPopup(): void {
    this.comments.hide();
  }

  private buildTabs(): void {
    if (!this.sheetTabs) return;
    this.sheetTabs.build(this.workbook.sheetNames, this.workbook.tabColors);
  }

  private updateTabActive(index: number): void {
    this.sheetTabs?.setActive(index);
  }

  /**
   * IX9 {@link ZoomableViewer} — set the cell/header scale (`1` = 100%; the
   * viewer's `cellScale`) and re-lay-out the current sheet. Clamped to the zoom
   * bounds and snapped to whole percent; keeps the slider thumb, percentage label
   * in sync, and fires `onScaleChange` when the resolved scale actually changes.
   */
  setScale(scale: number): void {
    const zoomMin = this.opts.zoomMin ?? 0.1;
    const zoomMax = this.opts.zoomMax ?? 4;
    // Snap to whole percent so the label and cellScale stay tidy.
    const pct = Math.min(
      Math.round(zoomMax * 100),
      Math.max(Math.round(zoomMin * 100), Math.round(scale * 100)),
    );
    const next = pct / 100;
    const prevScale = this.viewport.scale;
    // Consume the gesture-only pointer anchor (Ctrl/⌘+wheel set it just above)
    // FIRST — before the no-op early return — so a gesture whose setScale ends
    // up a NO-OP (pinned at zoomMin/zoomMax, or a small deltaY swallowed by the
    // whole-percent snap) can never leak a stale anchor into a later non-gesture
    // setScale (slider, steppers, fitWidth/fitPage, public API), which must keep
    // the historical START-anchored (top-left) preservation. `null` for every
    // non-gesture source.
    const gestureAnchor = this._pendingZoomAnchor;
    this._pendingZoomAnchor = null;
    if (next === prevScale) return;
    this.viewport.setScale(next);

    this.zoomControl?.sync(next, pct, zoomMin, zoomMax);

    if (this.currentWorksheet) {
      // Preserve the START-anchored effective scroll position across the zoom.
      // The spacer (scrollWidth) is re-sized below, which changes maxScrollLeft;
      // for RTL the native scrollLeft is the inverse of the effective position,
      // so we must re-derive scrollLeft from the preserved effective value or
      // the view would jump toward the start on every zoom step.
      const prevEffective = this.effectiveScrollLeft;
      const prevScrollTop = this.viewportTop;
      // Gutter extents scale with cellScale (XL4); re-lay them out before the
      // spacer/scroll math reads canvasArea's new inset size.
      this.layoutGutters();
      this.updateSpacerSize(this.currentWorksheet);

      if (gestureAnchor) {
        // POINTER-ANCHORED zoom (both axes). The header + frozen band are drawn
        // at a FIXED screen position and do NOT scroll (see getCellAt), but their
        // on-screen size is the UNSCALED extent K × cs — a SCALING lead-in. From
        // getCellAt, the logical row under screen-y `py` is
        //   (py + scrollTop)/cs − K            (K = HEADER_H + frozenH)
        // and requiring that to be invariant across cs makes the K·cs terms
        // cancel exactly:
        //   scrollTop' = ratio·(scrollTop + py) − py
        // — i.e. the RAW pointer is the anchor and the clamp is the native
        // [0, maxScroll] (see anchoredZoomOffset's LEAD-INS note; routing through
        // a lead-in-shifted virtual scroll would distort the low clamp and floor
        // scrollTop at K·cs near the sheet start).

        // Vertical: native scrollTop is start-anchored in both LTR/RTL.
        this.viewportTop = anchoredZoomOffset(prevScrollTop, gestureAnchor.y, prevScale, next, {
          maxScroll: this.maxScrollTop,
        });

        // Horizontal: anchor in the logical-LTR space the grid math uses (the
        // same cancellation holds for K = HEADER_W + frozenW), so RTL is handled
        // by translating the pointer through screenX (an involution) and
        // re-deriving the native scrollLeft from the effective (start-anchored)
        // position, exactly as the START-anchored branch does.
        const anchorLogicalX = this.screenX(gestureAnchor.x, 0);
        const maxLeftV = this.maxScrollLeft;
        const newEffective = anchoredZoomOffset(prevEffective, anchorLogicalX, prevScale, next, {
          maxScroll: maxLeftV,
        });
        this.setViewportLeft(newEffective);
      } else {
        this.setViewportLeft(prevEffective);
      }
    }
    void this.renderCurrentSheet().catch((error) => this._reportRenderError(error));
    this.updateSelectionOverlay();
    this.updateFindOverlay();
    this.sheetTabs?.updateNavButtons();
    // IX9 change notification (fired last, after the view is consistent). Only
    // reached when `next` differs from the prior scale (early-returned above).
    this.opts.onScaleChange?.(next);
  }

  /** IX9 {@link ZoomableViewer} — the current zoom factor (`1` = 100%). This is
   *  the viewer's `cellScale`; `1` before anything is set. */
  getScale(): number {
    return this.viewport.scale;
  }

  /** IX9 {@link ZoomableViewer} — step up to the next rung of the shared zoom
   *  ladder (clamped to `zoomMax` by {@link setScale}). */
  zoomIn(): void {
    this.setScale(nextZoomStep(this.getScale()));
  }

  /** IX9 {@link ZoomableViewer} — step down to the next lower ladder rung. */
  zoomOut(): void {
    this.setScale(prevZoomStep(this.getScale()));
  }

  /**
   * IX9 {@link ZoomableViewer} — fit the used data range's WIDTH to the canvas
   * area. The "content" is the natural (100%) width of the row header plus the
   * used columns; the container is `canvasArea.clientWidth`. A no-op (defers) when
   * nothing is loaded or the container is unlaid-out. Routes through
   * {@link setScale}, so the result is clamped/snapped and fires `onScaleChange`.
   */
  fitWidth(): void {
    this._fit('width');
  }

  /**
   * IX9 {@link ZoomableViewer} — fit the used data range's WIDTH AND HEIGHT inside
   * the canvas area (header + used columns/rows), so the whole used range is
   * visible without scrolling. Takes the tighter of the width- and height-fit
   * factors. Defers when unloaded / unlaid-out; routes through {@link setScale}.
   */
  fitPage(): void {
    this._fit('page');
  }

  /** Shared fit implementation for {@link fitWidth} / {@link fitPage}: derive the
   *  natural (cs=1) content extent of the used data range, ask core's pure
   *  {@link fitScale} for the factor, and apply it via {@link setScale}. */
  private _fit(mode: 'width' | 'page'): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const { width, height } = this._naturalContentExtent(ws);
    const scale = fitScale(
      {
        contentWidth: width,
        contentHeight: height,
        containerWidth: this.canvasArea.clientWidth,
        containerHeight: this.canvasArea.clientHeight,
      },
      mode,
    );
    if (scale <= 0) return; // unlaid-out / empty — defer (fitScale's 0 sentinel)
    this.setScale(scale);
  }

  /** Natural (unscaled, cs=1) CSS-px extent of a worksheet's used data range:
   *  the row/column header plus every used column width / row height. Uses the
   *  same content bounds and configured minimums as {@link updateSpacerSize},
   *  but deliberately excludes its trailing scroll headroom. */
  private _naturalContentExtent(ws: Worksheet): { width: number; height: number } {
    const { maxRow, maxCol } = worksheetContentBounds(ws, this.gridExtent);
    return getGridGeometryForWorksheet(ws).logicalContentExtent(
      maxRow,
      maxCol,
      HEADER_W,
      HEADER_H,
    );
  }

  private updateSpacerSize(ws: Worksheet): void {
    const cs = this.viewport.scale;
    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // Find actual scrollable data extent
    let { maxRow, maxCol } = worksheetContentBounds(ws, this.gridExtent);
    maxRow += this.gridExtent.marginRows;
    maxCol += this.gridExtent.marginCols;

    // Spacer = rounded header + cumulative per-band-rounded geometry.
    const extent = getGridGeometryForWorksheet(ws).roundedContentExtent(
      maxRow,
      maxCol,
      cs,
      HEADER_W,
      HEADER_H,
    );
    const totalW = extent.width;
    const totalH = extent.height;

    this.spacer.style.width = `${totalW}px`;
    this.spacer.style.height = `${totalH}px`;
    this.viewport.setViewportSize(this.scrollHost.clientWidth, this.scrollHost.clientHeight);
    this.viewport.setExtent(totalW, totalH);
    this.setViewportLeft(this.viewport.x);
    this.viewportTop = this.viewport.y;
  }

  /**
   * Coalesce a re-render into the next animation frame. Called from the
   * high-frequency event-driven paths (scroll, live column/row resize, drag-
   * selection, container resize); a burst of these within one frame schedules a
   * single {@link renderCurrentSheet}, avoiding the previous behavior where every
   * scroll event forced its own synchronous full redraw. Already-scheduled frames
   * are not re-scheduled — the one pending render reads the live scroll/scale
   * state when it runs, so the most recent position always wins without threading
   * a coordinate through. Falls back to a synchronous render when
   * `requestAnimationFrame` is unavailable (e.g. a non-DOM host), preserving the
   * old semantics there.
   */
  private scheduleRender(): void {
    this.renderDispatcher.schedule(() =>
      this.renderCurrentSheet().catch((error) => this._reportRenderError(error)));
  }

  private async renderCurrentSheet(): Promise<void> {
    const generation = this.renderDispatcher.begin();
    const rowPreview = this.selectionInput.rowResizePreview;
    try {
      await this._renderCurrentSheet(generation);
    } catch (err) {
      if (!this.renderDispatcher.isCurrent(generation)) return;
      if (rowPreview) this.selectionInput.abortRowResizePreview(rowPreview);
      throw err;
    }
  }

  /** Route a render failure to `onError`, or `console.error` when none is given
   *  (never fully silent), and never after teardown. Mirrors the scroll viewers'
   *  `_reportRenderError`. */
  private _reportRenderError(err: unknown): void {
    if (this._destroyed) return;
    const e = err instanceof Error ? err : new Error(String(err));
    if (this.opts.onError) this.opts.onError(e);
    else console.error('[ooxml] XlsxViewer render failed:', e);
  }

  private async _renderCurrentSheet(seq: number): Promise<void> {
    if (!this.currentWorksheet) return;
    if (this.previewCompletion) {
      if (this.firstPreviewRender) {
        const prepared = this.previewPreparedViewport;
        if (!prepared || this.canvasArea.clientWidth !== prepared.width ||
            this.canvasArea.clientHeight !== prepared.height ||
            this.viewport.scale !== prepared.scale ||
            this.viewportTop !== 0 || this.effectiveScrollLeft !== 0) return;
      } else await this.previewCompletion;
      if (!this.renderDispatcher.isCurrent(seq) || this._destroyed) return;
    }
    const ws = this.currentWorksheet;
    const w = this.canvasArea.clientWidth;
    const h = this.canvasArea.clientHeight;
    if (w <= 0 || h <= 0) return;

    // Claim a render generation up front so a later render started while this one
    // awaits the worker can mark this frame stale (worker mode only; see below).
    const cs = this.viewport.scale;
    const dpr = this.surface.dpr;

    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // DOM scrollLeft/scrollTop are in scaled (physical) CSS pixels.
    // Convert to logical pixels for cell-finding by dividing by cs. For RTL
    // sheets effectiveScrollLeft inverts the native scrollLeft so that 0 = col A
    // at the (mirrored) right edge — see the getter for the rationale.
    const visible = getGridGeometryForWorksheet(ws).visibleRange({
      width: w,
      height: h,
      scale: cs,
      scrollX: this.effectiveScrollLeft,
      scrollY: this.viewportTop,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
      buffer: 2,
    });
    const viewport: ViewportRange = visible.range;
    const { offsetX, offsetY } = visible;

    const { selectedRowRange, selectedColRange } = this.computeHeaderHighlight();

    const renderOpts = {
      width: w,
      height: h,
      dpr,
      imageResources: this.opts.imageResources,
      cellScale: cs,
      scrollOffsetX: offsetX,
      scrollOffsetY: offsetY,
      freezeRows,
      freezeCols,
      selectedRowRange,
      selectedColRange,
      chromeColors: this.chromeColors,
    };

    const sizeProjection = this.wireSizeOverrides();
    // #1713: the worker needs the bound reference even before any size edit.
    // It is fixed for this workbook/sheet, so the existing revision remains
    // the cache key (a changed reference without a revision bump is an error).
    // A degraded parser placeholder has its own graph. Keep the reference for
    // later valid reacquisition, but never transport it onto that placeholder.
    const initialAnchorSizes = ws.parseError
      ? undefined
      : this.initialAnchorReferenceFor(this.workbook, this.currentSheet) ?? undefined;
    const projection = sizeProjection || initialAnchorSizes
      ? {
          id: this.projectionId,
          revision: this.viewEdits.sizeRevision(this.currentSheet),
          autoRowHeightsPrepared: this.currentAutoRowHeightsPrepared,
          ...(initialAnchorSizes ? { initialAnchorSizes } : {}),
        }
      : undefined;
    const viewerRenderOpts = withViewerRenderContext(
      sizeProjection ? { ...renderOpts, sizeOverrides: sizeProjection.overrides } : renderOpts,
      getGridGeometryForWorksheet(ws).maximumDigitWidth,
      { worksheet: ws, projection },
    );

    if (this._mode === 'worker') {
      // Render the viewport off the main thread and paint the returned bitmap.
      // The selection overlay (geometry-based, from getCellRect) is unaffected.
      // Attach the cumulative view-only size overrides (outline collapse/
      // expand, drag resize) so the worker re-lays the mutated bands — its
      // local sheet cache never sees main-thread model writes on its own.
      const bmp = await this.workbook.renderViewportToBitmap(
        this.currentSheet,
        viewport,
        viewerRenderOpts,
      );
      if (!this.renderDispatcher.commitBitmap(seq, bmp, w, h)) return;
    } else {
      await this.workbook.renderViewport(
        this.canvas,
        this.currentSheet,
        viewport,
        withXlsxRenderCommitGuard(viewerRenderOpts, () =>
          !this._destroyed && this.renderDispatcher.isCurrent(seq),
        ),
      );
      if (!this.renderDispatcher.isCurrent(seq) || this._destroyed) return;
    }
    // XL4: repaint the outline gutters over the fresh grid frame, aligned to the
    // same scroll offset. No-op when the sheet has no outlining.
    this.renderGutters();
    this.firstPreviewRender = false;
    this.committedFrameCount++;
  }

  private computeHeaderHighlight(): {
    selectedRowRange: { start: number; end: number; strong: boolean } | null;
    selectedColRange: { start: number; end: number; strong: boolean } | null;
  } {
    return this.selectionController.headerHighlight();
  }

  get sheetNames(): string[] {
    return this.wb?.sheetNames ?? [];
  }

  /** The underlying <canvas> element the grid is drawn on. */
  get canvasElement(): HTMLCanvasElement {
    return this.canvas;
  }

  /** Latest content-free resource metrics for the loaded workbook. */
  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    if (!this.wb) throw new Error('Workbook not loaded');
    return await this.wb.getResourceMetrics();
  }

  /**
   * Tear down the viewer and release resources.
   *
   * The caller's container is returned to the state it had before construction
   * (empty): the entire wrapper subtree the constructor appended is removed.
   * Every listener, observer and frame is released: each collaborator detaches
   * the listeners it registered (viewport input, outline gutters, tab strip,
   * zoom control, validation panel and its document-level outside-click
   * handler, overlay host), and the chrome theme and comment popup disconnect
   * their observers. Safe to call more than once.
   *
   * NOTE: the shared `<style>` in the owning document is intentionally NOT removed —
   * it is a class constant that any still-live viewer may depend on, and one
   * leftover sheet is a bounded, harmless cost (see {@link ensureViewerStyleInjected}).
   */
  destroy(): void {
    if (this._destroyed) return;
    // First line: block any render rejection racing in from surfacing on a dead
    // viewer (checked at the top of _reportRenderError). The acquisition owner
    // invalidates any load still in flight below.
    this._destroyed = true;
    this.notifier.destroy();
    this.selectionInput.destroy();
    this.sheetRequestGeneration++;
    this.resizeObserver?.disconnect();
    this.chromeTheme.destroy();
    this.renderDispatcher.destroy();
    this.surface.destroy();
    this.overlayHost.destroy();
    this.sheetTabs?.destroy();
    this.zoomControl?.destroy();
    this.comments.destroy();
    this.validation.destroy();
    // IX2 — drop the find state (matches + cursor) so a stale
    // findNext()/findPrev() after teardown returns null instead of a match
    // pointing into a dead viewer (same fix as DocxViewer/PptxViewer.destroy).
    this.finder.destroy();
    this.releaseHostFonts();
    const releaseProjection = this.wb?.[releaseXlsxViewerProjection];
    if (typeof releaseProjection === 'function') {
      releaseProjection.call(this.wb, this.projectionId);
    }
    this.currentWorksheet = null;
    this.releaseCurrentWorksheet?.();
    this.releaseCurrentWorksheet = null;
    this.sheetViews.clear();
    this.releaseInitialAnchorReferences();
    this.viewEdits.destroy();
    this.currentSourceComments = [];
    this.sourceCommentMap.clear();
    this.hyperlinks.destroy();
    this.preparedWorkbook = null;
    this.outlineGutter.destroy();
    this.elementContext = null;
    this.selectionController.reset();
    this.acquisition.destroy();
    // Remove the whole UI subtree so the container is empty again.
    this.wrapper.remove();
  }

  private assertOpen(): void {
    if (this._destroyed) throw this.destroyedError();
  }

  private destroyedError(): Error {
    return new Error(this._mountKind === 'sheet'
      ? 'XlsxSheetViewer is destroyed'
      : 'XlsxViewer is destroyed');
  }
}

/** Workbook viewer mounted into a container with scrollable grid, sheet tabs,
 * outline gutters, and optional zoom chrome. */
export class XlsxViewer extends XlsxViewerEngine {
  /**
   * Create a workbook Viewer that borrows an already-loaded workbook.
   * Destroying the Viewer leaves the caller-owned workbook open.
   */
  static fromWorkbook(
    container: HTMLElement,
    workbook: XlsxWorkbook,
    opts: Omit<XlsxViewerOptions, keyof LoadOptions> = {},
  ): Omit<XlsxViewer, 'load'> {
    return new XlsxViewer(container, {
      ...opts,
      [borrowedWorkbookOption]: workbook,
    } as InternalXlsxViewerOptions);
  }

  constructor(container: HTMLElement, opts: XlsxViewerOptions = {}) {
    super(container, opts, { kind: 'composite' });
  }

  /** Load an OOXML workbook from a URL or ArrayBuffer. */
  async load(source: string | ArrayBuffer): Promise<void> {
    await this[loadXlsxViewerSource](source);
  }
}

type XlsxSheetViewerSnapshot = Readonly<{
  sheetIndex: number;
  sheetCount: number;
  sheetNames: string[];
  viewport: XlsxViewportOffset;
  selectionState: XlsxSelectionState | null;
  scale: number;
  hiddenSheetMode: HiddenSheetMode;
  visibleSheetCount: number;
}>;

/**
 * Canvas-mounted active-sheet viewer. It instantiates the same workbook,
 * acquisition, geometry, selection, overlay, and render-dispatch engine as
 * {@link XlsxViewer}, but mounts no workbook footer or sheet-tab chrome.
 */
export class XlsxSheetViewer implements ZoomableViewer {
  private readonly engine: XlsxViewerEngine;
  private readonly canvasMount: CallerCanvasMount;
  private destroyed = false;
  private snapshot: XlsxSheetViewerSnapshot;
  private lastMetrics: OoxmlResourceMetrics | undefined;

  /**
   * Create a sheet Viewer that borrows an already-loaded workbook.
   * Destroying the Viewer leaves the caller-owned workbook open.
   */
  static fromWorkbook(
    canvasElement: HTMLCanvasElement,
    workbook: XlsxWorkbook,
    options: Omit<XlsxSheetViewerOptions, keyof LoadOptions> = {},
  ): Omit<XlsxSheetViewer, 'load'> {
    return new XlsxSheetViewer(canvasElement, {
      ...options,
      [borrowedWorkbookOption]: workbook,
    } as InternalXlsxViewerOptions);
  }

  constructor(
    readonly canvasElement: HTMLCanvasElement,
    options: XlsxSheetViewerOptions = {},
  ) {
    const borrowedWorkbook = (options as InternalXlsxViewerOptions)[borrowedWorkbookOption];
    const mode = resolveCanvasViewerMode('XlsxSheetViewer', options.mode, borrowedWorkbook);
    const rect = canvasElement.getBoundingClientRect();
    this.canvasMount = new CallerCanvasMount(canvasElement, {
      wrapperCssText:
        `position:relative;display:inline-block;vertical-align:top;overflow:hidden;` +
        `width:${canvasElement.style.width || `${rect.width || canvasElement.width}px`};` +
        `height:${canvasElement.style.height || `${rect.height || canvasElement.height}px`};`,
      restoreMode: 'style-and-bitmap',
    });
    this.engine = new XlsxViewerEngine(this.canvasMount.wrapper, {
      ...options,
      onResourceMetrics: (metrics) => {
        this.lastMetrics = metrics;
        options.onResourceMetrics?.(metrics);
      },
    }, {
      kind: 'sheet',
      canvas: canvasElement,
      mode,
    });
    this.snapshot = {
      sheetIndex: 0,
      sheetCount: 0,
      sheetNames: [],
      viewport: { x: 0, y: 0 },
      selectionState: null,
      scale: this.engine.getScale(),
      hiddenSheetMode: this.engine.hiddenSheetMode,
      visibleSheetCount: 0,
    };
  }

  /**
   * Load an XLSX worksheet, or reuse the XLSX sheet renderer for one explicitly
   * selected delimited-text source. URL/ArrayBuffer input, reload replacement,
   * callbacks, and destroy ownership match XLSX loading.
   */
  async load(
    source: string | ArrayBuffer,
    options: XlsxSheetLoadOptions = {},
  ): Promise<void> {
    this.assertOpen();
    try {
      await this.engine[loadXlsxViewerSource](source, options);
    } finally {
      if (!this.destroyed) this.captureSnapshot();
    }
    this.assertOpen();
  }

  get sheetIndex(): number { return this.destroyed ? this.snapshot.sheetIndex : this.engine.sheetIndex; }
  get sheetCount(): number { return this.destroyed ? this.snapshot.sheetCount : this.engine.sheetCount; }
  get sheetNames(): string[] {
    return this.destroyed ? [...this.snapshot.sheetNames] : [...this.engine.sheetNames];
  }

  async goToSheet(index: number): Promise<void> {
    this.assertOpen();
    await this.engine.goToSheet(index);
    this.assertOpen();
    this.captureSnapshot();
  }

  async nextSheet(): Promise<void> {
    this.assertOpen();
    await this.engine.nextSheet();
    this.assertOpen();
    this.captureSnapshot();
  }

  async prevSheet(): Promise<void> {
    this.assertOpen();
    await this.engine.prevSheet();
    this.assertOpen();
    this.captureSnapshot();
  }

  getViewportOffset(): XlsxViewportOffset {
    return this.destroyed ? { ...this.snapshot.viewport } : this.engine.getViewportOffset();
  }

  async setViewportOffset(offset: XlsxViewportOffset): Promise<void> {
    this.assertOpen();
    await this.engine.setViewportOffset(offset);
    this.assertOpen();
    this.captureSnapshot();
  }

  async scrollToCell(ref: string, options?: XlsxScrollToCellOptions): Promise<void> {
    this.assertOpen();
    await this.engine.scrollToCell(ref, options);
    this.assertOpen();
    this.captureSnapshot();
  }

  async relayout(): Promise<void> {
    this.assertOpen();
    // The caller canvas remains the sizing authority. A caller may update its
    // inline/CSS box and then call relayout(); promote that box to the mount so
    // the shared engine measures the new viewport rather than its old wrapper.
    const rect = this.canvasElement.getBoundingClientRect();
    if (rect.width > 0) this.canvasMount.wrapper.style.width = `${rect.width}px`;
    if (rect.height > 0) this.canvasMount.wrapper.style.height = `${rect.height}px`;
    await this.engine.relayout();
    this.assertOpen();
    this.captureSnapshot();
  }

  getScale(): number { return this.destroyed ? this.snapshot.scale : this.engine.getScale(); }

  setScale(scale: number): void {
    this.assertOpen();
    this.engine.setScale(scale);
    this.captureSnapshot();
  }

  zoomIn(): void { this.assertOpen(); this.engine.zoomIn(); this.captureSnapshot(); }
  zoomOut(): void { this.assertOpen(); this.engine.zoomOut(); this.captureSnapshot(); }
  fitWidth(): void { this.assertOpen(); this.engine.fitWidth(); this.captureSnapshot(); }
  fitPage(): void { this.assertOpen(); this.engine.fitPage(); this.captureSnapshot(); }

  getCellAt(clientX: number, clientY: number): CellAddress | null {
    return this.destroyed ? null : this.engine.getCellAt(clientX, clientY);
  }

  getCellViewportRect(cell: CellAddress | string): XlsxCellViewportRect | null {
    return this.destroyed ? null : this.engine.getCellViewportRect(cell);
  }

  /** Detached comments for the current sheet, in authored order. */
  getComments(): readonly Readonly<XlsxComment>[] {
    this.assertOpen();
    return this.engine.getComments();
  }

  async goToComment(
    sheetIndex: number,
    cellRef: string,
    options?: XlsxScrollToCellOptions,
  ): Promise<boolean> {
    this.assertOpen();
    const found = await this.engine.goToComment(sheetIndex, cellRef, options);
    this.assertOpen();
    this.captureSnapshot();
    return found;
  }

  get selectionState(): XlsxSelectionState | null {
    const value = this.destroyed ? this.snapshot.selectionState : this.engine.selectionState;
    return value ? structuredClone(value) : null;
  }

  setSelection(selection: XlsxSelectionInput): void {
    this.assertOpen();
    this.engine.setSelection(selection);
    this.captureSnapshot();
  }

  getSelectionContext(options?: XlsxSelectionContextOptions): XlsxSelectionContext | null {
    this.assertOpen();
    return this.engine.getSelectionContext(options);
  }

  async copySelection(): Promise<XlsxCopyResult> {
    this.assertOpen();
    return await this.engine.copySelection();
  }

  setSelectionColor(color: string): void {
    this.assertOpen();
    this.engine.setSelectionColor(color);
  }

  async setHiddenSheetMode(mode: HiddenSheetMode): Promise<void> {
    this.assertOpen();
    await this.engine.setHiddenSheetMode(mode);
    this.assertOpen();
    this.captureSnapshot();
  }

  get hiddenSheetMode(): HiddenSheetMode {
    return this.destroyed ? this.snapshot.hiddenSheetMode : this.engine.hiddenSheetMode;
  }

  get visibleSheetCount(): number {
    return this.destroyed ? this.snapshot.visibleSheetCount : this.engine.visibleSheetCount;
  }

  async findText(
    query: string,
    options?: FindMatchesOptions,
  ): Promise<FindMatch<XlsxMatchLocation>[]> {
    this.assertOpen();
    const matches = await this.engine.findText(query, options);
    this.assertOpen();
    return matches;
  }

  async findNext(): Promise<FindMatch<XlsxMatchLocation> | null> {
    this.assertOpen();
    const match = await this.engine.findNext();
    this.assertOpen();
    this.captureSnapshot();
    return match;
  }

  async findPrev(): Promise<FindMatch<XlsxMatchLocation> | null> {
    this.assertOpen();
    const match = await this.engine.findPrev();
    this.assertOpen();
    this.captureSnapshot();
    return match;
  }

  clearFind(): void { this.assertOpen(); this.engine.clearFind(); }

  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    if (this.destroyed) {
      if (this.lastMetrics) return this.lastMetrics;
      throw this.destroyedError();
    }
    this.lastMetrics = await this.engine.getResourceMetrics();
    return this.lastMetrics;
  }

  destroy(): void {
    if (this.destroyed) return;
    this.captureSnapshot();
    this.destroyed = true;
    this.engine.destroy();

    this.canvasMount.restore();
  }

  private captureSnapshot(): void {
    const selectionState = this.engine.selectionState;
    this.snapshot = {
      sheetIndex: this.engine.sheetIndex,
      sheetCount: this.engine.sheetCount,
      sheetNames: [...this.engine.sheetNames],
      viewport: { ...this.engine.getViewportOffset() },
      selectionState: selectionState ? structuredClone(selectionState) : null,
      scale: this.engine.getScale(),
      hiddenSheetMode: this.engine.hiddenSheetMode,
      visibleSheetCount: this.engine.visibleSheetCount,
    };
  }

  private assertOpen(): void {
    if (this.destroyed) throw this.destroyedError();
  }

  private destroyedError(): Error {
    return new Error('XlsxSheetViewer is destroyed');
  }
}
