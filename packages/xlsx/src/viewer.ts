import {
  XlsxWorkbook,
  prepareXlsxViewerRowHeights,
  releaseXlsxViewerProjection,
  retainXlsxViewerFonts,
} from './workbook.js';
import type { LoadOptions } from './workbook.js';
import type { Cell, Hyperlink, Row, ViewportRange, Worksheet, XlsxChromeColors, XlsxComment } from './types.js';
import type { FindHighlightColors, HyperlinkTarget, FindMatch, FindMatchesOptions, OoxmlResourceMetrics, ViewerContextMenuEvent, ZoomableViewer } from '@silurus/ooxml-core';
import { nextVisibleIndex, resolveVisibleIndex, countVisible, zoomStepScale, anchoredZoomOffset, openExternalHyperlink, nextZoomStep, prevZoomStep, fitScale } from '@silurus/ooxml-core';
import {
  CallerCanvasMount,
  resolveCanvasViewerMode,
  type CanvasViewerRenderMode,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import type { ReadOnlyCommentThread } from '@silurus/ooxml-core/internal/read-only-comment-contract';
import {
  HEADER_W,
  HEADER_H,
  pxToColWidth,
  pxToRowHeight,
  invalidateAutoRowHeights,
  derivedAutoRowHeights,
  getGridGeometryForWorksheet,
  rtlMirrorX,
} from './renderer.js';
import { findListValidationAt } from './data-validation.js';
import { formatA1, parseA1 } from './a1.js';
import { resolveXlsxInternalHyperlink } from './internal-hyperlink.js';
import type {
  CellAddress,
  XlsxSelectionArea,
  XlsxSelectionContext,
  XlsxSelectionContextCell,
  XlsxSelectionContextOptions,
  XlsxElementContext,
  XlsxSelectionInput,
  XlsxSelectionState,
} from './selection.js';
import {
  hitTestXlsxElementContext,
  limitXlsxElementContext,
  projectXlsxElementContext,
  type XlsxElementHitViewport,
} from './element-context.js';
import {
  MAX_SELECTION_CONTEXT_CELLS,
  MAX_SELECTION_CONTEXT_TEXT_CHARACTERS,
  areaContainsCell,
  normalizeSelectionState,
  selectionCoordinateCountUpperBound,
  selectionStateFromReference,
  selectionStatesEqual,
} from './selection.js';
export type { CellAddress } from './selection.js';
import { XlsxFindController, type FindCell, type XlsxMatchLocation } from './find.js';
import { computeCommentPopupPosition } from './comment-popup.js';
import type { XlsxCommentsOptions } from './comment-card.js';
import {
  computeValidationPanelPosition,
  type ResolvedList,
} from './validation-list.js';
import { withViewerRenderContext, type WireSizeOverrides } from './worker-protocol.js';
import {
  buildOutlineLayout,
  toggleGroupHidden,
  levelButtonHidden,
  rowBands,
  colBands,
  summaryAfterFor,
  gutterExtentPx,
  outlineBracketSegments,
  outlineLevelButtonCenterPx,
  outlinePaneClipRect,
  OUTLINE_BUTTON_PX,
  OUTLINE_LANE_PX,
  type BandOutline,
  type OutlineGroup,
  type OutlineLayout,
  type OutlineAxis,
} from './outline.js';
import {
  GridGeometry,
  MAX_WORKSHEET_COL,
  MAX_WORKSHEET_ROW,
} from './internal/grid-geometry.js';
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
import { selectionAutoScrollVelocity } from './selection-auto-scroll.js';
import { worksheetContentBounds } from './internal/worksheet-content-bounds.js';

const borrowedWorkbookOption = Symbol('XlsxViewer.borrowedWorkbook');
type XlsxCommentUiRuntime = typeof import('./comment-ui-runtime.js');
let xlsxCommentUiRuntimePromise: Promise<XlsxCommentUiRuntime> | undefined;

function loadXlsxCommentUiRuntime(): Promise<XlsxCommentUiRuntime> {
  return xlsxCommentUiRuntimePromise ??= import('./comment-ui-runtime.js');
}

// Re-exported for the existing xlsx zoom tests (resize-zoom.test.ts imports it
// from this module) and any consumer that referenced it here before it moved to
// @silurus/ooxml-core. The single source of truth is core (design §5.2).
export { zoomStepScale } from '@silurus/ooxml-core';

/** Delay (ms) before a hovered comment popup appears. A short hover dwell
 *  prevents the popup from flickering while the cursor sweeps across many
 *  commented cells; ~150ms is the common tooltip-show threshold (responsive yet
 *  long enough to suppress transient passes). Excel itself uses a comparable
 *  short hover delay before showing a note. */
const COMMENT_POPUP_DELAY_MS = 150;
/** Max width of the comment popup body (CSS px). */
const COMMENT_POPUP_MAX_W = 280;
/** Max height before the body scrolls/clips (CSS px). */
const COMMENT_POPUP_MAX_H = 200;

/** Max width of the list-validation dropdown panel (CSS px). */
const VALIDATION_PANEL_MAX_W = 240;
/** Max height before the value list scrolls (CSS px). */
const VALIDATION_PANEL_MAX_H = 200;

const TAB_BAR_H = 30;
// Footer chrome stays in screen pixels: sheet zoom scales grid cells and their
// row/column headers, but must not resize the tab-navigation controls.
const TAB_NAV_W = HEADER_W;
// Gap between adjacent sheet tabs. The first tab also gets this much leading
// space so it is offset from the row-header boundary by the same margin that
// separates tabs from each other.
const TAB_GAP = 1;
let nextViewerProjectionId = 1;

/** How {@link XlsxViewer} presents hidden sheets (`<sheet state>`, §18.2.19). */
export type HiddenSheetMode = 'show' | 'skip' | 'dim';

/** `'dim'`-mode tab opacity: hidden/veryHidden tabs are greyed but selectable.
 *  A UI-presentation default (ECMA-376 defines no hidden-tab rendering); mirrors
 *  the named pptx `DEFAULT_HIDDEN_DIM` constant. */
const HIDDEN_TAB_DIM_OPACITY = 0.45;

/** Marker attribute on the single injected viewer stylesheet, so the module-
 *  level injector is idempotent and destroy() can leave it in place. */
const VIEWER_STYLE_ATTR = 'data-xlsx-viewer-styles';

/** Class-constant CSS shared by every XlsxViewer: it styles pseudo-elements
 *  (scrollbar, slider track/thumb) that inline `element.style` cannot reach, so
 *  it must live in a stylesheet rather than on the elements. */
const VIEWER_STYLE_CSS =
  `.xlsx-tab-strip::-webkit-scrollbar{display:none}` +
  // The viewport remains focusable so copy shortcuts belong to the active
  // Viewer. Pointer focus stays quiet, while keyboard focus remains visible at
  // the viewport boundary and distinct from the selected-cell border.
  `[data-xlsx-viewport-input]:focus{outline:none}` +
  `[data-xlsx-viewport-input]:focus-visible{outline:2px solid var(--ooxml-xlsx-focus-ring,#2563eb);outline-offset:-2px}` +
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

const XLSX_CHROME_COLOR_PROPERTIES = {
  background: '--ooxml-xlsx-chrome-background',
  surface: '--ooxml-xlsx-chrome-surface',
  mutedSurface: '--ooxml-xlsx-chrome-surface-muted',
  text: '--ooxml-xlsx-chrome-text',
  mutedText: '--ooxml-xlsx-chrome-text-muted',
  border: '--ooxml-xlsx-chrome-border',
  selectedSurface: '--ooxml-xlsx-chrome-selection-background',
  accent: '--ooxml-xlsx-chrome-accent',
} as const satisfies Record<keyof XlsxChromeColors, string>;

function sameChromeColors(left: XlsxChromeColors, right: XlsxChromeColors): boolean {
  return Object.keys(XLSX_CHROME_COLOR_PROPERTIES).every((key) =>
    left[key as keyof XlsxChromeColors] === right[key as keyof XlsxChromeColors]);
}

export interface XlsxSheetViewerOptions extends LoadOptions {
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

type SelectionInterval = Readonly<{ first: number; last: number }>;

function mergeSelectionIntervals(intervals: readonly SelectionInterval[]): SelectionInterval[] {
  const sorted = [...intervals].sort((a, b) => a.first - b.first || a.last - b.last);
  const merged: SelectionInterval[] = [];
  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (!previous || interval.first > previous.last + 1) {
      merged.push({ ...interval });
    } else if (interval.last > previous.last) {
      merged[merged.length - 1] = { first: previous.first, last: interval.last };
    }
  }
  return merged;
}

function intervalContains(intervals: readonly SelectionInterval[], value: number): boolean {
  let low = 0;
  let high = intervals.length - 1;
  while (low <= high) {
    const middle = (low + high) >>> 1;
    const interval = intervals[middle];
    if (value < interval.first) high = middle - 1;
    else if (value > interval.last) low = middle + 1;
    else return true;
  }
  return false;
}

function lowerBoundBy<T>(items: readonly T[], value: number, key: (item: T) => number): number {
  let low = 0;
  let high = items.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (key(items[middle]) < value) low = middle + 1;
    else high = middle;
  }
  return low;
}

function orderedBy<T>(items: readonly T[], key: (item: T) => number): readonly T[] {
  for (let index = 1; index < items.length; index++) {
    if (key(items[index - 1]) > key(items[index])) {
      return [...items].sort((left, right) => key(left) - key(right));
    }
  }
  return items;
}

/** Default cell-selection accent (Google blue), used when no `selectionColor`
 *  option is supplied. */
const DEFAULT_SELECTION_COLOR = '#1a73e8';

/** Half-width (CSS px) of the grab zone around a header border for
 *  drag-to-resize (issue #567), and the minimum size a column/row can be
 *  dragged to (logical px) so a collapsed band keeps a grabbable border. */
const RESIZE_GRAB_PX = 4;
const RESIZE_MIN_PX = 5;
// Keep clipboard materialization within the same hard cell-count envelope as a
// worksheet. A sparse range can span billions of coordinates even when only a
// handful of cells are populated, so its rectangular TSV must never be built.
const MAX_CLIPBOARD_CELLS = 250_000;
// Bound the retained TSV and its final joined copy. This is a resource-safety
// contract, not a worksheet semantic limit; callers can handle `too-large`
// without the viewer attempting an unbounded JavaScript string allocation.
const MAX_CLIPBOARD_UTF16_CODE_UNITS = 8 * 1_024 * 1_024;
const DEFAULT_SELECTION_CONTEXT_TEXT_CHARACTERS = 1 * 1_024 * 1_024;
const DEFAULT_SELECTION_CONTEXT_NOTIFICATION_TEXT_CHARACTERS = 65_536;
const MAX_SELECTION_CONTEXT_FIELD_CHARACTERS = 65_536;
const MAX_REENTRANT_SELECTION_NOTIFICATIONS = 100;

function safeUtf16Prefix(value: string, maxCodeUnits: number): string {
  let end = Math.min(value.length, Math.max(0, maxCodeUnits));
  if (end > 0 && end < value.length) {
    const previous = value.charCodeAt(end - 1);
    const next = value.charCodeAt(end);
    if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) end--;
  }
  return value.slice(0, end);
}

function encodeTsvFieldWithin(value: string, remaining: number): string | null {
  let quoteCount = 0;
  let needsQuotes = false;
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    if (code === 34) { quoteCount++; needsQuotes = true; }
    else if (code === 9 || code === 10 || code === 13) needsQuotes = true;
  }
  const length = value.length + (needsQuotes ? quoteCount + 2 : 0);
  if (length > remaining) return null;
  return needsQuotes ? `"${value.replace(/"/g, '""')}"` : value;
}

/**
 * Pure hit predicate for drag-to-resize (issue #567): given a pointer
 * coordinate `pt` (in the header-strip's CSS-px axis — already RTL-un-mirrored
 * by the caller) and the candidate band trailing edges `edges`, return the band
 * index whose edge is within `grabPx` of `pt`, or `null` if none qualifies.
 *
 * `edges` is the candidate list the caller builds — for the band the pointer is
 * over (`hit`) Excel lets you resize the band whose *trailing* border you grab,
 * so the caller passes both `hit - 1` and `hit` (the neighbour-to-the-far-side
 * and the band itself); the first edge within the grab zone wins, in the order
 * given. An edge that sits at or under the header strip (`edge <= headerExtent`,
 * i.e. scrolled behind the frozen corner) is rejected — you can't grab a border
 * hidden under the header. Kept pure (no DOM, no `this`) so the off-by-one
 * geometry — exact-on-edge, within-grab, just-outside, `[hit-1, hit]` neighbour
 * selection, header rejection — is unit-testable. {@link XlsxViewer.getResizeTarget}
 * does the DOM/geometry and calls this.
 */
export function resizeHitIndex(
  pt: number,
  edges: { index: number; edge: number }[],
  grabPx: number,
  headerExtent: number,
): number | null {
  for (const { index, edge } of edges) {
    if (edge <= headerExtent) continue; // scrolled behind the header strip
    if (Math.abs(pt - edge) <= grabPx) return index;
  }
  return null;
}

/**
 * Derive the selection rectangle's `border` and `background` CSS from a single
 * accent color: the border is the color verbatim and the fill is the same color
 * at 8% opacity via `color-mix`, so any CSS color string (`#rgb`, `rgb(...)`,
 * named) yields a matching translucent fill without the caller computing an
 * rgba. For the default `#1a73e8` this reproduces the historical
 * `rgba(26,115,232,0.08)` fill.
 */
export function selectionOverlayStyle(color: string): { border: string; background: string } {
  return {
    border: `2px solid ${color}`,
    background: `color-mix(in srgb, ${color} 8%, transparent)`,
  };
}

interface SelectionOverlayRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly top: boolean;
  readonly right: boolean;
  readonly bottom: boolean;
  readonly left: boolean;
}

interface SelectionBoundarySegment {
  readonly axis: 'h' | 'v';
  readonly fixed: number;
  readonly start: number;
  readonly end: number;
}

/**
 * Build the single-Area outline from its visible frozen-pane fragments.
 * Splitting collinear edges at every endpoint emits coincident fragment edges
 * only once. Work is bounded by the visible fragment count, not sheet size.
 */
function selectionBoundaryPath(rects: readonly SelectionOverlayRect[]): string {
  const raw: SelectionBoundarySegment[] = [];
  for (const rect of rects) {
    const x2 = rect.x + rect.width;
    const y2 = rect.y + rect.height;
    if (rect.top) raw.push({ axis: 'h', fixed: rect.y, start: rect.x, end: x2 });
    if (rect.right) raw.push({ axis: 'v', fixed: x2, start: rect.y, end: y2 });
    if (rect.bottom) raw.push({ axis: 'h', fixed: y2, start: rect.x, end: x2 });
    if (rect.left) raw.push({ axis: 'v', fixed: rect.x, start: rect.y, end: y2 });
  }

  const groups = new Map<string, SelectionBoundarySegment[]>();
  for (const segment of raw) {
    const key = `${segment.axis}:${segment.fixed}`;
    const group = groups.get(key);
    if (group) group.push(segment);
    else groups.set(key, [segment]);
  }

  const commands: string[] = [];
  for (const segments of groups.values()) {
    const points = [...new Set(segments.flatMap(({ start, end }) => [start, end]))]
      .sort((a, b) => a - b);
    let runStart: number | null = null;
    let runEnd = 0;
    const flush = () => {
      if (runStart === null || runEnd <= runStart) return;
      const { axis, fixed } = segments[0];
      commands.push(axis === 'h'
        ? `M${runStart} ${fixed}H${runEnd}`
        : `M${fixed} ${runStart}V${runEnd}`);
      runStart = null;
    };
    for (let index = 0; index + 1 < points.length; index++) {
      const start = points[index];
      const end = points[index + 1];
      const covered = segments.some((segment) => segment.start < end && segment.end > start);
      if (covered && runStart !== null && start === runEnd) {
        runEnd = end;
      } else {
        flush();
        if (covered) {
          runStart = start;
          runEnd = end;
        }
      }
    }
    flush();
  }
  return commands.join('');
}

let selectionMaskSequence = 0;

const DEFAULT_FIND_HIGHLIGHT = 'color-mix(in srgb, #ffb300 8%, transparent)';
const DEFAULT_FIND_ACTIVE_HIGHLIGHT = 'color-mix(in srgb, #fb8c00 8%, transparent)';

/** Resolve an XLSX find box without altering a caller-provided CSS background. */
export function findHighlightOverlayStyle(
  active: boolean,
  colors: FindHighlightColors = {},
): { border: string; background: string } {
  const accent = active ? '#fb8c00' : '#ffb300';
  const custom = active ? colors.active : colors.match;
  const background = custom ?? (active ? DEFAULT_FIND_ACTIVE_HIGHLIGHT : DEFAULT_FIND_HIGHLIGHT);
  return { border: `2px solid ${custom ?? accent}`, background };
}

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
  /** Left gutter canvas: row group brackets + toggles (XL4). */
  private rowGutter!: HTMLCanvasElement;
  /** Top gutter canvas: column group brackets + toggles (XL4). */
  private colGutter!: HTMLCanvasElement;
  /** Top-left corner canvas: numbered level buttons (XL4). */
  private cornerGutter!: HTMLCanvasElement;
  /** Cached extents (unscaled CSS px) of the current sheet's gutters; both 0 for
   *  an outline-free sheet. `w` insets {@link canvasArea} from the left, `h` from
   *  the top. */
  private gutter = { w: 0, h: 0 };
  /** Per-axis outline layout (group brackets + toggles) for the current sheet,
   *  recomputed on sheet switch and after each collapse/expand. `null` axis ⇒ no
   *  outlining on that axis. */
  private rowOutline: OutlineLayout | null = null;
  private colOutline: OutlineLayout | null = null;
  private rowOutlineBands: BandOutline[] = [];
  private colOutlineBands: BandOutline[] = [];
  /** Original row heights / column widths stashed the first time a band is
   *  collapsed, so expanding restores a custom size rather than the default.
   *  Keyed by band index; per current worksheet (cleared on sheet switch). */
  private stashedRowHeights = new Map<number, number | undefined>();
  private stashedColWidths = new Map<number, number | undefined>();
  /**
   * Per-sheet cumulative record of every view-only size mutation (outline
   * collapse/expand, drag-to-resize #567), keyed by sheet index. Value = the
   * band's current model size, or `null` when the model has no entry (default
   * size). Serialized as {@link WireSizeOverrides} with every render so both
   * modes draw from a render-local projection matching this viewer, while the
   * workbook cache remains immutable for sibling viewers. Entries are updated
   * in place and never removed; the whole store resets with a new workbook.
   */
  private sizeOverrideStore = new Map<
    number,
    {
      rows: Map<number, number | null>;
      automaticRows: Map<number, number>;
      cols: Map<number, number | null>;
      revision: number;
      wire?: WireSizeOverrides;
    }
  >();
  private readonly projectionId = nextViewerProjectionId++;
  private canvasArea: HTMLDivElement;
  private scrollHost: HTMLDivElement;
  private spacer: HTMLDivElement;
  private readonly surface: CanvasSurface;
  private readonly overlayHost: SheetOverlayHost;
  /** Composite-viewer chrome. These fields are initialized only for the
   *  container-mounted workbook viewer; sheet mounts create no footer DOM. */
  private tabBar!: HTMLDivElement;
  private tabStrip!: HTMLDivElement;
  /** Direction-aware flex row inside the LTR scroll host. Keeping direction on
   *  this inner row avoids browser-specific negative scrollLeft semantics. */
  private tabList!: HTMLDivElement;
  private navPrev!: HTMLButtonElement;
  private navNext!: HTMLButtonElement;
  private tabs: HTMLButtonElement[] = [];
  /** Per-tab colors parallel to `tabs`, from `<sheetPr><tabColor>`. */
  private tabColors: (string | null)[] = [];
  private zoomSlider: HTMLInputElement | null = null;
  private zoomLabel: HTMLSpanElement | null = null;
  private currentSheet = 0;
  /** Atomically commits an asynchronously acquired worksheet with its index.
   * Incremented by every navigation and teardown so late acquisitions are no-ops. */
  private sheetRequestGeneration = 0;
  private fontBindingGeneration = 0;
  private fontBinding: Readonly<{ workbook: XlsxWorkbook; release: () => void }> | null = null;
  private _hiddenSheetMode: HiddenSheetMode;
  private currentWorksheet: Worksheet | null = null;
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
  private chromeColors: XlsxChromeColors = {};
  private chromeStyleObserver: MutationObserver | null = null;
  private chromeSchemeMedia: MediaQueryList | null = null;
  private chromeSchemeListener: (() => void) | null = null;
  /** Last offset delivered to onViewportChange. Keeping this in the shared
   *  engine prevents a programmatic scroll followed by the browser's native
   *  scroll event from producing duplicate notifications. */
  private _lastViewportNotification: XlsxViewportOffset | null = null;
  private get anchorCell(): CellAddress | null {
    return this.selectionController.anchor;
  }

  private get activeCell(): CellAddress | null {
    return this.selectionController.active;
  }

  private get selectionMode(): SheetSelectionMode {
    return this.selectionController.mode;
  }

  private get isSelecting(): boolean {
    return this.selectionController.dragging;
  }

  private get selectionPointerId(): number | null {
    return this.selectionController.draggingPointerId;
  }

  /** Claim drag-selection ownership and discard deferred gestures from any
   * other pointer that began before this drag. */
  private beginSelectionDrag(pointerId: number): void {
    if (this.pendingTap?.pointerId !== pointerId) this.pendingTap = null;
    if (this.pendingClick?.pointerId !== pointerId) this.pendingClick = null;
    this.selectionController.beginDrag(pointerId);
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
  private lastNotifiedSelectionState: XlsxSelectionState | null = null;
  private emittingSelectionChange = false;
  private pendingSelectionChange = false;
  private selectionNotificationScheduled = false;
  private selectionNotificationCount = 0;
  private selectionContextNotificationFrame: number | null = null;
  private selectionContextNotificationMicrotask = false;
  // SpreadsheetML permits explicit row/cell references to appear out of
  // coordinate order. Cache a canonical view once per immutable parsed model
  // so range extraction can use binary search without silently skipping such
  // cells on every subsequent context read.
  private readonly selectionContextRows = new WeakMap<Worksheet, readonly Row[]>();
  private readonly selectionContextCells = new WeakMap<Row, readonly Cell[]>();
  private elementContext: XlsxElementContext | null = null;
  private selectionOverlay: HTMLDivElement;
  /** IX2 — find-highlight overlay (matched-cell boxes). */
  private findOverlay!: HTMLDivElement;
  /** IX2 — find state (matches + active cursor). */
  private _find!: XlsxFindController;
  private keydownHandler: ((e: KeyboardEvent) => void) | null = null;
  // Deferred selection press: committed on pointerup only if the pointer
  // neither moved beyond the tap threshold nor caused a scroll. Used for
  // touch/pen (swipe-to-scroll must not change the cell) and for mouse
  // presses inside the overlay-scrollbar band (a thumb drag must not select
  // the cell underneath).
  private pendingTap:
    | { x: number; y: number; shiftKey: boolean; additiveKey: boolean; pointerId: number }
    | null = null;
  // IX1 — mouse press bookkeeping for hyperlink activation: the down position and
  // the cell under it. On pointerup, if the pointer did not move beyond the tap
  // slop (a genuine click, not a drag-select), a hyperlink on that cell is
  // dispatched. Touch/pen activate through the pendingTap path instead.
  private pendingClick: { x: number; y: number; pointerId: number; cell: CellAddress } | null = null;
  private pendingElementClick:
    | { x: number; y: number; pointerId: number; context: XlsxElementContext }
    | null = null;
  // In-flight column/row resize drag (issue #567). `originScaled` is the fixed
  // LTR edge the resized band grows from (left edge for a column, top for a row)
  // in canvasArea CSS px; `mdw` is captured once so the live px→model-unit
  // conversion is stable across the drag. A resize is a *view-only* adjustment:
  // it mutates the in-memory worksheet's colWidths/rowHeights, never the file.
  private resizeDrag:
    | { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number; pointerId: number }
    | null = null;
  /** Last captured drag-selection pointer, retained while edge scrolling runs. */
  private selectionAutoScrollPointer:
    | { clientX: number; clientY: number; pointerId: number }
    | null = null;
  private selectionAutoScrollFrame: number | null = null;
  private selectionAutoScrollLastTime: number | null = null;

  // ─── Comment hover popup (Excel-style note) ───────────────────────────────
  /** DOM overlay element that shows the hovered cell's comment. */
  private commentPopup: HTMLDivElement;
  /** `"row:col"` → comment for the current sheet, rebuilt on every showSheet. */
  private commentMap = new Map<string, XlsxComment>();
  /** IX1 — `"row:col"` → hyperlink for the current sheet, rebuilt on every
   *  showSheet. Keys mirror the renderer's `hyperlinkMap` (1-based row/col, the
   *  first cell of a hyperlink `ref` range per the parser), so a `getCellAt`
   *  {row,col} looks up directly. */
  private hyperlinkMap = new Map<string, Hyperlink>();
  /** `"row:col"` of the cell whose popup is currently shown (or pending), so a
   *  pointermove within the same cell doesn't restart the show timer. */
  private commentPopupKey: string | null = null;
  /** Pending show timer (see {@link COMMENT_POPUP_DELAY_MS}). */
  private commentPopupTimer: ReturnType<typeof setTimeout> | null = null;
  private commentPopupCell: CellAddress | null = null;
  private commentPopupPositionScheduled = false;
  private commentPopupResizeObserver: ResizeObserver | null = null;
  private commentUi: XlsxCommentUiRuntime | null = null;
  private commentPopupRenderGeneration = 0;

  // ─── List data-validation dropdown panel (display-only) ───────────────────
  /** DOM overlay listing a list-validated cell's allowed values. Lives in
   *  canvasArea above the scrollHost; unlike the comment popup this is a click
   *  target (`pointer-events:auto`). Read-only: hovering an item highlights it
   *  but selecting does NOT change the cell. */
  private validationPanel: HTMLDivElement;
  /** `"row:col"` of the cell whose panel is pending or open, or null. Claiming
   *  the key before async range resolution lets a re-click cancel the request. */
  private validationPanelKey: string | null = null;
  private validationRequestGeneration = 0;
  /** Screen rect (canvasArea CSS px) of the dropdown arrow button last drawn by
   *  {@link maybeDrawValidationDropdown}, so pointerdown can hit-test it. Null
   *  when no arrow is currently visible. */
  private validationArrowRect: { x: number; y: number; w: number; h: number } | null = null;
  /** Document-level pointerdown listener that closes the panel on an outside
   *  click; installed only while the panel is open. */
  private validationOutsideHandler: ((e: PointerEvent) => void) | null = null;

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

    // Outline gutter canvases. Absolutely positioned inside gridRegion; sized /
    // shown per sheet in `layoutGutters`. `pointer-events:auto` on the gutters so
    // +/- toggles and level buttons are clickable; they are painted on the main
    // thread even in worker mode (cheap chrome, independent of the grid bitmap).
    const gutterStyle =
      `position:absolute;top:0;left:0;z-index:3;display:none;` +
      `background:var(--ooxml-xlsx-chrome-background,#f5f5f5);`;
    this.cornerGutter = this.hostDocument.createElement('canvas');
    this.cornerGutter.style.cssText = gutterStyle;
    this.cornerGutter.setAttribute('data-xlsx-outline', 'corner');
    this.colGutter = this.hostDocument.createElement('canvas');
    this.colGutter.style.cssText = gutterStyle;
    this.colGutter.setAttribute('data-xlsx-outline', 'col');
    this.rowGutter = this.hostDocument.createElement('canvas');
    this.rowGutter.style.cssText = gutterStyle;
    this.rowGutter.setAttribute('data-xlsx-outline', 'row');

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
    this.selectionOverlay = this.overlayHost.selection;
    this.findOverlay = this.overlayHost.find;
    this.commentPopup = this.overlayHost.comment;
    const ResizeObserverClass = this.hostDocument.defaultView?.ResizeObserver ??
      globalThis.ResizeObserver;
    if (ResizeObserverClass) {
      this.commentPopupResizeObserver = new ResizeObserverClass(() => {
        this.scheduleCommentPopupPosition();
      });
      this.commentPopupResizeObserver.observe(this.commentPopup);
    }
    this.validationPanel = this.overlayHost.validation;
    // Inject the shared viewer stylesheet once per module (idempotent). Both
    // mounts use it; the composite footer also hides its tab-strip scrollbar.
    ensureViewerStyleInjected(this.hostDocument);

    if (mount.kind === 'composite') {
      this.tabBar = this.hostDocument.createElement('div');
      this.tabBar.style.cssText =
        `display:flex;align-items:flex-end;height:${TAB_BAR_H}px;flex-shrink:0;` +
        `background:var(--ooxml-xlsx-chrome-background,#f0f0f0);` +
        `border-top:1px solid var(--ooxml-xlsx-chrome-border,#c8ccd0);`;

      // Excel-style scroll buttons. They scroll the tab strip; they do NOT change
      // the active sheet. Disabled (greyed) at the ends / when there is no overflow.
      this.navPrev = this.makeNavButton('◀', 'Scroll tabs left', () => this.scrollTabs(-1));
      this.navNext = this.makeNavButton('▶', 'Scroll tabs right', () => this.scrollTabs(1));
      this.navPrev.dataset.xlsxTabNav = 'prev';
      this.navNext.dataset.xlsxTabNav = 'next';

      // Keep the two-button footer control at the row-header width from the 100%
      // view. It is viewer chrome, so workbook zoom must not resize or shift it.
      const navGroup = this.hostDocument.createElement('div');
      navGroup.style.cssText =
        `display:flex;flex-shrink:0;width:${TAB_NAV_W}px;height:100%;`;
      navGroup.appendChild(this.navPrev);
      navGroup.appendChild(this.navNext);

      // The scrollable strip that actually holds the sheet tabs. position:relative
      // so each tab's offsetLeft is measured against the strip's scroll content.
      this.tabStrip = this.hostDocument.createElement('div');
      // Keep the scroll host itself LTR so scrollLeft is consistently 0..max in
      // every browser. The inner tabList owns visual LTR/RTL ordering.
      this.tabStrip.style.cssText =
        `position:relative;display:block;flex:1;min-width:0;height:100%;` +
        `margin-left:${TAB_GAP}px;overflow-x:auto;overflow-y:hidden;scrollbar-width:none;`;
      this.tabStrip.classList.add('xlsx-tab-strip');
      this.tabStrip.addEventListener('scroll', () => this.updateNavButtons());

      // width:max-content preserves overflow scrolling; min-width:100% makes a
      // short RTL tab row fill the strip so row-reverse can right-align it.
      this.tabList = this.hostDocument.createElement('div');
      this.tabList.style.cssText =
        `display:flex;align-items:flex-end;height:100%;` +
        `gap:${TAB_GAP}px;box-sizing:border-box;`;
      this.tabList.style.width = 'max-content';
      this.tabList.style.minWidth = '100%';
      this.tabStrip.appendChild(this.tabList);

      this.tabBar.appendChild(navGroup);
      this.tabBar.appendChild(this.tabStrip);
      if (this.opts.showZoomSlider !== false) {
        this.tabBar.appendChild(this.buildZoomControl());
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
    if (mount.kind === 'composite') this.wrapper.appendChild(this.tabBar);
    container.appendChild(this.wrapper);
    this.installChromeThemeRefresh();

    // Gutter click handling (XL4): +/- toggles and the numbered level banks
    // (each in its own gutter's header strip; the corner is inert background).
    // Registered once; no-op when a sheet has no gutter (extents 0 ⇒ hidden).
    this.rowGutter.addEventListener('pointerdown', (e) => this.onGutterPointerDown(e, 'row'));
    this.colGutter.addEventListener('pointerdown', (e) => this.onGutterPointerDown(e, 'col'));

    if (this._nativeScrollbars) this.surface.on('scroll', () => {
      // Any scroll cancels a deferred tap: the press that started it was a
      // scrollbar-thumb drag (overlay scrollbars) or a touch swipe, not a
      // cell click.
      this.pendingTap = null;
      this.pendingElementClick = null;
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
      this.updateNavButtons();
    });
    resizeObserver.observe(this.gridRegion);
    this.resizeObserver = resizeObserver;

    this.setupSelectionEvents();

    this._find = new XlsxFindController(
      () => this.sheetCount,
      (sheet) => this.wb?.sheetNames[sheet] ?? '',
      (sheet) => this._collectSheetCells(sheet),
    );

    if (borrowedWorkbook) {
      this.acquisition.install(borrowedWorkbook, false);
      if (this._mountKind === 'composite') {
        this.activateWorkbook(borrowedWorkbook).catch((error) => this._reportRenderError(error));
      }
    }

  }

  /**
   * Re-read the CSS custom properties that affect Canvas-painted Viewer chrome.
   * DOM chrome follows inherited CSS variables without help; row/column headers
   * and outline gutters need an explicit repaint because their colors are baked
   * into pixels.
   */
  private refreshChromeTheme(): void {
    if (this._destroyed) return;
    const getComputedStyle = this.hostWindow.getComputedStyle?.bind(this.hostWindow);
    if (!getComputedStyle) return;
    const computed = getComputedStyle(this.wrapper);
    const next: Record<string, string> = {};
    for (const [key, property] of Object.entries(XLSX_CHROME_COLOR_PROPERTIES)) {
      const value = computed.getPropertyValue(property).trim();
      if (value) next[key] = value;
    }
    const nextColors = next as XlsxChromeColors;
    if (sameChromeColors(this.chromeColors, nextColors)) return;
    this.chromeColors = nextColors;
    this.renderGutters();
    this.scheduleRender();
  }

  /** Observe the ordinary ways an application changes theme state. */
  private installChromeThemeRefresh(): void {
    this.refreshChromeTheme();

    const MutationObserverClass = this.hostWindow.MutationObserver ?? globalThis.MutationObserver;
    if (MutationObserverClass) {
      this.chromeStyleObserver = new MutationObserverClass(() => this.refreshChromeTheme());
      for (let target: HTMLElement | null = this.container; target; target = target.parentElement) {
        this.chromeStyleObserver.observe(target, {
          attributes: true,
          attributeFilter: ['class', 'style', 'data-theme'],
        });
      }
    }

    const media = this.hostWindow.matchMedia?.('(prefers-color-scheme: dark)') ?? null;
    if (media) {
      const listener = () => this.refreshChromeTheme();
      media.addEventListener?.('change', listener);
      this.chromeSchemeMedia = media;
      this.chromeSchemeListener = listener;
    }
  }

  /** Every non-empty cell of a sheet with its rendered display text (IX2 find
   *  source). Reads the parsed worksheet model directly — no render — so search
   *  covers the whole sheet, not just the on-screen viewport. */
  private async _collectSheetCells(sheet: number): Promise<FindCell[]> {
    const wb = this.wb;
    if (!wb) return [];
    const ws = await wb.getWorksheet(sheet);
    const cells: FindCell[] = [];
    for (const row of ws.rows) {
      for (const cell of row.cells) {
        const text = wb.cellText(ws, cell);
        if (text !== '') cells.push({ row: cell.row, col: cell.col, text });
      }
    }
    return cells;
  }

  /**
   * Load an XLSX from URL or ArrayBuffer and render the first sheet.
   *
   * Parse, load, and initial-render failures always reject this Promise.
   * `onError` is reserved for later Viewer-managed work that has no directly
   * awaitable method result, so one failure is never delivered twice.
   */
  async load(source: string | ArrayBuffer): Promise<void> {
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
      const wb = await this.acquisition.replace(() => XlsxWorkbook.load(source, {
          password: this.opts.password,
          useGoogleFonts: this.opts.useGoogleFonts,
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
          mode: this._mode,
        }), () => {
          // Claim every async-operation generation before closing the old
          // workbook. Rejections caused by its worker termination are stale
          // completion, not errors belonging to the new workbook.
          this.sheetRequestGeneration++;
          this.renderDispatcher.begin();
          this._find.invalidate();
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
    this._find.invalidate();
    this.sizeOverrideStore.clear();
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
    const generation = ++this.sheetRequestGeneration;
    const workbook = this.workbook;
    let worksheet: Worksheet;
    let sourceWorksheet: Worksheet;
    try {
      if (!await this.ensureHostFonts(workbook)) return;
      sourceWorksheet = await workbook.getWorksheet(index);
      worksheet = this.sheetViews.get(index) ?? this.createVisibleSheetView(sourceWorksheet);
      const prepareRowHeights = workbook[prepareXlsxViewerRowHeights];
      if (typeof prepareRowHeights === 'function') {
        const measureCanvas = this.hostDocument.createElement('canvas');
        const measureCtx = measureCanvas.getContext('2d');
        if (measureCtx) prepareRowHeights.call(workbook, worksheet, measureCtx);
      }
      this.syncAutomaticRowOverrides(index, worksheet);
      this.sheetViews.set(index, worksheet);
    } catch (error) {
      if (!this.isCurrentSheetRequest(generation, workbook)) return;
      throw error;
    }
    if (!this.isCurrentSheetRequest(generation, workbook)) return;

    this.currentSheet = index;
    this.currentWorksheet = worksheet;
    this.currentSourceComments = sourceWorksheet.comments ?? [];
    if (this.opts.comments !== false && this.currentSourceComments.length > 0) {
      void this.loadCommentUi().catch((error) => this._reportRenderError(error));
    }
    this.sourceCommentMap = this.createCommentMap(this.currentSourceComments);
    this.setElementContext(null);
    this.pendingElementClick = null;
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
    await this.renderCurrentSheet();
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

  // ─── Outline gutter (XL4: row/column grouping) ────────────────────────────

  /** Recompute the per-axis outline layout for `ws` and cache the band lists.
   *  Both axes are `null` (gutters collapse to 0) when the sheet has no
   *  outlining, so an outline-free sheet is untouched. */
  private buildOutline(ws: Worksheet): void {
    this.stashedRowHeights.clear();
    this.stashedColWidths.clear();
    this.rowOutlineBands = rowBands(ws);
    this.colOutlineBands = colBands(ws);
    const rowLayout = buildOutlineLayout(this.rowOutlineBands, summaryAfterFor(ws, 'row'));
    const colLayout = buildOutlineLayout(this.colOutlineBands, summaryAfterFor(ws, 'col'));
    this.rowOutline = rowLayout.maxLevel > 0 ? rowLayout : null;
    this.colOutline = colLayout.maxLevel > 0 ? colLayout : null;
  }

  /** Size and place the three gutter canvases (corner / col / row) from the
   *  current outline, and inset {@link canvasArea} by the gutter extents. When
   *  neither axis is grouped both extents are 0 and canvasArea covers the whole
   *  region — pixel-identical to a viewer built before XL4. */
  private layoutGutters(): void {
    const cs = this.viewport.scale;
    const gw = this.rowOutline ? Math.round(gutterExtentPx(this.rowOutline.maxLevel) * cs) : 0;
    const gh = this.colOutline ? Math.round(gutterExtentPx(this.colOutline.maxLevel) * cs) : 0;
    this.gutter = { w: gw, h: gh };

    // Attach the gutter canvases only while an outline exists; detach them
    // entirely for outline-free sheets. A hidden-but-attached canvas is NOT
    // neutral — DOM consumers that count/index `<canvas>` elements (e.g. the
    // layouts smoke's `page.locator('canvas').count()`) see it — so element
    // parity with the pre-outline viewer requires absence, not `display:none`.
    // The elements (and their pointer listeners) are constructed once and
    // survive detach/reattach across sheet switches.
    if (gw > 0 || gh > 0) {
      if (!this.colGutter.parentElement) {
        this.gridRegion.appendChild(this.colGutter);
        this.gridRegion.appendChild(this.rowGutter);
        this.gridRegion.appendChild(this.cornerGutter);
      }
    } else {
      this.colGutter.remove();
      this.rowGutter.remove();
      this.cornerGutter.remove();
    }

    // Inset canvasArea so the grid (and every geometry read that keys off its
    // client rect) starts after the gutters.
    this.canvasArea.style.left = `${gw}px`;
    this.canvasArea.style.top = `${gh}px`;

    const show = (el: HTMLCanvasElement, x: number, y: number, w: number, h: number) => {
      if (w <= 0 || h <= 0) { el.style.display = 'none'; return; }
      el.style.display = 'block';
      el.style.left = `${x}px`;
      el.style.top = `${y}px`;
      el.style.width = `${w}px`;
      el.style.height = `${h}px`;
    };
    const regionW = this.gridRegion.clientWidth;
    const regionH = this.gridRegion.clientHeight;
    // Corner holds the numbered level buttons; only meaningful where both a
    // horizontal and vertical gutter exist, but we always paint it to cover the
    // intersection so the two strips meet cleanly.
    show(this.cornerGutter, 0, 0, gw, gh);
    show(this.colGutter, gw, 0, Math.max(0, regionW - gw), gh);
    show(this.rowGutter, 0, gh, gw, Math.max(0, regionH - gh));
  }

  /** Paint all visible gutter strips for the current scroll offset. Called at the
   *  end of every grid render so the brackets track scroll / zoom exactly. */
  private renderGutters(): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    if (this.gutter.h > 0 && this.colOutline) this.paintAxisGutter('col');
    if (this.gutter.w > 0 && this.rowOutline) this.paintAxisGutter('row');
    if (this.gutter.w > 0 || this.gutter.h > 0) this.paintCornerGutter();
  }

  /** Draw one axis's group brackets and +/- toggles into its gutter canvas,
   *  aligned to the on-screen band positions via {@link getCellRect}. */
  private paintAxisGutter(axis: OutlineAxis): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const cs = this.viewport.scale;
    const isRow = axis === 'row';
    const canvas = isRow ? this.rowGutter : this.colGutter;
    const layout = isRow ? this.rowOutline : this.colOutline;
    if (!layout) return;
    const cssW = parseFloat(canvas.style.width) || 0;
    const cssH = parseFloat(canvas.style.height) || 0;
    if (cssW <= 0 || cssH <= 0) return;
    // Backing-store size at DPR; CSS size stays as laid out.
    const dpr = this.surface.sizeCanvas(canvas, cssW, cssH);
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);
    ctx.fillStyle = this.chromeColors.background ?? '#f5f5f5';
    ctx.fillRect(0, 0, cssW, cssH);

    const lanePx = OUTLINE_LANE_PX * cs;
    // The gutter canvas's cross-axis origin (0) sits at the grid's cell-area
    // origin: for the row gutter, y=0 aligns with the top of the row header +
    // gutter; getCellRect returns coordinates in canvasArea space, which is
    // offset from the gutter canvas by exactly `gutter.h` (col gutter is above).
    // The gutter canvas top is at gridRegion y = gutter.h, and canvasArea top is
    // also at gutter.h — so a band's canvasArea-space y maps 1:1 to gutter-canvas
    // y. Likewise x for the col gutter (offset by gutter.w).
    ctx.strokeStyle = this.chromeColors.border ?? '#808080';
    ctx.lineWidth = 1;
    ctx.fillStyle = this.chromeColors.text ?? '#404040';

    // Outline gutter geometry participates in the same header/frozen-pane split
    // as the worksheet canvas. Clip every logical run to its own pane so a
    // scrolled detail rail cannot leak upward into the column-letter header or
    // through the frozen-row boundary (and mirror the equivalent rule for RTL
    // frozen columns).
    const geometry = getGridGeometryForWorksheet(ws);
    const effective = geometry.effectiveFrozenBands({
      scale: cs,
      width: this.canvasArea.clientWidth,
      height: this.canvasArea.clientHeight,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
      rows: ws.freezeRows ?? 0,
      cols: ws.freezeCols ?? 0,
    });
    const axes = geometry.axesAtScale(cs);
    const frozenBandCount = isRow ? effective.rows : effective.cols;
    const frozenExtent = isRow
      ? axes.row.offsetOf(effective.rows + 1)
      : axes.col.offsetOf(effective.cols + 1);
    const headerExtent = (isRow ? HEADER_H : HEADER_W) * cs;
    const paneClip = (start: number, end: number) => outlinePaneClipRect(
      axis,
      start,
      end,
      frozenBandCount,
      headerExtent,
      frozenExtent,
      cssW,
      cssH,
      !isRow && ws.rightToLeft === true,
    );
    const clipContext = (start: number, end: number): boolean => {
      const clip = paneClip(start, end);
      if (clip.w <= 0 || clip.h <= 0) return false;
      ctx.save();
      ctx.beginPath();
      ctx.rect(clip.x, clip.y, clip.w, clip.h);
      ctx.clip();
      return true;
    };

    for (const g of layout.groups) {
      // Lane index for this level: lane 0 is the outermost (level 1). Buttons and
      // the outermost bracket sit nearest the grid edge? Excel draws level 1 in
      // the lane FARTHEST from the grid, deeper levels closer. We place level L in
      // lane (L-1) counted from the sheet-far edge.
      const laneFromFar = g.level - 1;
      const laneCenterCross = (laneFromFar + 0.5) * lanePx;

      // Detail run extent along the band axis, from on-screen cell rects.
      const startRect = isRow ? this._cellRect(g.start, 1) : this._cellRect(1, g.start);
      const endRect = isRow ? this._cellRect(g.end, 1) : this._cellRect(1, g.end);
      if (!startRect || !endRect) continue;
      const a = isRow ? startRect.y : this.screenX(startRect.x, startRect.w);
      const b = isRow ? endRect.y + endRect.h : this.screenX(endRect.x, endRect.w) + endRect.w;
      const runStart = Math.min(a, b);
      const runEnd = Math.max(a, b);

      // A collapsed group's detail run is hidden (zero visible extent) — Excel
      // draws only the +/- toggle, no bracket. Skip the bracket when the run has
      // negligible length.
      if (!g.collapsed && runEnd - runStart > 1) {
        if (clipContext(g.start, g.end)) {
          ctx.beginPath();
          for (const segment of outlineBracketSegments(axis, laneCenterCross, a, b, lanePx)) {
            ctx.moveTo(segment.x1, segment.y1);
            ctx.lineTo(segment.x2, segment.y2);
          }
          ctx.stroke();
          ctx.restore();
        }
      }

      // +/- toggle box on the summary band.
      if (g.summary != null) {
        const sRect = isRow ? this._cellRect(g.summary, 1) : this._cellRect(1, g.summary);
        if (sRect) {
          const along = isRow
            ? sRect.y + sRect.h / 2
            : this.screenX(sRect.x, sRect.w) + sRect.w / 2;
          if (clipContext(g.summary, g.summary)) {
            this.drawToggleBox(ctx, isRow ? laneCenterCross : along, isRow ? along : laneCenterCross, g.collapsed, cs);
            ctx.restore();
          }
        }
      }
    }

    // Numbered level buttons (1..maxLevel+1), one per lane, in this gutter's
    // header strip: the row bank sits beside the column-letter header (the
    // gutter's top HEADER_H band — no bracket ever draws there because band
    // y-coordinates start at the header edge), the column bank above the
    // row-number header (leftmost HEADER_W band). Placing each bank in its own
    // gutter (Excel's layout) keeps the two banks from ever sharing a cell —
    // the old corner placement collided at the shared bottom-right lane and
    // made the row expand-all button unreachable.
    const bankCross = isRow ? (HEADER_H * cs) / 2 : (HEADER_W * cs) / 2;
    for (let l = 1; l <= layout.maxLevel + 1; l++) {
      const buttonCenter = outlineLevelButtonCenterPx(l) * cs;
      if (buttonCenter + (OUTLINE_BUTTON_PX * cs) / 2 > (isRow ? cssW : cssH) + 0.5) break;
      this.drawLevelButton(
        ctx,
        isRow ? buttonCenter : bankCross,
        isRow ? bankCross : buttonCenter,
        String(l),
        cs,
      );
    }

    // Paint the pane separator last so it visibly cuts the outline rail at the
    // same coordinate as the main grid's separator. This also extends the line
    // through the outline gutter, making the frozen-row boundary continuous
    // from the gutter through the row-number header and cells.
    if (frozenBandCount > 0) {
      const divider = isRow
        ? headerExtent + frozenExtent
        : ws.rightToLeft === true
          ? cssW - headerExtent - frozenExtent
          : headerExtent + frozenExtent;
      ctx.save();
      ctx.strokeStyle = this.chromeColors.border ?? '#7a7a7a';
      ctx.lineWidth = 0.5;
      ctx.beginPath();
      if (isRow) {
        ctx.moveTo(0, divider);
        ctx.lineTo(cssW, divider);
      } else {
        ctx.moveTo(divider, 0);
        ctx.lineTo(divider, cssH);
      }
      ctx.stroke();
      ctx.restore();
    }
  }

  /** Draw a small square +/- toggle centered at (cx, cy) in gutter-canvas CSS px. */
  private drawToggleBox(
    ctx: CanvasRenderingContext2D,
    cx: number,
    cy: number,
    collapsed: boolean,
    cs: number,
  ): void {
    const s = Math.round(9 * cs);
    const x = Math.round(cx - s / 2);
    const y = Math.round(cy - s / 2);
    ctx.save();
    ctx.fillStyle = this.chromeColors.surface ?? '#ffffff';
    ctx.strokeStyle = this.chromeColors.border ?? '#808080';
    ctx.lineWidth = 1;
    ctx.fillRect(x + 0.5, y + 0.5, s, s);
    ctx.strokeRect(x + 0.5, y + 0.5, s, s);
    ctx.strokeStyle = this.chromeColors.text ?? '#404040';
    ctx.beginPath();
    // horizontal stroke (present for both + and -)
    ctx.moveTo(x + 2.5, y + s / 2 + 0.5);
    ctx.lineTo(x + s - 1.5, y + s / 2 + 0.5);
    if (collapsed) {
      // vertical stroke makes it a "+"
      ctx.moveTo(x + s / 2 + 0.5, y + 2.5);
      ctx.lineTo(x + s / 2 + 0.5, y + s - 1.5);
    }
    ctx.stroke();
    ctx.restore();
  }

  /** Draw one numbered level button centered at (cx, cy) in gutter-canvas CSS
   *  px. Shared by the row bank (in the row gutter's top strip) and the column
   *  bank (in the column gutter's left strip). */
  private drawLevelButton(
    ctx: CanvasRenderingContext2D,
    cx: number,
    cy: number,
    label: string,
    cs: number,
  ): void {
    const s = Math.round(OUTLINE_BUTTON_PX * cs);
    const x = Math.round(cx - s / 2);
    const y = Math.round(cy - s / 2);
    ctx.save();
    ctx.font = `${Math.round(9 * cs)}px sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillStyle = this.chromeColors.surface ?? '#ffffff';
    ctx.strokeStyle = this.chromeColors.border ?? '#808080';
    ctx.lineWidth = 1;
    ctx.fillRect(x + 0.5, y + 0.5, s, s);
    ctx.strokeRect(x + 0.5, y + 0.5, s, s);
    ctx.fillStyle = this.chromeColors.text ?? '#404040';
    ctx.fillText(label, cx, cy + 0.5);
    ctx.restore();
  }

  /** Paint the corner (intersection of the two gutters) as plain background.
   *  The numbered level banks live in each axis gutter's own header strip
   *  (see paintAxisGutter), so the corner carries no interactive content. */
  private paintCornerGutter(): void {
    const canvas = this.cornerGutter;
    const cssW = parseFloat(canvas.style.width) || 0;
    const cssH = parseFloat(canvas.style.height) || 0;
    if (cssW <= 0 || cssH <= 0) { return; }
    const dpr = this.surface.sizeCanvas(canvas, cssW, cssH);
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);
    ctx.fillStyle = this.chromeColors.background ?? '#f5f5f5';
    ctx.fillRect(0, 0, cssW, cssH);
  }

  /** Handle a click in a row/col gutter: hit-test the +/- toggles and toggle the
   *  matching group's collapse state. */
  private onGutterPointerDown(e: PointerEvent, axis: OutlineAxis): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const isRow = axis === 'row';
    const layout = isRow ? this.rowOutline : this.colOutline;
    if (!layout) return;
    const canvas = isRow ? this.rowGutter : this.colGutter;
    const rect = canvas.getBoundingClientRect();
    const px = e.clientX - rect.left;
    const py = e.clientY - rect.top;
    const cs = this.viewport.scale;
    const lanePx = OUTLINE_LANE_PX * cs;
    const hitR = 7 * cs; // generous grab radius around a +/- button center

    // Numbered level bank first: it lives in this gutter's header strip (row
    // bank beside the column-letter header, column bank above the row-number
    // header — mirrors paintAxisGutter), where no +/- toggle can be.
    const bankCross = isRow ? (HEADER_H * cs) / 2 : (HEADER_W * cs) / 2;
    const inBankStrip = (isRow ? py : px) <= (isRow ? HEADER_H : HEADER_W) * cs;
    if (inBankStrip) {
      for (let l = 1; l <= layout.maxLevel + 1; l++) {
        const buttonCenter = outlineLevelButtonCenterPx(l) * cs;
        const cx = isRow ? buttonCenter : bankCross;
        const cy = isRow ? bankCross : buttonCenter;
        const buttonHitR = (OUTLINE_BUTTON_PX * cs) / 2;
        if (Math.abs(px - cx) <= buttonHitR && Math.abs(py - cy) <= buttonHitR) {
          e.preventDefault();
          this.applyLevelButton(l, axis);
          return;
        }
      }
      return; // header strip carries no toggles — don't fall through
    }

    for (const g of layout.groups) {
      if (g.summary == null) continue;
      const laneCenterCross = (g.level - 1 + 0.5) * lanePx;
      const sRect = isRow ? this._cellRect(g.summary, 1) : this._cellRect(1, g.summary);
      if (!sRect) continue;
      const along = isRow
        ? sRect.y + sRect.h / 2
        : this.screenX(sRect.x, sRect.w) + sRect.w / 2;
      const cx = isRow ? laneCenterCross : along;
      const cy = isRow ? along : laneCenterCross;
      if (Math.abs(px - cx) <= hitR && Math.abs(py - cy) <= hitR) {
        e.preventDefault();
        this.applyGroupToggle(g, axis);
        return;
      }
    }
  }

  /** Flip a single group's collapse state in the in-memory model, then rebuild
   *  the outline + repaint. View-only: the file is never written. */
  private applyGroupToggle(group: OutlineGroup, axis: OutlineAxis): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const bands = axis === 'row' ? this.rowOutlineBands : this.colOutlineBands;
    const { hide, show, nowCollapsed } = toggleGroupHidden(group, bands);
    for (const i of hide) this.setBandHidden(axis, i, true);
    for (const i of show) this.setBandHidden(axis, i, false);
    // Reflect the new collapsed state on the summary band so the next toggle
    // reads the correct direction and the +/- glyph flips.
    if (group.summary != null) this.setBandCollapsed(axis, group.summary, nowCollapsed);
    // Collapsing removes the detail bands before the summary. Anchor that
    // surviving summary band at the viewport start after geometry has been
    // rebuilt; otherwise the browser clamps the shortened scroll extent and
    // leaves an unrelated partial row at the top.
    this.afterOutlineMutation(
      ws,
      nowCollapsed && group.summary != null ? { axis, summary: group.summary } : undefined,
    );
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

  /** Collapse/expand the whole sheet to `level` on one axis. */
  private applyLevelButton(level: number, axis: OutlineAxis): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const bands = axis === 'row' ? this.rowOutlineBands : this.colOutlineBands;
    const { hide, show } = levelButtonHidden(bands, level);
    for (const i of hide) this.setBandHidden(axis, i, true);
    for (const i of show) this.setBandHidden(axis, i, false);
    // Update each group's summary-band collapsed flag from the new state: a group
    // at lane L is collapsed exactly when its detail (level >= L) is now hidden,
    // i.e. `L >= level`. Driving this off the layout's groups (rather than the
    // band list) also reaches level-0 summary bands, which are not in `bands`.
    const layout = axis === 'row' ? this.rowOutline : this.colOutline;
    if (layout) {
      for (const g of layout.groups) {
        if (g.summary != null) this.setBandCollapsed(axis, g.summary, g.level >= level);
      }
    }
    this.afterOutlineMutation(ws);
  }

  /** Set a row/column hidden by mapping to the size-0 encoding the axis/renderer
   *  already understand, stashing the original size so expand can restore it. */
  private setBandHidden(axis: OutlineAxis, index: number, hidden: boolean): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    if (axis === 'row') {
      if (hidden) {
        if (!this.stashedRowHeights.has(index)) {
          this.stashedRowHeights.set(index, ws.rowHeights[index]);
        }
        ws.rowHeights[index] = 0;
      } else {
        if (this.stashedRowHeights.has(index)) {
          const orig = this.stashedRowHeights.get(index);
          if (orig === undefined) delete ws.rowHeights[index];
          else ws.rowHeights[index] = orig;
          this.stashedRowHeights.delete(index);
        } else if (ws.rowHeights[index] === 0) {
          // Was hidden in the source file (height 0) with no stash — reveal at
          // the default height.
          delete ws.rowHeights[index];
        }
      }
    } else {
      if (hidden) {
        if (!this.stashedColWidths.has(index)) {
          this.stashedColWidths.set(index, ws.colWidths[index]);
        }
        ws.colWidths[index] = 0;
      } else {
        if (this.stashedColWidths.has(index)) {
          const orig = this.stashedColWidths.get(index);
          if (orig === undefined) delete ws.colWidths[index];
          else ws.colWidths[index] = orig;
          this.stashedColWidths.delete(index);
        } else if (ws.colWidths[index] === 0) {
          delete ws.colWidths[index];
        }
      }
    }
    // Mirror the post-mutation model value into the render override channel so
    // both modes can draw this viewer's projection without mutating the shared
    // workbook cache.
    this.recordSizeOverride(axis, index);
  }

  /** Record band `index`'s CURRENT model size (or `null` = no entry) in the
   *  per-sheet override store. Called after every view-only size mutation so
   *  both render modes receive this viewer's independent projection. */
  private recordSizeOverride(axis: OutlineAxis, index: number): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    let entry = this.sizeOverrideStore.get(this.currentSheet);
    if (!entry) {
      entry = { rows: new Map(), automaticRows: new Map(), cols: new Map(), revision: 0 };
      this.sizeOverrideStore.set(this.currentSheet, entry);
    }
    const target = axis === 'row' ? entry.rows : entry.cols;
    if (axis === 'row') entry.automaticRows.delete(index);
    const value = axis === 'row' ? ws.rowHeights[index] ?? null : ws.colWidths[index] ?? null;
    if (target.get(index) === value && target.has(index)) return;
    target.set(index, value);
    entry.revision++;
    entry.wire = undefined;
  }

  /** The current sheet's override store serialized for the wire, or undefined
   *  when nothing has been mutated (keeps the request payload unchanged). */
  private wireSizeOverrides(): Readonly<{
    overrides: WireSizeOverrides;
    revision: number;
  }> | undefined {
    const entry = this.sizeOverrideStore.get(this.currentSheet);
    if (!entry || (entry.rows.size === 0 && entry.automaticRows.size === 0 && entry.cols.size === 0)) {
      return undefined;
    }
    if (!entry.wire) {
      const wire: WireSizeOverrides = {};
      if (entry.rows.size > 0 || entry.automaticRows.size > 0) {
        wire.rows = Object.fromEntries([...entry.automaticRows, ...entry.rows]);
      }
      if (entry.cols.size > 0) wire.cols = Object.fromEntries(entry.cols);
      entry.wire = wire;
    }
    return { overrides: entry.wire, revision: entry.revision };
  }

  /** Mirror only display-derived heights into the worker projection channel.
   * Manual/authored sizes remain in `rows`, so a later column refit can replace
   * automatic values without reclassifying a user's row resize. */
  private syncAutomaticRowOverrides(sheetIndex: number, worksheet: Worksheet): void {
    const next = new Map(derivedAutoRowHeights(worksheet));
    let entry = this.sizeOverrideStore.get(sheetIndex);
    if (!entry && next.size === 0) return;
    if (!entry) {
      entry = { rows: new Map(), automaticRows: new Map(), cols: new Map(), revision: 0 };
      this.sizeOverrideStore.set(sheetIndex, entry);
    }
    entry.automaticRows = next;
    entry.revision++;
    entry.wire = undefined;
  }

  /** Update the `collapsed` flag on a band's model entry so the outline rebuild
   *  reflects the new state. */
  private setBandCollapsed(axis: OutlineAxis, index: number, collapsed: boolean): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    if (axis === 'row') {
      const row = ws.rows.find((r) => r.index === index);
      if (row) row.collapsed = collapsed;
    } else {
      ws.colCollapsed = ws.colCollapsed ?? {};
      if (collapsed) ws.colCollapsed[index] = true;
      else delete ws.colCollapsed[index];
    }
  }

  /** Shared tail of a gutter interaction: invalidate the axis cache, rebuild the
   *  outline (collapsed flags changed), refresh dependent geometry, re-render. */
  private afterOutlineMutation(
    ws: Worksheet,
    anchor?: { axis: OutlineAxis; summary: number },
  ): void {
    GridGeometry.invalidate(ws);
    this.buildOutlineLayoutOnly(ws);
    this.updateSpacerSize(ws);
    if (anchor) this.scrollOutlineSummaryToStart(anchor.axis, anchor.summary);
    this.updateSelectionOverlay();
    this.updateFindOverlay();
    this.scheduleRender();
    if (anchor) this.emitViewportChange();
  }

  /** Rebuild only the layout + band lists (not the stashes) after a collapse
   *  state change, so the +/- glyphs and bracket set stay in sync. */
  private buildOutlineLayoutOnly(ws: Worksheet): void {
    this.rowOutlineBands = rowBands(ws);
    this.colOutlineBands = colBands(ws);
    const rowLayout = buildOutlineLayout(this.rowOutlineBands, summaryAfterFor(ws, 'row'));
    const colLayout = buildOutlineLayout(this.colOutlineBands, summaryAfterFor(ws, 'col'));
    this.rowOutline = rowLayout.maxLevel > 0 ? rowLayout : null;
    this.colOutline = colLayout.maxLevel > 0 ? colLayout : null;
  }

  /** True when the current sheet's grid is laid out right-to-left. */
  private get isRtl(): boolean {
    return this.currentWorksheet?.rightToLeft === true;
  }

  /** Mirror the workbook footer around the sheet-tab strip for an RTL sheet.
   *  The DOM order remains navigation → tabs → zoom, which is also the
   *  logical reading order; `row-reverse` places that sequence right-to-left.
   *  Move the strip's leading gap with it so the spacing stays symmetric. */
  private updateFooterDirection(): void {
    if (this._mountKind !== 'composite') return;
    this.tabBar.style.flexDirection = this.isRtl ? 'row-reverse' : 'row';
    this.tabStrip.style.marginLeft = this.isRtl ? '0' : `${TAB_GAP}px`;
    this.tabStrip.style.marginRight = this.isRtl ? `${TAB_GAP}px` : '0';
    this.tabList.style.flexDirection = this.isRtl ? 'row-reverse' : 'row';
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
    if (this.elementContext) {
      return limitXlsxElementContext(this.elementContext, options.maxTextCharacters);
    }
    const worksheet = this.currentWorksheet;
    const selection = this.selectionState;
    if (!worksheet || !selection) return null;
    const requestedMax = options.maxCells ?? 1_000;
    if (!Number.isFinite(requestedMax) || requestedMax < 0) {
      throw new RangeError('maxCells must be a finite non-negative number.');
    }
    const maxCells = Math.min(MAX_SELECTION_CONTEXT_CELLS, Math.floor(requestedMax));
    const requestedTextMax = options.maxTextCharacters ?? DEFAULT_SELECTION_CONTEXT_TEXT_CHARACTERS;
    if (!Number.isFinite(requestedTextMax) || requestedTextMax < 0) {
      throw new RangeError('maxTextCharacters must be a finite non-negative number.');
    }
    const maxTextCharacters = Math.min(
      MAX_SELECTION_CONTEXT_TEXT_CHARACTERS,
      Math.floor(requestedTextMax),
    );
    let textCharacters = 0;
    let textTruncated = false;
    const boundedField = (input: string | readonly Readonly<{ text: string }>[]): string => {
      const parts: readonly (string | Readonly<{ text: string }>)[] =
        typeof input === 'string' ? [input] : input;
      const chunks: string[] = [];
      let fieldCharacters = 0;
      for (let index = 0; index < parts.length; index++) {
        const sourcePart = parts[index];
        const part = typeof sourcePart === 'string' ? sourcePart : sourcePart.text;
        const allowed = Math.max(0, Math.min(
          MAX_SELECTION_CONTEXT_FIELD_CHARACTERS - fieldCharacters,
          maxTextCharacters - textCharacters,
        ));
        const chunk = safeUtf16Prefix(part, allowed);
        chunks.push(chunk);
        fieldCharacters += chunk.length;
        textCharacters += chunk.length;
        if (chunk.length < part.length || index + 1 < parts.length && allowed === 0) {
          textTruncated = true;
          break;
        }
      }
      return chunks.join('');
    };
    const sheetSelected = selection.areas.some((area) => area.kind === 'sheet');
    const rowIntervals = mergeSelectionIntervals(selection.areas.flatMap((area) =>
      area.kind === 'rows' ? [{ first: area.firstRow, last: area.lastRow }] : []));
    const columnIntervals = mergeSelectionIntervals(selection.areas.flatMap((area) =>
      area.kind === 'columns'
        ? [{ first: area.firstColumn, last: area.lastColumn }]
        : []));
    const rectangles = selection.areas.flatMap((area) => area.kind === 'cells' ? [area] : []);
    const events = rectangles.flatMap((area, index) => [
      { row: area.top, index, active: true },
      { row: area.bottom + 1, index, active: false },
    ]).sort((a, b) => a.row - b.row || Number(a.active) - Number(b.active));
    const activeRectangles = new Set<number>();
    let eventIndex = 0;
    let activeColumnIntervals: SelectionInterval[] = [];
    const cells: XlsxSelectionContextCell[] = [];
    let cellsTruncated = false;
    const selectedRowIntervals = sheetSelected || columnIntervals.length > 0
      ? [{ first: 1, last: MAX_WORKSHEET_ROW }]
      : mergeSelectionIntervals([
          ...rowIntervals,
          ...rectangles.map((area) => ({ first: area.top, last: area.bottom })),
        ]);
    let rows = this.selectionContextRows.get(worksheet);
    if (!rows) {
      rows = orderedBy(worksheet.rows, (row) => row.index);
      this.selectionContextRows.set(worksheet, rows);
    }

    cellScan: for (const selectedRows of selectedRowIntervals) {
      let rowIndex = lowerBoundBy(rows, selectedRows.first, (row) => row.index);
      while (rowIndex < rows.length) {
        const row = rows[rowIndex++];
        if (row.index > selectedRows.last) break;
        let changed = false;
        while (eventIndex < events.length && events[eventIndex].row <= row.index) {
          const event = events[eventIndex++];
          if (event.active) activeRectangles.add(event.index);
          else activeRectangles.delete(event.index);
          changed = true;
        }
        if (changed) {
          activeColumnIntervals = mergeSelectionIntervals([...activeRectangles].map((index) => ({
            first: rectangles[index].left,
            last: rectangles[index].right,
          })));
        }
        const wholeRow = sheetSelected || intervalContains(rowIntervals, row.index);
        const selectedColumns = wholeRow
          ? [{ first: 1, last: MAX_WORKSHEET_COL }]
          : mergeSelectionIntervals([...columnIntervals, ...activeColumnIntervals]);
        for (const selectedColumnsInterval of selectedColumns) {
          let rowCells = this.selectionContextCells.get(row);
          if (!rowCells) {
            rowCells = orderedBy(row.cells, (cell) => cell.col);
            this.selectionContextCells.set(row, rowCells);
          }
          let cellIndex = lowerBoundBy(rowCells, selectedColumnsInterval.first, (cell) => cell.col);
          while (cellIndex < rowCells.length) {
            const cell = rowCells[cellIndex++];
            if (cell.col > selectedColumnsInterval.last) break;
            const raw = cell.value;
            const sourceComment = this.sourceCommentMap.get(`${cell.row}:${cell.col}`);
            if (raw.type === 'empty' && cell.formula === undefined && !sourceComment) continue;
            if (cells.length >= maxCells) { cellsTruncated = true; break cellScan; }
            const displayText = boundedField(this.wb?.cellText(worksheet, cell) ?? '');
            const value = raw.type === 'text'
              ? boundedField(raw.runs ?? raw.text)
              : raw.type === 'number'
                ? raw.number
                : raw.type === 'bool'
                  ? raw.bool
                  : raw.type === 'error'
                    ? boundedField(raw.error)
                  : null;
            const comment = sourceComment ? {
              root: {
                id: sourceComment.id,
                author: sourceComment.author,
                date: sourceComment.date,
                text: boundedField(sourceComment.rootText ?? sourceComment.text),
                status: sourceComment.resolved ? 'resolved' as const : 'active' as const,
              },
              replies: (sourceComment.replies ?? []).map((reply) => ({
                id: reply.id,
                author: reply.author,
                date: reply.date,
                text: boundedField(reply.text),
                status: reply.resolved ? 'resolved' as const : 'active' as const,
              })),
            } : undefined;
            cells.push({
              address: { row: cell.row, col: cell.col },
              displayText,
              valueType: raw.type,
              value,
              ...(cell.formula === undefined ? {} : { formula: boundedField(cell.formula) }),
              ...(comment === undefined ? {} : { comment }),
            });
            if (textTruncated) break cellScan;
          }
        }
      }
    }
    const truncationReasons: Array<'cells' | 'text'> = [];
    if (cellsTruncated) truncationReasons.push('cells');
    if (textTruncated) truncationReasons.push('text');
    return {
      format: 'xlsx',
      kind: 'range',
      sheetIndex: this.currentSheet,
      sheetName: worksheet.name,
      selection,
      coordinateCountUpperBound: selectionCoordinateCountUpperBound(selection),
      cells,
      truncated: truncationReasons.length > 0,
      truncationReasons,
      maxCells,
      textCharacters,
      maxTextCharacters,
    };
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
    if (!this.opts.onSelectionContextChange || this._destroyed ||
        this.selectionContextNotificationFrame !== null ||
        this.selectionContextNotificationMicrotask) return;
    const notify = () => {
      this.selectionContextNotificationFrame = null;
      this.selectionContextNotificationMicrotask = false;
      if (this._destroyed) return;
      const context = this.getSelectionContext({
        maxTextCharacters: DEFAULT_SELECTION_CONTEXT_NOTIFICATION_TEXT_CHARACTERS,
      });
      this.opts.onSelectionContextChange?.(context ? structuredClone(context) : null);
    };
    if (typeof this.hostWindow.requestAnimationFrame === 'function') {
      this.selectionContextNotificationFrame = this.hostWindow.requestAnimationFrame(notify);
    } else {
      this.selectionContextNotificationMicrotask = true;
      queueMicrotask(notify);
    }
  }

  private emitSelectionChange(): void {
    const state = this.selectionState;
    if (!selectionStatesEqual(state, this.lastNotifiedSelectionState)) {
      this.scheduleSelectionContextNotification();
    }
    if (this.emittingSelectionChange) {
      this.pendingSelectionChange = true;
      this.scheduleSelectionNotification();
      return;
    }
    this.pendingSelectionChange = false;
    if (selectionStatesEqual(state, this.lastNotifiedSelectionState)) {
      this.finishSelectionNotificationChain();
      return;
    }

    if (this.selectionNotificationCount >= MAX_REENTRANT_SELECTION_NOTIFICATIONS) {
      // A callback feedback cycle must not monopolize the main thread. The
      // canonical state remains authoritative; only notifications beyond the
      // documented per-chain safety limit are suppressed.
      this.lastNotifiedSelectionState = state ? structuredClone(state) : null;
      this.finishSelectionNotificationChain();
      return;
    }
    this.selectionNotificationCount++;
    this.lastNotifiedSelectionState = state ? structuredClone(state) : null;
    this.emittingSelectionChange = true;
    try {
      this.opts.onSelectionStateChange?.(state ? structuredClone(state) : null);
    } finally {
      this.emittingSelectionChange = false;
      if (this.pendingSelectionChange ||
          !selectionStatesEqual(this.selectionState, this.lastNotifiedSelectionState)) {
        this.scheduleSelectionNotification();
      } else {
        this.finishSelectionNotificationChain();
      }
    }
  }

  private scheduleSelectionNotification(): void {
    if (this.selectionNotificationScheduled || this._destroyed) return;
    this.selectionNotificationScheduled = true;
    queueMicrotask(() => {
      this.selectionNotificationScheduled = false;
      if (!this._destroyed) this.emitSelectionChange();
    });
  }

  private finishSelectionNotificationChain(): void {
    this.pendingSelectionChange = false;
    this.selectionNotificationCount = 0;
  }

  /**
   * Returns what the header area contains at the given client coordinates.
   * Returns null when the point is in the cell grid (not a header).
   */
  private getHeaderHit(
    clientX: number,
    clientY: number,
  ): { kind: 'corner' } | { kind: 'row'; row: number } | { kind: 'col'; col: number } | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.viewport.scale;
    const rect = this.canvasArea.getBoundingClientRect();
    // Same RTL un-mirror as getCellAt: map the screen x back to the logical-LTR
    // layout (row header on the left) before the header math below.
    const lx = this.screenX(clientX - rect.left, 0);
    const ly = clientY - rect.top;

    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const inRowHeader = lx < headerW;
    const inColHeader = ly < headerH;
    if (!inRowHeader && !inColHeader) return null;
    if (inRowHeader && inColHeader) return { kind: 'corner' };

    const geometry = getGridGeometryForWorksheet(ws);

    if (inRowHeader) {
      // Determine which row was clicked
      const innerY = ly - headerH;
      if (innerY < 0) return { kind: 'corner' };
      const r = geometry.rowAt(innerY, this.viewportTop, cs);
      return r === null ? null : { kind: 'row', row: r };
    }

    // inColHeader
    const innerX = lx - headerW;
    if (innerX < 0) return { kind: 'corner' };
    const c = geometry.colAt(innerX, this.effectiveScrollLeft, cs);
    return c === null ? null : { kind: 'col', col: c };
  }

  /**
   * If the pointer sits on a column/row-header border (within {@link
   * RESIZE_GRAB_PX}), return the resize target: which index to resize and the
   * fixed LTR edge it grows from (in canvasArea CSS px). Excel resizes the band
   * whose *trailing* border you grab — the column to the left of a vertical
   * border, the row above a horizontal one — so both that band and its
   * neighbour-to-the-far-side are checked. Geometry comes straight from {@link
   * getCellRect}, so the grab line always coincides with the drawn border at any
   * scroll offset / zoom / RTL. Returns null off the header borders.
   */
  private getResizeTarget(
    clientX: number,
    clientY: number,
  ): { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number } | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.viewport.scale;
    const rect = this.canvasArea.getBoundingClientRect();
    // Un-mirror the screen x to the logical-LTR space getCellRect draws in (the
    // same transform getHeaderHit uses), so the comparison holds for RTL sheets.
    const ptX = this.screenX(clientX - rect.left, 0);
    const ptY = clientY - rect.top;
    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const mdw = getGridGeometryForWorksheet(ws).maximumDigitWidth;

    // Column borders live in the column-header strip, right of the corner.
    if (ptY <= headerH && ptX > headerW) {
      const hit = this.getHeaderHit(clientX, clientY);
      if (hit?.kind !== 'col') return null;
      const origins = new Map<number, number>(); // index -> fixed LTR origin edge
      const edges: { index: number; edge: number }[] = [];
      for (const c of [hit.col - 1, hit.col]) {
        if (c < 1) continue;
        const r = this._cellRect(1, c); // x is independent of the row
        if (!r) continue;
        origins.set(c, r.x);
        edges.push({ index: c, edge: r.x + r.w }); // trailing (right) border
      }
      const index = resizeHitIndex(ptX, edges, RESIZE_GRAB_PX, headerW);
      if (index === null) return null;
      return { kind: 'col', index, originScaled: origins.get(index) as number, mdw };
    }

    // Row borders live in the row-header strip, below the corner.
    if (ptX <= headerW && ptY > headerH) {
      const hit = this.getHeaderHit(clientX, clientY);
      if (hit?.kind !== 'row') return null;
      const origins = new Map<number, number>(); // index -> fixed LTR origin edge
      const edges: { index: number; edge: number }[] = [];
      for (const rIdx of [hit.row - 1, hit.row]) {
        if (rIdx < 1) continue;
        const r = this._cellRect(rIdx, 1); // y is independent of the column
        if (!r) continue;
        origins.set(rIdx, r.y);
        edges.push({ index: rIdx, edge: r.y + r.h }); // trailing (bottom) border
      }
      const index = resizeHitIndex(ptY, edges, RESIZE_GRAB_PX, headerH);
      if (index === null) return null;
      return { kind: 'row', index, originScaled: origins.get(index) as number, mdw };
    }

    return null;
  }

  /**
   * Apply a live resize drag: size the band from its fixed origin edge to the
   * current pointer, clamp to {@link RESIZE_MIN_PX}, and write the result back
   * into the in-memory worksheet model in its native unit (Excel column widths /
   * points). This is a *view-only* mutation — the file is never written. The
   * memoized axis cache for this sheet is invalidated so every geometry read
   * (spacer, hit-test, overlay, renderer) sees the new size on the next frame.
   */
  private applyResize(clientX: number, clientY: number): void {
    const drag = this.resizeDrag;
    const ws = this.currentWorksheet;
    if (!drag || !ws) return;
    const cs = this.viewport.scale;
    const rect = this.canvasArea.getBoundingClientRect();

    if (drag.kind === 'col') {
      const ptX = this.screenX(clientX - rect.left, 0);
      const sizePx = Math.max(RESIZE_MIN_PX, Math.round((ptX - drag.originScaled) / cs));
      ws.colWidths[drag.index] = pxToColWidth(sizePx, drag.mdw);
      this.recordSizeOverride('col', drag.index);
    } else {
      const ptY = clientY - rect.top;
      const sizePx = Math.max(RESIZE_MIN_PX, Math.round((ptY - drag.originScaled) / cs));
      ws.rowHeights[drag.index] = pxToRowHeight(sizePx);
      this.recordSizeOverride('row', drag.index);
    }

    GridGeometry.invalidate(ws); // sizes changed → rebuild the cumulative-offset axes
    this.updateSpacerSize(ws);
    this.updateSelectionOverlay();
    // Live resize drag fires per pointermove; coalesce the canvas repaint into
    // one frame. The spacer (scrollbar extent) and overlay updates are cheap DOM
    // writes that must track the drag immediately, so they stay synchronous.
    this.scheduleRender();
  }

  /** Refit automatic rows once after a column-resize gesture. Doing this on
   * every pointermove would turn a drag into O(sheet cells × pointer events),
   * while Excel's observable result only needs to be committed at release. */
  private refitAutoRowsAfterColumnResize(): void {
    const ws = this.currentWorksheet;
    const workbook = this.preparedWorkbook;
    if (!ws || !workbook) return;
    const manualRows = this.sizeOverrideStore.get(this.currentSheet)?.rows.keys() ?? [];
    invalidateAutoRowHeights(ws, manualRows);
    const prepareRowHeights = workbook[prepareXlsxViewerRowHeights];
    if (typeof prepareRowHeights !== 'function') return;
    const measureCanvas = this.hostDocument.createElement('canvas');
    const measureCtx = measureCanvas.getContext('2d');
    if (!measureCtx) return;
    prepareRowHeights.call(workbook, ws, measureCtx);
    this.syncAutomaticRowOverrides(this.currentSheet, ws);
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
    const ws = this.currentWorksheet;
    const state = this.selectionState;
    if (!ws || !state) return { status: 'empty-selection' };
    if (state.areas.length !== 1) return { status: 'unsupported-multiple-areas' };
    const area = state.areas[0];

    // Whole-row/column/sheet selections are unbounded Excel concepts. Copying
    // narrows them to used cells without changing the logical selection.
    let maxRow = 1, maxCol = 1;
    for (const row of ws.rows) {
      if (row.index > maxRow) maxRow = row.index;
      for (const cell of row.cells) {
        if (cell.col > maxCol) maxCol = cell.col;
      }
    }

    const { r1, r2, c1, c2 } = area.kind === 'sheet'
      ? { r1: 1, r2: maxRow, c1: 1, c2: maxCol }
      : area.kind === 'rows'
        ? { r1: area.firstRow, r2: area.lastRow, c1: 1, c2: maxCol }
        : area.kind === 'columns'
          ? { r1: 1, r2: maxRow, c1: area.firstColumn, c2: area.lastColumn }
          : { r1: area.top, r2: area.bottom, c1: area.left, c2: area.right };

    const rowCount = r2 - r1 + 1;
    const colCount = c2 - c1 + 1;
    if (rowCount > Math.floor(MAX_CLIPBOARD_CELLS / colCount)) {
      return { status: 'too-large', limit: 'cells' };
    }
    const cellCount = rowCount * colCount;

    let utf16CodeUnits = Math.max(0, rowCount - 1) + rowCount * Math.max(0, colCount - 1);
    if (utf16CodeUnits > MAX_CLIPBOARD_UTF16_CODE_UNITS) {
      return { status: 'too-large', limit: 'text' };
    }
    const cellMap = new Map<number, Map<number, string>>();
    for (const row of ws.rows) {
      if (row.index < r1 || row.index > r2) continue;
      for (const cell of row.cells) {
        if (cell.col < c1 || cell.col > c2) continue;
        const v = cell.value;
        let text = this.wb?.cellText(ws, cell) ?? '';
        if (!this.wb) {
          if (v.type === 'text') text = v.runs ? v.runs.map((r) => r.text).join('') : v.text;
          else if (v.type === 'number') text = String(v.number);
          else if (v.type === 'bool') text = v.bool ? 'TRUE' : 'FALSE';
          else if (v.type === 'error') text = v.error;
        }
        if (text) {
          const encoded = encodeTsvFieldWithin(
            text,
            MAX_CLIPBOARD_UTF16_CODE_UNITS - utf16CodeUnits,
          );
          if (encoded === null) return { status: 'too-large', limit: 'text' };
          utf16CodeUnits += encoded.length;
          let values = cellMap.get(row.index);
          if (!values) { values = new Map(); cellMap.set(row.index, values); }
          values.set(cell.col, encoded);
        }
      }
    }

    const lines: string[] = [];
    for (let r = r1; r <= r2; r++) {
      const cols: string[] = [];
      const values = cellMap.get(r);
      for (let c = c1; c <= c2; c++) {
        const value = values?.get(c) ?? '';
        cols.push(value);
      }
      lines.push(cols.join('\t'));
    }
    const clipboard = this.hostWindow.navigator.clipboard;
    if (!clipboard) return { status: 'clipboard-unavailable' };
    try {
      await clipboard.writeText(lines.join('\n'));
      return { status: 'copied', cellCount, utf16CodeUnits };
    } catch {
      return { status: 'clipboard-denied' };
    }
  }

  private updateSelectionOverlay(): void {
    this.overlayHost.clearSelection();
    if (this.elementContext) {
      this.drawElementContextOverlay();
      return;
    }
    const state = this.selectionState;
    if (!state) return;
    const cs = this.viewport.scale;
    const ws = this.currentWorksheet;
    if (!ws) return;
    const sp = (px: number) => Math.round(px * cs);
    const headerW = sp(HEADER_W);
    const headerH = sp(HEADER_H);
    const width = this.canvasArea.clientWidth;
    const height = this.canvasArea.clientHeight;
    const geometry = getGridGeometryForWorksheet(ws);
    // Match renderViewport's physical freeze materialization. A legal freeze
    // count may cover the full sheet; it must never create million-row overlay
    // geometry when only a handful of bands can reach this viewport.
    const effective = geometry.effectiveFrozenBands({
      scale: cs, width, height, headerWidth: HEADER_W, headerHeight: HEADER_H,
      rows: ws.freezeRows ?? 0, cols: ws.freezeCols ?? 0,
    });
    const axes = geometry.axesAtScale(cs);
    const frozenW = axes.col.offsetOf(effective.cols + 1);
    const frozenH = axes.row.offsetOf(effective.rows + 1);
    const xPanes = effective.cols > 0
      ? [
          { first: 1, last: effective.cols, start: headerW, end: Math.min(width, headerW + frozenW) },
          { first: effective.cols + 1, last: MAX_WORKSHEET_COL, start: Math.min(width, headerW + frozenW), end: width },
        ]
      : [{ first: 1, last: MAX_WORKSHEET_COL, start: headerW, end: width }];
    const yPanes = effective.rows > 0
      ? [
          { first: 1, last: effective.rows, start: headerH, end: Math.min(height, headerH + frozenH) },
          { first: effective.rows + 1, last: MAX_WORKSHEET_ROW, start: Math.min(height, headerH + frozenH), end: height },
        ]
      : [{ first: 1, last: MAX_WORKSHEET_ROW, start: headerH, end: height }];
    const selectionColor = this.opts.selectionColor ?? DEFAULT_SELECTION_COLOR;
    const { background } = selectionOverlayStyle(selectionColor);
    const seenFragments = new Set<string>();
    const fillSubpaths: string[] = [];
    const overlayRects: SelectionOverlayRect[] = [];

    for (const area of state.areas) {
      const bounds = area.kind === 'cells'
        ? { top: area.top, bottom: area.bottom, left: area.left, right: area.right,
            topEdge: true, bottomEdge: true, leftEdge: true, rightEdge: true }
        : area.kind === 'rows'
          ? { top: area.firstRow, bottom: area.lastRow, left: 1, right: MAX_WORKSHEET_COL,
              topEdge: true, bottomEdge: true, leftEdge: false, rightEdge: false }
          : area.kind === 'columns'
            ? { top: 1, bottom: MAX_WORKSHEET_ROW, left: area.firstColumn, right: area.lastColumn,
                topEdge: false, bottomEdge: false, leftEdge: true, rightEdge: true }
            : { top: 1, bottom: MAX_WORKSHEET_ROW, left: 1, right: MAX_WORKSHEET_COL,
                topEdge: false, bottomEdge: false, leftEdge: false, rightEdge: false };

      for (const yp of yPanes) for (const xp of xPanes) {
        if (xp.end <= xp.start || yp.end <= yp.start) continue;
        const top = Math.max(bounds.top, yp.first);
        const bottom = Math.min(bounds.bottom, yp.last);
        const left = Math.max(bounds.left, xp.first);
        const right = Math.min(bounds.right, xp.last);
        if (top > bottom || left > right) continue;
        const tl = this._cellRect(top, left);
        const br = this._cellRect(bottom, right);
        if (!tl || !br) continue;
        const rawLeft = tl.x;
        const rawTop = tl.y;
        const rawRight = br.x + br.w;
        const rawBottom = br.y + br.h;
        const x = Math.max(rawLeft, xp.start);
        const y = Math.max(rawTop, yp.start);
        const x2 = Math.min(rawRight, xp.end);
        const y2 = Math.min(rawBottom, yp.end);
        const fragmentW = x2 - x;
        const fragmentH = y2 - y;
        if (fragmentW <= 0 || fragmentH <= 0) continue;

        // Only paint a border where the logical selection itself ends. Pane and
        // viewport clips are not selection edges and must not create fake lines.
        const topBorder = bounds.topEdge && top === bounds.top && rawTop >= yp.start;
        const bottomBorder = bounds.bottomEdge && bottom === bounds.bottom && rawBottom <= yp.end;
        const leftBorder = bounds.leftEdge && left === bounds.left && rawLeft >= xp.start;
        const rightBorder = bounds.rightEdge && right === bounds.right && rawRight <= xp.end;
        const screenLeft = this.screenX(x, fragmentW);
        const physicalLeftBorder = this.isRtl ? rightBorder : leftBorder;
        const physicalRightBorder = this.isRtl ? leftBorder : rightBorder;
        const fragmentKey = [
          screenLeft, y, fragmentW, fragmentH,
          topBorder, physicalRightBorder, bottomBorder, physicalLeftBorder,
        ].join('|');
        if (seenFragments.has(fragmentKey)) continue;
        seenFragments.add(fragmentKey);
        // Paint every fragment as a subpath in one SVG fill operation. With a
        // single non-zero fill, overlapping selection areas form a visual union
        // instead of stacking translucent backgrounds and becoming darker.
        fillSubpaths.push(
          `M${screenLeft} ${y}h${fragmentW}v${fragmentH}h${-fragmentW}Z`,
        );
        overlayRects.push({
          x: screenLeft,
          y,
          width: fragmentW,
          height: fragmentH,
          top: topBorder,
          right: physicalRightBorder,
          bottom: bottomBorder,
          left: physicalLeftBorder,
        });
      }
    }

    if (fillSubpaths.length > 0) {
      const svgNamespace = 'http://www.w3.org/2000/svg';
      const svg = this.hostDocument.createElementNS(svgNamespace, 'svg');
      svg.setAttribute('data-xlsx-selection-fill', '');
      svg.style.cssText =
        'position:absolute;inset:0;width:100%;height:100%;overflow:hidden;pointer-events:none;';
      const isMultipleAreaSelection = state.areas.length > 1;
      const activeRect = this._cellRect(state.activeCell.row, state.activeCell.col);
      const maskId = `xlsx-selection-mask-${++selectionMaskSequence}`;
      const defs = this.hostDocument.createElementNS(svgNamespace, 'defs');
      const mask = this.hostDocument.createElementNS(svgNamespace, 'mask');
      mask.setAttribute('id', maskId);
      mask.setAttribute('maskUnits', 'userSpaceOnUse');
      mask.setAttribute('x', '0');
      mask.setAttribute('y', '0');
      mask.setAttribute('width', String(width));
      mask.setAttribute('height', String(height));
      const selectedPath = this.hostDocument.createElementNS(svgNamespace, 'path');
      selectedPath.setAttribute('d', fillSubpaths.join(''));
      selectedPath.setAttribute('fill', '#fff');
      mask.appendChild(selectedPath);

      // Excel leaves ActiveCell unshaded so it remains distinct from the
      // selected cells. ActiveCell stays at the drag origin; only the Area's
      // opposite corner changes during extension.
      if (activeRect) {
        for (const yp of yPanes) for (const xp of xPanes) {
          const clippedX = Math.max(activeRect.x, xp.start);
          const clippedY = Math.max(activeRect.y, yp.start);
          const clippedX2 = Math.min(activeRect.x + activeRect.w, xp.end);
          const clippedY2 = Math.min(activeRect.y + activeRect.h, yp.end);
          if (clippedX2 <= clippedX || clippedY2 <= clippedY) continue;
          const cutout = this.hostDocument.createElementNS(svgNamespace, 'rect');
          cutout.setAttribute('data-xlsx-active-cell-cutout', '');
          cutout.setAttribute('x', String(this.screenX(clippedX, clippedX2 - clippedX)));
          cutout.setAttribute('y', String(clippedY));
          cutout.setAttribute('width', String(clippedX2 - clippedX));
          cutout.setAttribute('height', String(clippedY2 - clippedY));
          cutout.setAttribute('fill', '#000');
          mask.appendChild(cutout);
        }
      }
      defs.appendChild(mask);
      svg.appendChild(defs);

      const fill = this.hostDocument.createElementNS(svgNamespace, 'rect');
      fill.setAttribute('x', '0');
      fill.setAttribute('y', '0');
      fill.setAttribute('width', String(width));
      fill.setAttribute('height', String(height));
      fill.setAttribute('fill', background);
      fill.setAttribute('mask', `url(#${maskId})`);
      svg.appendChild(fill);

      const boundaryPath = isMultipleAreaSelection ? '' : selectionBoundaryPath(overlayRects);
      if (boundaryPath) {
        const boundary = this.hostDocument.createElementNS(svgNamespace, 'path');
        boundary.setAttribute('data-xlsx-selection-border', '');
        boundary.setAttribute('d', boundaryPath);
        boundary.setAttribute('fill', 'none');
        boundary.setAttribute('stroke', selectionColor);
        boundary.setAttribute('stroke-width', '2');
        boundary.setAttribute('stroke-linecap', 'square');
        boundary.setAttribute('stroke-linejoin', 'miter');
        svg.appendChild(boundary);
      }
      if (activeRect && isMultipleAreaSelection) {
        for (const yp of yPanes) for (const xp of xPanes) {
          const clippedX = Math.max(activeRect.x, xp.start);
          const clippedY = Math.max(activeRect.y, yp.start);
          const clippedX2 = Math.min(activeRect.x + activeRect.w, xp.end);
          const clippedY2 = Math.min(activeRect.y + activeRect.h, yp.end);
          if (clippedX2 <= clippedX || clippedY2 <= clippedY) continue;
          const focus = this.hostDocument.createElementNS(svgNamespace, 'rect');
          focus.setAttribute('data-xlsx-active-cell-border', '');
          focus.setAttribute('x', String(this.screenX(clippedX, clippedX2 - clippedX)));
          focus.setAttribute('y', String(clippedY));
          focus.setAttribute('width', String(clippedX2 - clippedX));
          focus.setAttribute('height', String(clippedY2 - clippedY));
          focus.setAttribute('fill', 'none');
          focus.setAttribute('stroke', selectionColor);
          focus.setAttribute('stroke-width', '1');
          svg.appendChild(focus);
        }
      }
      this.overlayHost.appendSelection(svg as unknown as HTMLElement);
    }

    // List data-validation dropdown arrow (ECMA-376 §18.3.1.33). Excel shows an
    // in-cell dropdown button only while the cell is *selected* and only for
    // `list`-type rules — so it is drawn here (selection overlay) rather than in
    // the canvas renderer. The button itself is non-interactive
    // (pointer-events:none); clicks are hit-tested against its rect in the
    // pointerdown handler, which opens a panel listing the allowed values
    // (display only — picking a value never changes the cell).
    this.maybeDrawValidationDropdown();
  }

  private drawElementContextOverlay(): void {
    const context = this.elementContext;
    const worksheet = this.currentWorksheet;
    const viewport = this.elementContextViewport();
    if (!context || !worksheet || !viewport || context.sheetIndex !== this.currentSheet) return;
    const projection = projectXlsxElementContext(worksheet, context, viewport);
    if (!projection) return;
    const clip = this.hostDocument.createElement('div');
    clip.setAttribute('data-xlsx-element-context-clip', '');
    clip.style.cssText =
      `position:absolute;left:${projection.clip.x}px;top:${projection.clip.y}px;` +
      `width:${projection.clip.width}px;height:${projection.clip.height}px;` +
      'overflow:hidden;pointer-events:none;';
    const frame = this.hostDocument.createElement('div');
    frame.setAttribute('data-xlsx-element-context-outline', context.elementType);
    const color = this.opts.selectionColor ?? DEFAULT_SELECTION_COLOR;
    frame.style.cssText =
      `position:absolute;left:${projection.rect.x - projection.clip.x}px;` +
      `top:${projection.rect.y - projection.clip.y}px;` +
      `width:${projection.rect.width}px;height:${projection.rect.height}px;` +
      `box-sizing:border-box;border:2px solid ${color};` +
      `background:color-mix(in srgb, ${color} 6%, transparent);` +
      `transform:rotate(${projection.rotation}deg);transform-origin:center;pointer-events:none;`;
    clip.appendChild(frame);
    this.overlayHost.appendSelection(clip);
  }

  /** Draw the Excel list-validation dropdown button just outside the
   *  bottom-right corner of the *active* cell when that cell is covered by a
   *  `list` data-validation rule. Anchored to the single active cell (not the
   *  whole range) to mirror Excel, which attaches the button to the active
   *  cell of the selection. */
  private maybeDrawValidationDropdown(): void {
    // The overlay is rebuilt on every selection / scroll change, so the
    // arrow's hit-test rect is recomputed here each time (cleared when no arrow
    // is currently shown).
    this.validationArrowRect = null;
    if (this.selectionMode !== 'cells') return;
    const ws = this.currentWorksheet;
    const active = this.activeCell;
    if (!ws || !active) return;
    const dv = findListValidationAt(ws.dataValidations, active.row, active.col);
    if (!dv) return;

    const rect = this._cellRect(active.row, active.col);
    if (!rect) return;

    // Excel's dropdown button is a fixed square sized to the cell height,
    // clamped to a sensible range so it stays usable at small zoom and doesn't
    // dominate tall rows. The arrow glyph is centered inside.
    const cs = this.viewport.scale;
    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const side = Math.max(14, Math.min(rect.h, 22 * cs));
    // Button sits flush to the right of the cell, top-aligned with it.
    const btnLogicalX = rect.x + rect.w;
    const btnY = rect.y;
    // Cull when the active cell (hence its button) is scrolled behind the
    // fixed headers.
    if (btnLogicalX + side <= headerW || btnY + side <= headerH) return;

    const screenLeft = this.screenX(btnLogicalX, side);

    const btn = this.hostDocument.createElement('div');
    btn.setAttribute('data-xlsx-validation-dropdown', '');
    btn.style.cssText =
      `position:absolute;` +
      `left:${screenLeft}px;top:${btnY}px;width:${side}px;height:${side}px;` +
      `box-sizing:border-box;display:flex;align-items:center;justify-content:center;` +
      // Match Excel's grey button chrome; non-interactive (display only).
      `background:#f0f0f0;border:1px solid #7f7f7f;pointer-events:none;`;
    const arrow = Math.max(4, Math.round(side * 0.42));
    btn.innerHTML =
      `<svg width="${arrow}" height="${arrow}" viewBox="0 0 10 6" aria-hidden="true">` +
      `<path d="M0 0 L10 0 L5 6 Z" fill="#333"/></svg>`;
    this.overlayHost.appendSelection(btn);

    // Record the arrow's on-screen rect (canvasArea space) for pointer
    // hit-testing. The button element has pointer-events:none, so clicks fall
    // through to the scrollHost where the pointerdown handler tests this rect.
    this.validationArrowRect = { x: screenLeft, y: btnY, w: side, h: side };

    // Keep an already-open panel glued to the arrow as the grid scrolls. If the
    // active cell's validation differs from the open panel (selection moved),
    // close it instead.
    if (this.validationPanel.style.display !== 'none') {
      if (this.validationPanelKey === `${active.row}:${active.col}`) {
        this.positionValidationPanel();
      } else {
        this.hideValidationPanel();
      }
    }
  }

  // ─── IX2 find-highlight overlay ──────────────────────────────────────────

  /**
   * Redraw the find-highlight overlay: one translucent box per matched cell on
   * the current sheet, the active match in a stronger colour. Uses the SAME
   * `getCellRect` + `screenX` + header/frozen clamp the selection overlay uses,
   * so a box lands exactly on the drawn cell at any scroll offset / zoom / RTL.
   * Rebuilt on every render and scroll (cheap DOM geometry, no canvas paint).
   */
  private updateFindOverlay(): void {
    this.overlayHost.clearFind();
    const ws = this.currentWorksheet;
    if (!ws) return;
    const cs = this.viewport.scale;
    const sp = (px: number) => Math.round(px * cs);
    const headerW = sp(HEADER_W);
    const headerH = sp(HEADER_H);
    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;
    const frozen = getGridGeometryForWorksheet(ws).roundedFrozenExtent(cs);
    const frozenBoundX = headerW + frozen.width;
    const frozenBoundY = headerH + frozen.height;

    // A match accent: same single-color → border + translucent fill derivation
    // the selection overlay uses. The active match uses a warm accent so it is
    // distinguishable from other hits and from the (blue) selection box.
    const other = findHighlightOverlayStyle(false, this.opts.findHighlightColors);
    const active = findHighlightOverlayStyle(true, this.opts.findHighlightColors);

    for (const hl of this._find.sheetHighlights(this.currentSheet)) {
      const rect = this._cellRect(hl.row, hl.col);
      if (!rect) continue;
      let { x, y, w, h } = rect;
      // Clamp against headers + the frozen-pane boundary (scrollable cells that
      // scrolled behind the frozen area are clipped there), mirroring the
      // selection overlay so a highlight never spills over fixed regions.
      if (x < headerW) { w -= headerW - x; x = headerW; }
      if (y < headerH) { h -= headerH - y; y = headerH; }
      if (hl.col > freezeCols && x < frozenBoundX) { w -= frozenBoundX - x; x = frozenBoundX; }
      if (hl.row > freezeRows && y < frozenBoundY) { h -= frozenBoundY - y; y = frozenBoundY; }
      if (w <= 0 || h <= 0) continue;
      const screenLeft = this.screenX(x, w);
      const { border, background } = hl.active ? active : other;
      const box = this.hostDocument.createElement('div');
      box.style.cssText =
        `position:absolute;` +
        `left:${screenLeft}px;top:${y}px;width:${w}px;height:${h}px;` +
        `box-sizing:border-box;border:${border};background:${background};pointer-events:none;`;
      this.overlayHost.appendFind(box);
    }
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
    if (!this.wb) return [];
    const matches = await this._find.find(query, opts);
    this.updateFindOverlay();
    return matches;
  }

  /**
   * IX2 — move to the next match (wrap-around), switching sheets and scrolling
   * the matched cell into view as needed, and highlight it as the active match.
   * Returns the now-active match, or `null` when there are none. Call
   * {@link findText} first.
   */
  async findNext(): Promise<FindMatch<XlsxMatchLocation> | null> {
    return this._activateMatch(this._find.next());
  }

  /** IX2 — move to the previous match (wrap-around). */
  async findPrev(): Promise<FindMatch<XlsxMatchLocation> | null> {
    return this._activateMatch(this._find.prev());
  }

  /** IX2 — clear all highlights and reset the find state. */
  clearFind(): void {
    this._find.invalidate();
    this.updateFindOverlay();
  }

  private async _activateMatch(
    match: FindMatch<XlsxMatchLocation> | null,
  ): Promise<FindMatch<XlsxMatchLocation> | null> {
    if (!match) {
      this.updateFindOverlay();
      return null;
    }
    const { sheet, row, col } = match.location;
    if (sheet !== this.currentSheet) {
      // showSheet resets scroll/selection and re-renders; the find state (and so
      // the highlights) survive because they live on the controller, not the
      // sheet. updateFindOverlay runs after the sheet switch below.
      await this.goToSheet(sheet);
    }
    this._scrollCellIntoView(row, col);
    // Scrolling schedules a coalesced render; draw the highlights now so the
    // active box is visible immediately without waiting a frame.
    this.updateFindOverlay();
    return match;
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

  // ─── List data-validation dropdown panel (display-only) ───────────────────

  /** Toggle the dropdown panel for the active cell's list validation. Called
   *  from pointerdown when the arrow rect is hit. Re-clicking the same arrow
   *  closes it. */
  private toggleValidationPanel(): void {
    const ws = this.currentWorksheet;
    const active = this.activeCell;
    if (!ws || !active) return;
    const key = `${active.row}:${active.col}`;
    if (this.validationPanelKey === key) {
      this.hideValidationPanel();
      return;
    }
    const dv = findListValidationAt(ws.dataValidations, active.row, active.col);
    if (!dv) return;
    this.hideValidationPanel();
    this.validationPanelKey = key;
    void this.openValidationPanel(active, dv.formula1);
  }

  /** Resolve the allowed values for `formula1` (relative to the current sheet)
   *  and render them in the panel anchored below the active cell. Async because
   *  cross-sheet range references may need a lazily-parsed worksheet. */
  private async openValidationPanel(cell: CellAddress, formula1: string | undefined): Promise<void> {
    const generation = ++this.validationRequestGeneration;
    const workbook = this.wb;
    const sheet = this.currentSheet;
    if (!workbook || this._destroyed) return;
    let resolved: ResolvedList;
    try {
      resolved = await workbook.resolveValidationList(sheet, formula1);
    } catch {
      if (!this.isCurrentValidationRequest(generation, workbook, sheet, cell)) return;
      // A resolution failure (e.g. a missing sheet) must not break the viewer;
      // fall back to disclosing the raw formula.
      resolved = { kind: 'formula', formula: formula1 ?? '' };
    }
    if (!this.isCurrentValidationRequest(generation, workbook, sheet, cell)) return;

    this.renderValidationPanel(resolved);
    this.positionValidationPanel();
    this.installValidationOutsideHandler();
  }

  private isCurrentValidationRequest(
    generation: number,
    workbook: XlsxWorkbook,
    sheet: number,
    cell: CellAddress,
  ): boolean {
    const active = this.activeCell;
    return !this._destroyed
      && generation === this.validationRequestGeneration
      && this.wb === workbook
      && this.currentSheet === sheet
      && this.validationPanelKey === `${cell.row}:${cell.col}`
      && active?.row === cell.row
      && active?.col === cell.col;
  }

  /** Build the panel's children. Uses textContent throughout (no HTML injection
   *  from cell values). Items highlight on hover but are NOT selectable —
   *  this is a read-only viewer, so clicking a value must not change the cell. */
  private renderValidationPanel(resolved: ResolvedList): void {
    const panel = this.validationPanel;
    panel.textContent = '';
    if (resolved.kind === 'formula' || resolved.values.length === 0) {
      // Unresolved operand (named range / complex formula) or an empty range:
      // disclose the formula / a placeholder rather than showing a blank box.
      const note = this.hostDocument.createElement('div');
      note.style.cssText = 'padding:4px 8px;color:#666;font-style:italic;white-space:pre-wrap;word-break:break-word;';
      note.textContent =
        resolved.kind === 'formula'
          ? (resolved.formula ? `= ${resolved.formula}` : '(no list)')
          : '(empty list)';
      panel.appendChild(note);
      return;
    }
    for (const value of resolved.values) {
      const item = this.hostDocument.createElement('div');
      item.setAttribute('data-xlsx-validation-item', '');
      item.style.cssText = 'padding:3px 8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:default;';
      item.textContent = value;
      // Hover highlight only — no click/select (read-only viewer).
      item.addEventListener('pointerenter', () => {
        item.style.background = '#cfe3ff';
      });
      item.addEventListener('pointerleave', () => {
        item.style.background = '';
      });
      panel.appendChild(item);
    }
  }

  /** Position the (already-populated, visible-or-becoming-visible) panel below
   *  the dropdown arrow / active cell using the pure geometry calculator. */
  private positionValidationPanel(): void {
    const active = this.activeCell;
    if (!active) return;
    const rect = this._cellRect(active.row, active.col);
    if (!rect) return;
    const screenLeft = this.screenX(rect.x, rect.w);
    // Make it measurable off-screen first so offsetWidth/Height reflect content.
    this.validationPanel.style.left = '-9999px';
    this.validationPanel.style.top = '-9999px';
    this.validationPanel.style.display = 'block';
    const pos = computeValidationPanelPosition({
      cell: { x: screenLeft, y: rect.y, w: rect.w, h: rect.h },
      panel: { w: this.validationPanel.offsetWidth, h: this.validationPanel.offsetHeight },
      viewport: { w: this.canvasArea.clientWidth, h: this.canvasArea.clientHeight },
      rtl: this.isRtl,
    });
    this.overlayHost.showValidation(pos.left, pos.top);
  }

  /** Install a document-level pointerdown listener that closes the panel on a
   *  click outside it (and outside the arrow, which toggles via its own path).
   *  Removed by {@link hideValidationPanel}. */
  private installValidationOutsideHandler(): void {
    if (this.validationOutsideHandler) return;
    this.validationOutsideHandler = (e: PointerEvent) => {
      const target = e.target as Node | null;
      if (target && this.validationPanel.contains(target)) return; // inside panel
      // A click on the arrow is handled by the scrollHost pointerdown (toggle);
      // don't double-handle it here. Detect by hit-testing the arrow rect.
      const { x: ax, y: ay } = this.surface.localPoint(e.clientX, e.clientY);
      const ar = this.validationArrowRect;
      if (ar && ax >= ar.x && ax <= ar.x + ar.w && ay >= ar.y && ay <= ar.y + ar.h) {
        return;
      }
      this.hideValidationPanel();
    };
    // Capture phase so we see the click before it mutates selection.
    this.hostDocument.addEventListener('pointerdown', this.validationOutsideHandler, true);
  }

  /** Hide the panel and detach its outside-click listener. Called on re-click,
   *  outside click, Esc, scroll, selection change, sheet switch and destroy. */
  private hideValidationPanel(): void {
    this.validationRequestGeneration++;
    this.overlayHost.hideValidation();
    this.validationPanelKey = null;
    if (this.validationOutsideHandler) {
      this.hostDocument.removeEventListener('pointerdown', this.validationOutsideHandler, true);
      this.validationOutsideHandler = null;
    }
  }

  // ─── Comment hover popup ──────────────────────────────────────────────────

  /** Build the `"row:col"` → comment index for the given sheet. Parses each
   *  `XlsxComment.cellRef` with the shared {@link parseA1}; later refs win on a
   *  collision (Excel allows at most one note per cell, so this is moot in
   *  practice). */
  private buildCommentMap(ws: Worksheet): void {
    this.commentMap = this.createCommentMap(ws.comments ?? []);
  }

  private createCommentMap(comments: readonly XlsxComment[]): Map<string, XlsxComment> {
    const map = new Map<string, XlsxComment>();
    for (const c of comments) {
      const p = parseA1(c.cellRef);
      if (p) map.set(`${p.row}:${p.col}`, c);
    }
    return map;
  }

  private createVisibleSheetView(source: Worksheet): Worksheet {
    const worksheet = createSheetViewModel(source);
    if (this.opts.comments === false) {
      return { ...worksheet, commentRefs: [], comments: [] };
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
    return {
      ...worksheet,
      commentRefs: worksheet.commentRefs?.filter((ref) => !resolved.has(ref)),
      comments: worksheet.comments?.filter((comment) => !resolved.has(comment.cellRef)),
    };
  }

  /** IX1 — index the current sheet's hyperlinks by `"row:col"` (1-based, first
   *  cell of the `ref` range) so a clicked/hovered cell resolves in O(1). Keys
   *  match the renderer's `hyperlinkMap` exactly (`${hl.row}:${hl.col}`). */
  private buildHyperlinkMap(ws: Worksheet): void {
    this.hyperlinkMap = new Map();
    for (const hl of ws.hyperlinks ?? []) {
      this.hyperlinkMap.set(`${hl.row}:${hl.col}`, hl);
    }
  }

  /** IX1 — the hyperlink at a cell, or null. `getCellAt` returns 1-based
   *  {row,col}, matching the parser/renderer keying.
   *
   *  Returns null unconditionally when `enableHyperlinks` is `false`: this is the
   *  single gate that disables hyperlink interactivity. Both consumers — the
   *  pointermove pointer-cursor affordance and the click dispatch
   *  ({@link dispatchHyperlink}) — funnel through this hit-test, so a null result
   *  means no cursor change, no default navigation, and no `onHyperlinkClick`. */
  private hyperlinkAtCell(cell: CellAddress): Hyperlink | null {
    if (this.opts.enableHyperlinks === false) return null;
    return this.hyperlinkMap.get(`${cell.row}:${cell.col}`) ?? null;
  }

  /**
   * IX1 — dispatch a click on a hyperlinked cell. Builds a
   * {@link HyperlinkTarget} from the parsed hyperlink (external `url` wins over
   * internal `location`, matching Excel: a `<hyperlink>` carrying both navigates
   * to the external target) and routes it to the caller's `onHyperlinkClick`
   * (which fully owns behaviour) or the built-in default. Returns true when a
   * hyperlink was found and dispatched.
   */
  private dispatchHyperlink(cell: CellAddress): boolean {
    const hl = this.hyperlinkAtCell(cell);
    if (!hl) return false;
    let target: HyperlinkTarget;
    if (hl.url) {
      target = { kind: 'external', url: hl.url };
    } else if (hl.location) {
      target = { kind: 'internal', ref: hl.location };
    } else {
      return false; // parser only emits a hyperlink with url or location
    }
    const custom = this.opts.onHyperlinkClick;
    if (custom) {
      custom(target);
      return true;
    }
    // Built-in default. External: open in a new tab, sanitised against the safe
    // scheme allowlist (a blocked scheme like `javascript:` is a no-op, not a
    // navigation). Internal: best-effort sheet navigation, below.
    if (target.kind === 'external') {
      openExternalHyperlink(target.url, undefined, this.hostWindow);
    } else {
      void this.navigateInternalHyperlink(target.ref).catch(
        (error) => this._reportRenderError(error),
      );
    }
    return true;
  }

  /**
   * IX1 default handler for an internal `location` target (§18.3.1.47): resolve
   * a direct cell/range or an in-scope defined name (§18.2.5), switch sheets when
   * needed, then scroll the first referenced cell into view.
   */
  private async navigateInternalHyperlink(location: string): Promise<void> {
    const target = resolveXlsxInternalHyperlink(
      location,
      this.currentSheet,
      this.sheetNames,
      this.currentWorksheet?.definedNames ?? [],
    );
    if (!target) return;
    if (target.sheetIndex !== this.currentSheet) {
      await this.goToSheet(target.sheetIndex);
    }
    await this.scrollToCell(target.cellRef);
  }

  /** Show the popup for the comment on `cell` after the hover dwell, anchored to
   *  the cell's current on-screen rect. No-op when the cell carries no comment.
   *  Re-hovering the same cell does not restart the timer. */
  private scheduleCommentPopup(cell: CellAddress): void {
    const key = `${cell.row}:${cell.col}`;
    const comment = this.commentMap.get(key);
    if (!comment) {
      this.hideCommentPopup();
      return;
    }
    if (this.commentPopupKey === key) return; // already shown / pending here
    this.hideCommentPopup();
    this.commentPopupKey = key;
    this.commentPopupTimer = setTimeout(() => {
      this.commentPopupTimer = null;
      void this.renderCommentPopup(cell, comment).catch((error) => this._reportRenderError(error));
    }, COMMENT_POPUP_DELAY_MS);
  }

  private async loadCommentUi(): Promise<XlsxCommentUiRuntime> {
    const commentUi = this.commentUi ?? await loadXlsxCommentUiRuntime();
    if (!this._destroyed) this.commentUi = commentUi;
    return commentUi;
  }

  /** Immediately render the popup for `comment` anchored to `cell` (used by the
   *  hover-dwell timer and by touch selection, which has no hover). */
  private async renderCommentPopup(cell: CellAddress, comment: XlsxComment): Promise<void> {
    if (!this._cellRect(cell.row, cell.col)) return;
    const generation = ++this.commentPopupRenderGeneration;
    const commentUi = await this.loadCommentUi();
    if (this._destroyed || generation !== this.commentPopupRenderGeneration) return;
    if (!this._cellRect(cell.row, cell.col)) return;
    this.commentPopupCell = cell;

    // Use the same card structure and default theme as the DOCX/PPTX margins;
    // XLSX owns only the cell-anchored popup geometry.
    const occurrenceKey = `sheet:${this.currentSheet}:cell:${comment.cellRef}:comment:${comment.id ?? 'root'}`;
    const thread: ReadOnlyCommentThread = {
      occurrenceKey,
      root: {
        messageKey: `${occurrenceKey}:root`,
        sourceId: comment.id,
        author: comment.author,
        date: comment.date,
        text: comment.rootText ?? comment.text,
        status: comment.resolved ? 'resolved' : 'active',
      },
      replies: (comment.replies ?? []).map((reply, index) => ({
        messageKey: `${occurrenceKey}:reply:${reply.id ?? index}`,
        sourceId: reply.id,
        author: reply.author,
        date: reply.date,
        text: reply.text,
        status: reply.resolved ? 'resolved' : 'active',
      })),
    };
    commentUi.paintReadOnlyCommentCard(this.commentPopup, thread, {
      interactive: false,
      standalone: true,
    });
    const rootText = (comment.rootText ?? comment.text).trim();
    const byAuthor = comment.author?.trim() ? ` by ${comment.author.trim()}` : '';
    const replyCount = comment.replies?.length ?? 0;
    const replies = replyCount === 0
      ? ''
      : `; ${replyCount} ${replyCount === 1 ? 'reply' : 'replies'}`;
    this.overlayHost.announceComment(
      `Comment on ${comment.cellRef}${byAuthor}${rootText ? `: ${rootText}` : ''}${replies}`,
    );
    this.commentPopup.dataset.ooxmlCommentUi = 'popup';
    this.commentPopup.style.maxWidth = `${COMMENT_POPUP_MAX_W}px`;
    this.commentPopup.style.maxHeight = `${COMMENT_POPUP_MAX_H}px`;

    // Anchor to the cell's *screen* rect (RTL already mirrored by screenX), then
    // run the pure position calc against the popup's measured size. Make it
    // visible (off-screen) first so offsetWidth/Height reflect the wrapped text.
    this.commentPopup.style.left = '-9999px';
    this.commentPopup.style.top = '-9999px';
    this.commentPopup.style.display = '';
    this.positionCommentPopup();
  }

  private scheduleCommentPopupPosition(): void {
    if (this.commentPopupPositionScheduled || !this.commentPopupCell) return;
    this.commentPopupPositionScheduled = true;
    const position = (): void => {
      this.commentPopupPositionScheduled = false;
      this.positionCommentPopup();
    };
    const ownerWindow = this.hostDocument.defaultView;
    if (ownerWindow?.requestAnimationFrame) ownerWindow.requestAnimationFrame(position);
    else queueMicrotask(position);
  }

  private positionCommentPopup(): void {
    const cell = this.commentPopupCell;
    if (!cell || this.commentPopup.style.display === 'none') return;
    const rect = this._cellRect(cell.row, cell.col);
    if (!rect) return;
    const screenLeft = this.screenX(rect.x, rect.w);
    const pos = computeCommentPopupPosition({
      cell: { x: screenLeft, y: rect.y, w: rect.w, h: rect.h },
      popup: { w: this.commentPopup.offsetWidth, h: this.commentPopup.offsetHeight },
      viewport: { w: this.canvasArea.clientWidth, h: this.canvasArea.clientHeight },
      rtl: this.isRtl,
    });
    this.overlayHost.showComment(pos.left, pos.top);
  }

  /** Hide the popup and cancel any pending show. Called on cell-out, scroll,
   *  sheet switch and destroy. */
  private hideCommentPopup(): void {
    this.commentPopupRenderGeneration++;
    if (this.commentPopupTimer !== null) {
      clearTimeout(this.commentPopupTimer);
      this.commentPopupTimer = null;
    }
    this.commentPopupKey = null;
    this.commentPopupCell = null;
    this.overlayHost.hideComment();
    this.commentPopup.replaceChildren();
  }

  private applyPointerSelection(
    clientX: number,
    clientY: number,
    shiftKey: boolean,
    additiveKey: boolean,
    pointerId: number,
    allowDrag: boolean,
  ): void {
    const headerHit = this.getHeaderHit(clientX, clientY);

    if (headerHit) {
      if (headerHit.kind === 'corner') {
        // Select all — no drag extension needed
        this.selectionController.select({ row: 1, col: 1 }, 'all');
        this.selectionController.endDrag();
      } else if (headerHit.kind === 'row') {
        if (shiftKey && this.anchorCell && this.selectionMode === 'rows') {
          this.selectionController.extend({ row: headerHit.row, col: 1 });
        } else {
          const selected = additiveKey
            ? this.selectionController.add({ row: headerHit.row, col: 1 }, 'rows')
            : (this.selectionController.select({ row: headerHit.row, col: 1 }, 'rows'), true);
          if (allowDrag && selected) {
            this.beginSelectionDrag(pointerId);
            this.scrollHost.setPointerCapture(pointerId);
          }
        }
      } else {
        if (shiftKey && this.anchorCell && this.selectionMode === 'cols') {
          this.selectionController.extend({ row: 1, col: headerHit.col });
        } else {
          const selected = additiveKey
            ? this.selectionController.add({ row: 1, col: headerHit.col }, 'cols')
            : (this.selectionController.select({ row: 1, col: headerHit.col }, 'cols'), true);
          if (allowDrag && selected) {
            this.beginSelectionDrag(pointerId);
            this.scrollHost.setPointerCapture(pointerId);
          }
        }
      }
      this.updateSelectionOverlay();
      void this.renderCurrentSheet().catch((error) => this._reportRenderError(error));
      this.emitSelectionChange();
      return;
    }

    const cell = this.getCellAt(clientX, clientY);
    if (!cell) return;

    let selected = true;
    if (shiftKey && this.anchorCell && this.selectionMode === 'cells') {
      this.selectionController.extend(cell);
    } else {
      selected = additiveKey
        ? this.selectionController.add(cell, 'cells')
        : (this.selectionController.select(cell, 'cells'), true);
    }
    if (allowDrag && selected) {
      this.beginSelectionDrag(pointerId);
      this.scrollHost.setPointerCapture(pointerId);
    }
    this.updateSelectionOverlay();
    if (this.wb) {
      this.renderCurrentSheet().catch((error) => this._reportRenderError(error));
    }
    this.emitSelectionChange();
  }

  /** Browser-visible input box, excluding classic native scrollbar gutters. */
  private viewportInputBounds(): { left: number; top: number; width: number; height: number } {
    const rect = this.canvasArea.getBoundingClientRect();
    const left = rect.left + this.scrollHost.clientLeft;
    const top = rect.top + this.scrollHost.clientTop;
    const availableWidth = Math.max(0, rect.width - this.scrollHost.clientLeft);
    const availableHeight = Math.max(0, rect.height - this.scrollHost.clientTop);
    return {
      left,
      top,
      width: Math.min(availableWidth, this.scrollHost.clientWidth || availableWidth),
      height: Math.min(availableHeight, this.scrollHost.clientHeight || availableHeight),
    };
  }

  /** Extend the active drag selection to the pointer's cell. Captured pointers
   * outside the canvas and auto-scroll ticks clamp to the visible data edge so
   * selection never jumps ahead of the viewport. */
  private extendDragSelection(
    clientX: number,
    clientY: number,
    clampToViewport: boolean,
  ): boolean {
    let pointerX = clientX;
    let pointerY = clientY;
    const bounds = this.viewportInputBounds();
    const outsideViewport = clientX < bounds.left || clientX >= bounds.left + bounds.width ||
      clientY < bounds.top || clientY >= bounds.top + bounds.height;
    if (clampToViewport || outsideViewport) {
      const cs = this.viewport.scale;
      const headerW = Math.round(HEADER_W * cs);
      const headerH = Math.round(HEADER_H * cs);
      const dataLeft = bounds.left + (this.isRtl ? 0 : headerW);
      const dataRight = bounds.left + bounds.width - (this.isRtl ? headerW : 0);
      pointerX = Math.min(dataRight - 1, Math.max(dataLeft + 1, pointerX));
      pointerY = Math.min(
        bounds.top + bounds.height - 1,
        Math.max(bounds.top + headerH + 1, pointerY),
      );
    }

    if (this.selectionMode === 'rows') {
      const hit = clampToViewport ? null : this.getHeaderHit(pointerX, pointerY);
      const row = hit?.kind === 'row'
        ? hit.row
        : this.getCellAt(pointerX, pointerY)?.row;
      if (!row || row === this.activeCell?.row) return false;
      this.selectionController.extend({ row, col: 1 });
      return true;
    }

    if (this.selectionMode === 'cols') {
      const hit = clampToViewport ? null : this.getHeaderHit(pointerX, pointerY);
      const col = hit?.kind === 'col'
        ? hit.col
        : this.getCellAt(pointerX, pointerY)?.col;
      if (!col || col === this.activeCell?.col) return false;
      this.selectionController.extend({ row: 1, col });
      return true;
    }

    const cell = this.getCellAt(pointerX, pointerY);
    if (!cell || (cell.row === this.activeCell?.row && cell.col === this.activeCell?.col)) {
      return false;
    }
    this.selectionController.extend(cell);
    return true;
  }

  private selectionAutoScrollSpeed(): { x: number; y: number } {
    const pointer = this.selectionAutoScrollPointer;
    if (!pointer) return { x: 0, y: 0 };
    const bounds = this.viewportInputBounds();
    return selectionAutoScrollVelocity(
      { x: pointer.clientX - bounds.left, y: pointer.clientY - bounds.top },
      { width: bounds.width, height: bounds.height },
      this.isRtl,
      this.selectionMode,
    );
  }

  private trackSelectionAutoScroll(e: PointerEvent): void {
    if (e.pointerId !== this.selectionPointerId) return;
    this.selectionAutoScrollPointer = {
      clientX: e.clientX,
      clientY: e.clientY,
      pointerId: e.pointerId,
    };
    const speed = this.selectionAutoScrollSpeed();
    if (speed.x === 0 && speed.y === 0) {
      this.stopSelectionAutoScroll();
      return;
    }
    if (this.selectionAutoScrollFrame !== null) return;
    this.selectionAutoScrollLastTime = null;
    this.selectionAutoScrollFrame = this.hostWindow.requestAnimationFrame(
      (time) => this.runSelectionAutoScroll(time),
    );
  }

  private runSelectionAutoScroll(time: number): void {
    this.selectionAutoScrollFrame = null;
    const pointer = this.selectionAutoScrollPointer;
    if (
      !pointer ||
      pointer.pointerId !== this.selectionPointerId ||
      !this.isSelecting ||
      this._destroyed
    ) {
      this.stopSelectionAutoScroll();
      return;
    }

    const speed = this.selectionAutoScrollSpeed();
    if (speed.x === 0 && speed.y === 0) {
      this.stopSelectionAutoScroll();
      return;
    }

    const previousTime = this.selectionAutoScrollLastTime;
    const elapsedSeconds = previousTime === null
      ? 1 / 60
      : Math.min(0.05, Math.max(0, time - previousTime) / 1000);
    this.selectionAutoScrollLastTime = time;

    const beforeX = this.effectiveScrollLeft;
    const beforeY = this.viewportTop;
    this.setViewportLeft(beforeX + speed.x * elapsedSeconds);
    this.viewportTop = beforeY + speed.y * elapsedSeconds;
    const moved = this.effectiveScrollLeft !== beforeX || this.viewportTop !== beforeY;
    const extended = moved && this.extendDragSelection(pointer.clientX, pointer.clientY, true);

    if (moved) {
      this.updateSelectionOverlay();
      this.updateFindOverlay();
      this.scheduleRender();
      this.emitViewportChange();
      if (extended) this.emitSelectionChange();
    }

    if (!moved) {
      this.stopSelectionAutoScroll();
      return;
    }
    this.selectionAutoScrollFrame = this.hostWindow.requestAnimationFrame(
      (nextTime) => this.runSelectionAutoScroll(nextTime),
    );
  }

  private stopSelectionAutoScroll(): void {
    if (this.selectionAutoScrollFrame !== null) {
      this.hostWindow.cancelAnimationFrame(this.selectionAutoScrollFrame);
      this.selectionAutoScrollFrame = null;
    }
    this.selectionAutoScrollPointer = null;
    this.selectionAutoScrollLastTime = null;
  }

  private contextMenuTargetIsSelected(clientX: number, clientY: number): boolean {
    const selection = this.selectionState;
    if (!selection) return false;
    const header = this.getHeaderHit(clientX, clientY);
    if (header?.kind === 'corner') {
      return selection.areas.some((area) => area.kind === 'sheet');
    }
    if (header?.kind === 'row') {
      return selection.areas.some((area) => area.kind === 'sheet' ||
        (area.kind === 'rows' && header.row >= area.firstRow && header.row <= area.lastRow));
    }
    if (header?.kind === 'col') {
      return selection.areas.some((area) => area.kind === 'sheet' ||
        (area.kind === 'columns' &&
          header.col >= area.firstColumn && header.col <= area.lastColumn));
    }
    const cell = this.getCellAt(clientX, clientY);
    return cell !== null && selection.areas.some((area) => areaContainsCell(area, cell));
  }

  private resolveContextMenuContext(event: MouseEvent): Promise<XlsxSelectionContext | null> {
    if (this._destroyed) return Promise.resolve(null);
    const element = this.elementContextAt(event.clientX, event.clientY);
    if (element) {
      this.setElementContext(element);
    } else {
      this.setElementContext(null);
      if (!this.contextMenuTargetIsSelected(event.clientX, event.clientY)) {
        this.applyPointerSelection(event.clientX, event.clientY, false, false, -1, false);
      }
    }
    const context = this.getSelectionContext();
    return Promise.resolve(context ? structuredClone(context) : null);
  }

  private setupSelectionEvents(): void {
    // Distance (CSS px) beyond which a touch/pen pointerdown→pointerup is treated as a swipe (scroll), not a tap.
    const TAP_SLOP = 8;

    if (this.opts.onContextMenu) {
      this.surface.on('contextmenu', (event: MouseEvent) => {
        let context: Promise<XlsxSelectionContext | null> | undefined;
        this.opts.onContextMenu?.({
          originalEvent: event,
          getContext: () => context ??= this.resolveContextMenuContext(event),
        });
      });
    }

    this.surface.on('pointerdown', (e: PointerEvent) => {
      this.scrollHost.focus?.({ preventScroll: true });
      if (e.button !== 0) return;
      if (this.isSelecting && e.pointerId !== this.selectionPointerId) return;

      // Drag-to-resize a column/row from its header border (issue #567). Checked
      // before selection so grabbing the border never moves the cell selection.
      // Gated by the `resizable` option (default true); when off, a header-border
      // press falls through to normal selection behavior.
      const resize = (this.opts.resizable ?? true)
        ? this.getResizeTarget(e.clientX, e.clientY)
        : null;
      if (resize) {
        e.preventDefault();
        this.resizeDrag = { ...resize, pointerId: e.pointerId };
        this.scrollHost.setPointerCapture(e.pointerId);
        this.hideCommentPopup();
        return;
      }

      // List-validation dropdown arrow: if the press lands on the (display-only)
      // arrow button drawn on the active cell, toggle the value panel instead of
      // re-selecting the cell. The arrow's rect is in canvasArea space, so map
      // the client point through canvasArea's box.
      const ar = this.validationArrowRect;
      if (ar) {
        const { x: ax, y: ay } = this.surface.localPoint(e.clientX, e.clientY);
        if (ax >= ar.x && ax <= ar.x + ar.w && ay >= ar.y && ay <= ar.y + ar.h) {
          e.preventDefault();
          this.toggleValidationPanel();
          return;
        }
      }

      // A pointerdown on the native scrollbar must not move the cell
      // selection — dragging the thumb would otherwise select whatever cell
      // sits underneath it. Two scrollbar styles need different handling:
      // classic scrollbars reserve layout space, so the press lands in the
      // band between the content box (clientWidth/Height) and the border-box
      // edge and can be rejected exactly; OS overlay scrollbars (macOS
      // "show when scrolling") float over the content without affecting
      // client sizes, so a press near a scrollable edge is geometrically
      // indistinguishable from a cell click. For that case we defer the
      // selection to pointerup via the pendingTap path and cancel it when a
      // scroll event arrives first (the press was a thumb drag). A plain
      // click in the band still selects the cell on release.
      const hostRect = this.scrollHost.getBoundingClientRect();
      const localX = e.clientX - hostRect.left - this.scrollHost.clientLeft;
      const localY = e.clientY - hostRect.top - this.scrollHost.clientTop;
      if (localX >= this.scrollHost.clientWidth || localY >= this.scrollHost.clientHeight) {
        return; // classic scrollbar gutter
      }
      // Overlay scrollbar hit band (~15 CSS px on macOS / Windows 11).
      const OVERLAY_SCROLLBAR_BAND = 16;
      const inOverlayBand = this._nativeScrollbars && (
        (this.scrollHost.scrollWidth > this.scrollHost.clientWidth &&
          this.scrollHost.clientHeight - localY <= OVERLAY_SCROLLBAR_BAND) ||
        (this.scrollHost.scrollHeight > this.scrollHost.clientHeight &&
          this.scrollHost.clientWidth - localX <= OVERLAY_SCROLLBAR_BAND));

      const elementContext = this.elementContextAt(e.clientX, e.clientY);
      if (elementContext) {
        this.pendingTap = null;
        this.pendingClick = null;
        this.pendingElementClick = {
          x: e.clientX,
          y: e.clientY,
          pointerId: e.pointerId,
          context: elementContext,
        };
        return;
      }
      // A cell/header/empty-space press leaves object focus and returns the
      // authoritative context to the existing cell-selection state.
      this.setElementContext(null);

      // Touch / pen: defer selection until pointerup so swipe-to-scroll doesn't change the cell.
      // Mouse: select immediately to preserve drag-to-extend behavior.
      if (e.pointerType !== 'mouse' || inOverlayBand) {
        this.pendingTap = {
          x: e.clientX,
          y: e.clientY,
          shiftKey: e.shiftKey,
          additiveKey: e.ctrlKey || e.metaKey,
          pointerId: e.pointerId,
        };
        return;
      }

      // IX1 — remember the cell under a mouse press so a click (no drag) can
      // activate its hyperlink on release. Recorded before selection so a
      // shift-click extend still tracks the destination cell.
      const downCell = this.getCellAt(e.clientX, e.clientY);
      this.pendingClick = downCell
        ? { x: e.clientX, y: e.clientY, pointerId: e.pointerId, cell: downCell }
        : null;

      this.applyPointerSelection(
        e.clientX,
        e.clientY,
        e.shiftKey,
        e.ctrlKey || e.metaKey,
        e.pointerId,
        true,
      );
    });

    this.surface.on('pointermove', (e: PointerEvent) => {
      // Live column/row resize takes priority over every other pointer behavior.
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        e.preventDefault();
        this.applyResize(e.clientX, e.clientY);
        return;
      }

      // Resize-handle affordance: show the col/row-resize cursor when hovering a
      // header border (mouse only — touch/pen have no hover). Skipped mid-select
      // and when the `resizable` option (default true) is off, so no resize
      // cursor is shown when drag-resize is disabled.
      if (e.pointerType === 'mouse' && !this.isSelecting && (this.opts.resizable ?? true)) {
        const rt = this.getResizeTarget(e.clientX, e.clientY);
        this.scrollHost.style.cursor = rt ? (rt.kind === 'col' ? 'col-resize' : 'row-resize') : '';
        if (rt) {
          this.hideCommentPopup();
          return;
        }
      }

      // Cancel a pending tap once the pointer moves beyond the slop — the user is scrolling.
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) {
          this.pendingTap = null;
        }
      }

      // IX1 — a mouse press that turns into a drag (beyond the slop) is a
      // selection, not a hyperlink click: drop the pending activation.
      if (this.pendingClick && this.pendingClick.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingClick.x;
        const dy = e.clientY - this.pendingClick.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) {
          this.pendingClick = null;
        }
      }
      if (this.pendingElementClick?.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingElementClick.x;
        const dy = e.clientY - this.pendingElementClick.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) this.pendingElementClick = null;
      }

      // Comment hover popup (mouse only — touch/pen have no hover, so they get
      // the popup on selection instead, below). Suppressed while drag-selecting
      // so the popup doesn't fight the selection rect. A header hover hides it.
      if (e.pointerType === 'mouse' && !this.isSelecting) {
        const hovered = this.getCellAt(e.clientX, e.clientY);
        if (hovered) this.scheduleCommentPopup(hovered);
        else this.hideCommentPopup();
        // IX1 — pointer cursor over a hyperlinked cell. Reached only when the
        // pointer is NOT over a resize border (that path returns above), so the
        // resize cursor is never clobbered. Otherwise clear back to default.
        this.scrollHost.style.cursor =
          hovered && this.hyperlinkAtCell(hovered) ? 'pointer' : '';
      }

      if (!this.isSelecting || e.pointerId !== this.selectionPointerId) return;

      this.trackSelectionAutoScroll(e);
      if (!this.extendDragSelection(e.clientX, e.clientY, false)) return;

      this.updateSelectionOverlay();
      // Drag-select fires per pointermove; coalesce the canvas repaint (the
      // header-highlight bands the renderer draws) into one frame. The overlay
      // rect and the selection-change callback stay synchronous.
      this.scheduleRender();
      this.emitSelectionChange();
    });

    this.surface.on('pointerup', (e: PointerEvent) => {
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        if (this.resizeDrag.kind === 'col') this.refitAutoRowsAfterColumnResize();
        this.scrollHost.releasePointerCapture(e.pointerId);
        this.resizeDrag = null;
        return;
      }
      if (this.pendingElementClick?.pointerId === e.pointerId) {
        const pending = this.pendingElementClick;
        this.pendingElementClick = null;
        const dx = e.clientX - pending.x;
        const dy = e.clientY - pending.y;
        const current = dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP
          ? this.elementContextAt(e.clientX, e.clientY)
          : null;
        if (
          current &&
          current.sheetIndex === pending.context.sheetIndex &&
          current.elementType === pending.context.elementType &&
          current.elementIndex === pending.context.elementIndex &&
          current.shapeIndex === pending.context.shapeIndex
        ) this.setElementContext(current);
        return;
      }
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP) {
          this.applyPointerSelection(
            e.clientX,
            e.clientY,
            this.pendingTap.shiftKey,
            this.pendingTap.additiveKey,
            e.pointerId,
            false,
          );
          // Touch / pen have no hover, so surface the comment popup on a tap
          // (the active cell after the selection commit). Mouse uses hover.
          if (e.pointerType !== 'mouse' && this.activeCell) {
            const key = `${this.activeCell.row}:${this.activeCell.col}`;
            const comment = this.commentMap.get(key);
            if (comment) {
              this.hideCommentPopup();
              void this.renderCommentPopup(this.activeCell, comment)
                .catch((error) => this._reportRenderError(error));
            } else {
              this.hideCommentPopup();
            }
          }
          // IX1 — a touch/pen tap on a hyperlinked cell activates it.
          if (this.activeCell) this.dispatchHyperlink(this.activeCell);
        }
        this.pendingTap = null;
      }
      const endsSelectionDrag = e.pointerId === this.selectionPointerId;
      if (endsSelectionDrag) this.stopSelectionAutoScroll();
      // IX1 — a mouse click (press+release without a drag) on a hyperlinked cell
      // activates it. The release must still land on the same cell the press did.
      if (this.pendingClick && this.pendingClick.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingClick.x;
        const dy = e.clientY - this.pendingClick.y;
        const upCell = this.getCellAt(e.clientX, e.clientY);
        if (
          dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP &&
          upCell &&
          upCell.row === this.pendingClick.cell.row &&
          upCell.col === this.pendingClick.cell.col
        ) {
          this.dispatchHyperlink(this.pendingClick.cell);
        }
        this.pendingClick = null;
      }
      if (endsSelectionDrag) this.selectionController.endDrag(e.pointerId);
    });

    this.surface.on('pointercancel', (e: PointerEvent) => {
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        if (this.resizeDrag.kind === 'col') this.refitAutoRowsAfterColumnResize();
        this.resizeDrag = null;
      }
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        this.pendingTap = null;
      }
      if (this.pendingClick && this.pendingClick.pointerId === e.pointerId) {
        this.pendingClick = null;
      }
      if (this.pendingElementClick?.pointerId === e.pointerId) {
        this.pendingElementClick = null;
      }
      if (e.pointerId === this.selectionPointerId) {
        this.stopSelectionAutoScroll();
        this.selectionController.endDrag(e.pointerId);
      }
    });

    // Ctrl/⌘ + mouse wheel (and trackpad pinch, which the browser reports as a
    // ctrl-wheel) zooms the grid, matching Excel. preventDefault stops the
    // browser's own page zoom. A plain wheel still scrolls the grid natively.
    // The step is exponential in deltaY (see zoomStepScale) so a trackpad
    // pinch — a high-frequency stream of small-deltaY events — does not zoom
    // away; the total zoom tracks the gesture distance, not the event count.
    this.surface.on(
      'wheel',
      (e: WheelEvent) => {
        if (!(e.ctrlKey || e.metaKey)) {
          if (!this._nativeScrollbars) {
            e.preventDefault();
            const unit = e.deltaMode === WheelEvent.DOM_DELTA_LINE
              ? 16
              : e.deltaMode === WheelEvent.DOM_DELTA_PAGE
                ? Math.max(1, this.scrollHost.clientHeight)
                : 1;
            const horizontal = (e.shiftKey ? e.deltaY : e.deltaX) * unit;
            const vertical = (e.shiftKey ? 0 : e.deltaY) * unit;
            this.setViewportLeft(this.effectiveScrollLeft + horizontal);
            this.viewportTop += vertical;
            this.scheduleRender();
            this.updateSelectionOverlay();
            this.updateFindOverlay();
            this.emitViewportChange();
          }
          return;
        }
        e.preventDefault();
        if (e.deltaY === 0) return;
        // Pointer-anchored zoom: pivot on the cursor, not the top-left corner.
        // Record the pointer relative to the grid's top-left (canvasArea rect,
        // which the scrollHost overlays with inset:0) so `setScale` keeps the
        // cell under the cursor fixed. `scrollHost` and `canvasArea` share a rect.
        // A malformed event (no clientX/Y) yields a non-finite anchor; drop it so
        // `setScale` falls back to the historical START-anchored preservation.
        const { x: ax, y: ay } = this.surface.localPoint(e.clientX, e.clientY);
        this._pendingZoomAnchor =
          Number.isFinite(ax) && Number.isFinite(ay) ? { x: ax, y: ay } : null;
        this.setScale(zoomStepScale(this.viewport.scale, e.deltaY));
      },
      { passive: false },
    );

    this.surface.on('pointerleave', (event: PointerEvent) => {
      const next = event.relatedTarget as Node | null;
      if (next && this.commentPopup.contains(next)) return;
      this.hideCommentPopup();
    });

    // A canvas-backed sheet has no native focused cell. Establish the ordinary
    // A1 selection when its viewport receives keyboard focus, then reuse the
    // public selection contract for Arrow-key movement below.
    this.surface.on('focus', () => {
      if (this.currentWorksheet && !this.activeCell) this.setSelection('A1');
    });

    this.keydownHandler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
        if (e.defaultPrevented || e.isComposing) return;
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        if (target?.isContentEditable || tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
        e.preventDefault();
        void this.copySelection();
      } else if (
        !e.defaultPrevented && !e.isComposing &&
        !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey &&
        (e.key === 'ArrowUp' || e.key === 'ArrowDown' ||
          e.key === 'ArrowLeft' || e.key === 'ArrowRight')
      ) {
        const current = this.activeCell;
        const rowDelta = e.key === 'ArrowUp' ? -1 : e.key === 'ArrowDown' ? 1 : 0;
        const colDelta = e.key === 'ArrowLeft'
          ? (this.isRtl ? 1 : -1)
          : e.key === 'ArrowRight'
            ? (this.isRtl ? -1 : 1)
            : 0;
        const next = current ? {
          row: Math.max(1, Math.min(MAX_WORKSHEET_ROW, current.row + rowDelta)),
          col: Math.max(1, Math.min(MAX_WORKSHEET_COL, current.col + colDelta)),
        } : { row: 1, col: 1 };
        e.preventDefault();
        this.hideCommentPopup();
        const ref = formatA1(next.row, next.col);
        this.setSelection(ref);
        // Selection already schedules the paint. Reuse the ordinary viewport
        // geometry without starting a second immediate render for every key.
        this._scrollCellIntoView(next.row, next.col);
        this.updateSelectionOverlay();
        this.updateFindOverlay();
        this.emitViewportChange();
      } else if (e.key === 'Escape' && this.validationPanel.style.display !== 'none') {
        this.hideValidationPanel();
      } else if (e.key === 'Escape' && this.commentPopup.style.display !== 'none') {
        this.hideCommentPopup();
      } else if (
        e.key === 'Enter' && this.activeCell &&
        !e.defaultPrevented && !e.isComposing &&
        !e.ctrlKey && !e.metaKey && !e.altKey
      ) {
        const comment = this.commentMap.get(`${this.activeCell.row}:${this.activeCell.col}`);
        if (comment) {
          e.preventDefault();
          this.hideCommentPopup();
          void this.renderCommentPopup(this.activeCell, comment)
            .catch((error) => this._reportRenderError(error));
        }
      }
    };
    this.surface.on('keydown', this.keydownHandler);
  }

  private buildTabs(): void {
    if (this._mountKind === 'sheet') return;
    this.tabList.innerHTML = '';
    this.tabs = [];
    this.tabColors = this.workbook.tabColors;
    this.workbook.sheetNames.forEach((name, i) => {
      const btn = this.hostDocument.createElement('button');
      btn.textContent = name;
      btn.title = name;
      btn.style.cssText = this.tabCss(i, false);
      btn.addEventListener('click', () => {
        void this.goToSheet(i).catch((error) => this._reportRenderError(error));
      });
      this.tabList.appendChild(btn);
      this.tabs.push(btn);
    });
    this.updateNavButtons();
  }

  private makeNavButton(glyph: string, label: string, onClick: () => void): HTMLButtonElement {
    const btn = this.hostDocument.createElement('button');
    btn.textContent = glyph;
    btn.setAttribute('aria-label', label);
    btn.title = label;
    btn.classList.add('xlsx-tab-nav');
    btn.style.cssText = this.navButtonStyle(false);
    btn.addEventListener('click', onClick);
    return btn;
  }

  private navButtonStyle(disabled: boolean): string {
    // Plain triangle icons — no border / tab chrome. The background (incl. the
    // hover tint) lives in the injected `.xlsx-tab-nav` stylesheet so the inline
    // style does not shadow the `:hover` rule.
    const base =
      `flex:1;height:100%;padding:0;` +
      `display:flex;align-items:center;justify-content:center;` +
      `border:none;color:var(--ooxml-xlsx-chrome-text-muted,#666);font-size:9px;line-height:1;` +
      `box-sizing:border-box;outline:none;`;
    return disabled
      ? base + `opacity:0.3;cursor:default;pointer-events:none;`
      : base + `cursor:pointer;`;
  }

  private scrollTabs(dir: -1 | 1): void {
    const strip = this.tabStrip;
    const viewLeft = strip.scrollLeft;
    const viewRight = viewLeft + strip.clientWidth;
    let target: number | null = null;
    if (dir === 1) {
      // Nearest tab clipped on the physical right; align its right edge.
      let nearestRight = Number.POSITIVE_INFINITY;
      for (const tab of this.tabs) {
        const right = tab.offsetLeft + tab.offsetWidth;
        if (right > viewRight + 1) nearestRight = Math.min(nearestRight, right);
      }
      if (Number.isFinite(nearestRight)) target = nearestRight - strip.clientWidth;
    } else {
      // Nearest tab clipped on the physical left; align its left edge. Search
      // by geometry, not DOM order, because RTL reverses the visual tab row.
      let nearestLeft = Number.NEGATIVE_INFINITY;
      for (const tab of this.tabs) {
        const left = tab.offsetLeft;
        if (left < viewLeft - 1) nearestLeft = Math.max(nearestLeft, left);
      }
      if (Number.isFinite(nearestLeft)) target = nearestLeft;
    }
    if (target !== null) {
      // Instant (not smooth) so the disabled state is consistent the moment the
      // click resolves — keeps the interaction deterministic to drive/test.
      strip.scrollLeft = Math.max(0, Math.min(target, strip.scrollWidth - strip.clientWidth));
    }
    this.updateNavButtons();
  }

  private updateNavButtons(): void {
    if (this._mountKind === 'sheet') return;
    const strip = this.tabStrip;
    const atStart = strip.scrollLeft <= 0;
    const atEnd = strip.scrollLeft + strip.clientWidth >= strip.scrollWidth - 1;
    // No overflow => scrollWidth ≈ clientWidth => both ends true => both disabled.
    this.navPrev.style.cssText = this.navButtonStyle(atStart);
    this.navNext.style.cssText = this.navButtonStyle(atEnd);
  }

  private updateTabActive(index: number): void {
    this.tabs.forEach((btn, i) => {
      btn.style.cssText = this.tabCss(i, i === index);
    });
    // Keep the active tab visible by scrolling the tab strip HORIZONTALLY only.
    // `scrollIntoView` walks every scrollable ancestor, so it also scrolls the
    // page vertically — on first load that jumped the whole page down to the
    // tab bar (the active sheet is set during load). Adjust the strip's
    // scrollLeft directly so the page never moves.
    // `offsetParent === null` for a `display:none` tab (a hidden sheet reached
    // by an explicit goToSheet in 'skip' mode). Its getBoundingClientRect is all
    // zeros, which would spuriously scroll the strip — skip the scroll for it.
    const tab = this.tabs[index];
    if (tab && tab.offsetParent !== null) {
      const strip = this.tabStrip;
      const tabRect = tab.getBoundingClientRect();
      const stripRect = strip.getBoundingClientRect();
      if (tabRect.left < stripRect.left) {
        strip.scrollLeft -= stripRect.left - tabRect.left;
      } else if (tabRect.right > stripRect.right) {
        strip.scrollLeft += tabRect.right - stripRect.right;
      }
    }
    this.updateNavButtons();
  }

  private tabStyle(active: boolean, tabColor?: string | null): string {
    // Active tab renders taller than inactive so the selected sheet draws the
    // eye. Tabs align to flex-end, so shorter inactive tabs sit lower and the
    // active tab sticks up. Font size also bumps a hair on active.
    const activeH = TAB_BAR_H - 2;
    const inactiveH = TAB_BAR_H - 5;
    const base =
      `display:inline-block;flex:none;padding:0 14px;position:relative;` +
      `border:1px solid var(--ooxml-xlsx-chrome-border,#c8ccd0);border-bottom:none;` +
      `border-radius:3px 3px 0 0;` +
      `cursor:pointer;white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;` +
      `outline:none;box-sizing:border-box;`;
    // `<sheetPr><tabColor>` renders as a color bar along the tab's bottom edge
    // (Excel's "sheet tab color" treatment), drawn as an inset bottom shadow so
    // it doesn't fight the tab's own border/background. The active tab keeps a
    // thinner bar since its bottom merges into the white sheet body.
    const bar = tabColor
      ? `box-shadow:inset 0 -${active ? 2 : 3}px 0 0 ${tabColor};`
      : '';
    return active
      ? base +
        `height:${activeH}px;font-size:13px;` +
        `background:var(--ooxml-xlsx-chrome-surface,#fff);` +
        `color:var(--ooxml-xlsx-chrome-text,#000);` +
        `border-bottom:1px solid var(--ooxml-xlsx-chrome-surface,#fff);` +
        `font-weight:600;top:1px;` +
        bar
      : base +
        `height:${inactiveH}px;font-size:11px;` +
        `background:var(--ooxml-xlsx-chrome-surface-muted,#e0e0e0);` +
        `color:var(--ooxml-xlsx-chrome-text-muted,#555);` +
        bar;
  }

  /**
   * Full inline style for the tab of sheet `i`, honoring the hidden-sheet mode:
   * `'skip'` hides the tab of a hidden/veryHidden sheet (`display:none`); `'dim'`
   * greys it but leaves it clickable; `'show'` styles every tab normally. Used
   * by both buildTabs and updateTabActive so navigation never wipes the styling.
   */
  private tabCss(i: number, active: boolean): string {
    let css = this.tabStyle(active, this.tabColors[i]);
    if (this._hiddenSheetMode !== 'show' && this.wb?.isHidden(i)) {
      css += this._hiddenSheetMode === 'skip' ? 'display:none;' : `opacity:${HIDDEN_TAB_DIM_OPACITY};`;
    }
    return css;
  }

  /** Excel-style zoom control pinned to the footer's logical end:
   *  `−  [────slider────]  +  100%`. Live-updates the cell scale on input. */
  private buildZoomControl(): HTMLDivElement {
    const zoomMin = this.opts.zoomMin ?? 0.1;
    const zoomMax = this.opts.zoomMax ?? 4;
    const cur = this.viewport.scale;

    const wrap = this.hostDocument.createElement('div');
    wrap.style.cssText =
      `display:flex;align-items:center;flex-shrink:0;gap:2px;` +
      `padding:0 10px;height:100%;` +
      `color:var(--ooxml-xlsx-chrome-text-muted,#555);font-size:12px;user-select:none;`;

    // The steppers walk the shared IX9 zoom ladder (ZOOM_STEP_LADDER via
    // zoomIn/zoomOut) so the built-in chrome and a host's own buttons wired to
    // the ZoomableViewer contract land on identical scales (issue #842).
    // Pre-IX9 these stepped ±0.1 linearly.
    const mkBtn = (glyph: string, label: string, step: () => void): HTMLButtonElement => {
      const b = this.hostDocument.createElement('button');
      b.type = 'button';
      b.textContent = glyph;
      b.setAttribute('aria-label', label);
      b.title = label;
      b.style.cssText =
        `width:18px;height:18px;padding:0;border:none;background:transparent;` +
        `color:var(--ooxml-xlsx-chrome-text-muted,#555);` +
        `font-size:14px;line-height:1;cursor:pointer;border-radius:3px;`;
      b.addEventListener('click', step);
      return b;
    };

    // The slider works in "position" units [0,100]; 50 is dead-center and maps
    // to 100% so each half is its own linear segment (zoomMin→1 on the left,
    // 1→zoomMax on the right), mirroring Excel's status-bar zoom where 100% sits
    // in the middle even though the range (10%–400%) is asymmetric.
    const slider = this.hostDocument.createElement('input');
    slider.type = 'range';
    slider.min = '0';
    slider.max = '100';
    slider.step = 'any';
    slider.value = String(this.zoomScaleToPos(cur, zoomMin, zoomMax));
    slider.setAttribute('aria-label', 'Zoom');
    slider.title = 'Zoom';
    slider.classList.add('xlsx-zoom-slider');
    slider.style.cssText = `width:90px;cursor:pointer;`;
    slider.addEventListener('input', () =>
      this.setScale(this.zoomPosToScale(Number(slider.value), zoomMin, zoomMax)),
    );

    const label = this.hostDocument.createElement('span');
    label.textContent = `${Math.round(cur * 100)}%`;
    label.style.cssText = `min-width:42px;margin-left:6px;text-align:right;font-variant-numeric:tabular-nums;`;

    wrap.appendChild(mkBtn('−', 'Zoom out', () => this.zoomOut()));
    wrap.appendChild(slider);
    wrap.appendChild(mkBtn('+', 'Zoom in', () => this.zoomIn()));
    wrap.appendChild(label);

    this.zoomSlider = slider;
    this.zoomLabel = label;
    return wrap;
  }

  /** Map a slider position [0,100] to a scale factor. 50 → 1.0 (100%), with a
   *  separate linear segment on each side so the center is always 100%. */
  private zoomPosToScale(pos: number, min: number, max: number): number {
    return pos <= 50
      ? min + (pos / 50) * (1 - min)
      : 1 + ((pos - 50) / 50) * (max - 1);
  }

  /** Inverse of {@link zoomPosToScale}: scale factor → slider position [0,100]. */
  private zoomScaleToPos(scale: number, min: number, max: number): number {
    const clamped = Math.min(max, Math.max(min, scale));
    return clamped <= 1
      ? ((clamped - min) / (1 - min)) * 50
      : 50 + ((clamped - 1) / (max - 1)) * 50;
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

    if (this.zoomSlider) this.zoomSlider.value = String(this.zoomScaleToPos(next, zoomMin, zoomMax));
    if (this.zoomLabel) this.zoomLabel.textContent = `${pct}%`;

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
    this.updateNavButtons();
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
   *  the row/column header plus every used column width / row height. Mirrors
   *  {@link updateSpacerSize} at cs=1 (same used-range detection) so the fit
   *  targets exactly the region the spacer/scroll extent covers. */
  private _naturalContentExtent(ws: Worksheet): { width: number; height: number } {
    const { maxRow, maxCol } = worksheetContentBounds(ws);
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
    let { maxRow, maxCol } = worksheetContentBounds(ws);
    maxRow += 30;
    maxCol += 10;

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
    try {
      await this._renderCurrentSheet(generation);
    } catch (err) {
      if (!this.renderDispatcher.isCurrent(generation)) return;
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
    const viewerRenderOpts = withViewerRenderContext(
      sizeProjection ? { ...renderOpts, sizeOverrides: sizeProjection.overrides } : renderOpts,
      getGridGeometryForWorksheet(ws).maximumDigitWidth,
      {
        worksheet: ws,
        projection: sizeProjection
          ? { id: this.projectionId, revision: sizeProjection.revision, autoRowHeightsPrepared: true }
          : undefined,
      },
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
   * All document-level listeners are detached — the keydown handler here, and
   * the validation-panel outside-click handler via {@link hideValidationPanel}.
   * Listeners on elements inside the wrapper (scrollHost, tabs, …) need no
   * explicit removal: removing the subtree makes them unreachable and eligible
   * for GC. Safe to call more than once.
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
    if (this.selectionContextNotificationFrame !== null) {
      this.hostWindow.cancelAnimationFrame(this.selectionContextNotificationFrame);
      this.selectionContextNotificationFrame = null;
    }
    this.selectionContextNotificationMicrotask = false;
    this.stopSelectionAutoScroll();
    this.sheetRequestGeneration++;
    this.resizeObserver?.disconnect();
    this.chromeStyleObserver?.disconnect();
    this.chromeStyleObserver = null;
    if (this.chromeSchemeMedia && this.chromeSchemeListener) {
      this.chromeSchemeMedia.removeEventListener?.('change', this.chromeSchemeListener);
    }
    this.chromeSchemeMedia = null;
    this.chromeSchemeListener = null;
    this.commentPopupResizeObserver?.disconnect();
    this.commentPopupResizeObserver = null;
    this.renderDispatcher.destroy();
    this.surface.destroy();
    this.hideCommentPopup();
    this.hideValidationPanel();
    // IX2 — drop the find state (matches + cursor) so a stale
    // findNext()/findPrev() after teardown returns null instead of a match
    // pointing into a dead viewer (same fix as DocxViewer/PptxViewer.destroy).
    this._find.invalidate();
    this.releaseHostFonts();
    const releaseProjection = this.wb?.[releaseXlsxViewerProjection];
    if (typeof releaseProjection === 'function') {
      releaseProjection.call(this.wb, this.projectionId);
    }
    this.currentWorksheet = null;
    this.currentSourceComments = [];
    this.sourceCommentMap.clear();
    this.elementContext = null;
    this.pendingElementClick = null;
    this.selectionController.reset();
    this.lastNotifiedSelectionState = null;
    this.finishSelectionNotificationChain();
    this.acquisition.destroy();
    // Remove the whole UI subtree so the container is empty again. This also
    // detaches every listener bound to elements within it (scrollHost pointer/
    // wheel handlers, tab clicks, zoom slider) without per-element cleanup.
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

  async load(source: string | ArrayBuffer): Promise<void> {
    this.assertOpen();
    try {
      await this.engine.load(source);
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
