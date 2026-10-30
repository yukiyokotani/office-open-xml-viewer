import type { FindHighlightColors, HyperlinkTarget, ViewerContextMenuEvent } from '@silurus/ooxml-core';
import type { LoadOptions } from './document';
import type { RenderPageOptions } from './types';
import type { DocxSelectionContext } from './selection-context';
import type { DocxCommentsOptions } from './comment-margin';

/**
 * Options for {@link DocxScrollViewer}. Extends `RenderPageOptions` (per-page
 * render knobs, minus `onTextRun`) and `LoadOptions` (parse/worker knobs). See
 * design §8.1.
 *
 * `onTextRun` is omitted deliberately: the viewer drives it internally per
 * mounted slot to build the optional per-page selection overlay (gated by
 * `enableTextSelection`), so exposing it here would let a caller's callback be
 * silently overridden.
 */
export interface DocxScrollViewerOptions extends Omit<RenderPageOptions, 'onTextRun'>, LoadOptions {
  /** Base fit width in CSS px → base zoom scale. Default: the container's width
   *  at first non-zero layout (design §7/§11 zero-width deferral). */
  width?: number;
  /** Vertical gap (px) between consecutive pages. Default 16. */
  gap?: number;
  /** Desk padding (px) ABOVE the FIRST page — the margin a PDF reader leaves
   *  between the top of the scroll surface and the first sheet. Default: `gap`
   *  (uniform desk rhythm — the first page sits the same distance from the top as
   *  pages sit from each other). Pass `0` for a flush-top layout. */
  paddingTop?: number;
  /** Desk padding (px) BELOW the LAST page — the margin below the final sheet.
   *  Default: `gap`. Pass `0` for a flush-bottom layout. */
  paddingBottom?: number;
  /** Desk gutter (px) to the LEFT of the pages — the horizontal margin between the
   *  left edge of the scroll surface and a page sitting flush-left (i.e. once
   *  zoomed wide enough that centering no longer applies). Default: `gap` (uniform
   *  desk rhythm — the horizontal gutters match the vertical ones). It also shrinks
   *  the container-derived FIT width so a page sits inside the gutters at 100%
   *  (an EXPLICIT `opts.width` is the page's CSS-width contract and is NOT reduced;
   *  the gutters still apply around placement). Pass `0` for a flush-left layout. */
  paddingLeft?: number;
  /** Desk gutter (px) to the RIGHT of the pages. Default: `gap`. Shrinks the
   *  container-derived fit width symmetrically with `paddingLeft`. Pass `0` for a
   *  flush-right layout. */
  paddingRight?: number;
  /** Pages kept mounted beyond the viewport on each side. Default 1. */
  overscan?: number;
  /**
   * Paint the document's opening pages as soon as they are laid out, instead of
   * waiting for the whole document.
   *
   * A large document otherwise shows nothing until every page has been
   * paginated. With this on, the viewer mounts the first pages within a few
   * hundred milliseconds regardless of document length, the scrollbar grows as
   * the rest of the layout arrives, and the viewer relays out once it lands.
   * Published pages are provisional: later header/footer, field, anchored-object,
   * or section convergence can still replace them. Mounted pages repaint when
   * the authoritative layout lands.
   *
   * `pageCount` therefore starts small and grows; {@link findText} waits for the
   * full layout internally. Ignored by `fromDocument`, whose document is already
   * loaded. Works in both render modes and in either tracked-changes view; in
   * `mode: 'worker'` it additionally keeps the remaining pagination off the
   * main thread.
   */
  progressiveLayout?: boolean;
  /**
   * Lay the document out in slices, keeping the thread responsive while a large
   * document paginates. `load()` still resolves only once layout is complete —
   * use {@link progressiveLayout} to paint before then.
   */
  sliceLayout?: boolean;
  /** Per-page transparent text-selection overlay. IX6 — works in BOTH render
   *  modes: in worker mode the per-run geometry is collected off-thread and
   *  shipped back beside the page bitmap, so the overlay is populated identically
   *  to main mode (no more empty overlay / one-time warning). */
  enableTextSelection?: boolean;
  /** Show the built-in read-only comments. Pass options to configure them. Default false. */
  comments?: boolean | DocxCommentsOptions;
  /**
   * Enable read-only selection of mounted pictures, charts, and shapes. The
   * selected object exposes element context and receives a non-editable outline.
   */
  enableElementSelection?: boolean;
  /** Emits bounded, detached text or element context suitable for read-only AI/MCP use. */
  onSelectionContextChange?: (context: DocxSelectionContext | null) => void;
  /**
   * Called synchronously for a browser `contextmenu` event. The original event
   * can suppress the native menu; `getContext()` resolves the text or element
   * context established at the event target.
   */
  onContextMenu?: (event: ViewerContextMenuEvent<DocxSelectionContext>) => void;
  /** CSS backgrounds for ordinary and active in-document search matches. */
  findHighlightColors?: FindHighlightColors;
  /** Minimum zoom scale (px-per-pt multiplier floor). A smaller width-fit base
   * remains reachable as the effective minimum. Default 0.1. */
  zoomMin?: number;
  /** Maximum zoom scale. Default 4. */
  zoomMax?: number;
  /** Enable `Ctrl`/`Cmd`+wheel zoom. Default true. */
  enableZoom?: boolean;
  /**
   * Re-fit the document to the container width when the container is resized.
   * Default true. Set false to preserve the current absolute scale, including
   * an explicit pre-load `setScale(1)`, independently of the viewport width.
   * Explicit `fitWidth()` and `fitPage()` calls remain available.
   */
  refitOnResize?: boolean;
  /**
   * CSS `background` shorthand for the scroll surface (the "desk") visible
   * behind and between pages — the gray a PDF reader paints around the sheet.
   * Applied to the viewer-owned scroll host. The pages themselves are always
   * drawn on the document's own white canvas and are unaffected. Default
   * `undefined`: the scroll surface stays transparent so the host container's
   * background shows through (non-breaking).
   */
  background?: string;
  /**
   * CSS `box-shadow` painted on every page CANVAS (not the wrapper — the
   * text-selection overlay must not cast its own shadow). The soft drop shadow a
   * PDF reader leaves under each sheet.
   *
   * - Default (`undefined`): `'0 1px 3px rgba(0,0,0,0.2)'` — the recipe look, so
   *   the scroll viewer reproduces the Examples appearance with zero config.
   * - `false`: NO shadow (flat pages).
   * - A custom string is applied verbatim. A spread-only ring such as
   *   `'0 0 0 1px #c8ccd0'` gives a crisp 1px BORDER look — and because
   *   `box-shadow` never affects layout (unlike `border`, which would grow the
   *   box and shift every offset), a border and a drop shadow are the SAME knob
   *   here rather than two competing options.
   */
  pageShadow?: string | false;
  /** Fires when the top-most visible page OR the document's page count changes.
   *  `topIndex` from `computeVisibleRange` (the first page intersecting the
   *  viewport top, EXCLUDING overscan).
   *
   *  `layoutComplete` is false while progressive layout is still running, and
   *  `total` is then the pages laid out SO FAR, not the document's total — a
   *  "page X of Y" indicator should mark it provisional (Word shows an
   *  unsettled count the same way during background repagination). The count
   *  is watched as well as the index precisely so that indicator updates when
   *  the rest of the document arrives without the user scrolling. */
  onVisiblePageChange?: (
    topIndex: number,
    total: number,
    layoutComplete: boolean,
  ) => void;
  /** IX9 — fires whenever the zoom factor actually changes (`1` = 100% = a page
   *  at its natural pt→px size): from {@link DocxScrollViewer.setScale},
   *  `zoomIn`/`zoomOut`, `fitWidth`/`fitPage`, a Ctrl/⌘+wheel gesture, or a
   *  container-resize re-fit (when `refitOnResize` is enabled). Named
   *  `onScaleChange` to match the single-canvas viewers so all five share one
   *  notification shape. */
  onScaleChange?: (scale: number) => void;
  /** IX1 (design decision — NOT user-confirmed, integrator may veto). Called when
   *  a hyperlink run is clicked. When omitted, the default is: external → open in a
   *  new tab via core `openExternalHyperlink` (sanitised, noopener,noreferrer);
   *  internal → jump to the page whose text contains the bookmark (best-effort). */
  onHyperlinkClick?: (target: HyperlinkTarget) => void;
  /** IX1 — master switch for hyperlink interactivity. Default `true`. When
   *  `false`, the hyperlink machinery is not wired at all: no overlay hit region
   *  is installed for link runs, so there is no pointer cursor, no title tooltip,
   *  no default navigation (external new-tab / internal bookmark jump), and
   *  `onHyperlinkClick` is never called. Links still render exactly as authored
   *  but are inert, like plain text. */
  enableHyperlinks?: boolean;
  /** Receives asynchronous Viewer-managed failures that cannot be observed by
   *  awaiting the method that started them. `load()` failures always reject and
   *  are not also delivered here. Virtualized per-slot render failures (both
   *  main `renderPage` and worker `renderPageToBitmap` rejections) invoke it; a
   *  failed page is left blank rather than crashing the loop. Without an
   *  `onError`, render failures are logged via
   *  `console.error` so they are never fully silent. Stable cases can be
   *  narrowed with `OoxmlError`, `OoxmlResourceLimitError`, or
   *  `OoxmlDecodedImageLimitError` re-exported by this package. Other failures
   *  remain `Error` values; a `code` of `parser-crashed` identifies a recognized
   *  WASM trap, not a reliably classified OOM. */
  onError?: (err: Error) => void;
  /**
   * Called once per loaded document when a package part failed to parse and a
   * placeholder page is shown instead (#1974). Delivered before `load()`
   * resolves. Never routed to `onError`; `load()` still resolves. The same
   * records are available from `DocxScrollViewer.partFailures`.
   */
  onPartFailure?: (failure: Readonly<import('./part-failures').DocxPartFailure>) => void;
}
