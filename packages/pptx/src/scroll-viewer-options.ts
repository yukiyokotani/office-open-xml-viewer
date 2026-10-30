import type { FindHighlightColors, HyperlinkTarget, ViewerContextMenuEvent } from '@silurus/ooxml-core';
import type { LoadOptions, RenderSlideOptions } from './presentation';
import type { PptxSelectionContext } from './element-selection';
import type { PptxCommentsOptions } from './comment-margin';

/**
 * Options for {@link PptxScrollViewer}. Only the `width` and `dpr` per-slide
 * render knobs apply to this virtualized Viewer; it owns text-run collection,
 * media controls, and hidden-slide dimming itself.
 */
export interface PptxScrollViewerOptions extends Pick<RenderSlideOptions, 'width' | 'dpr' | 'imageResources'>, LoadOptions {
  /** Base fit width in CSS px → base zoom scale. Default: the container's width
   *  at first non-zero layout (design §7/§11 zero-width deferral). */
  width?: number;
  /** Vertical gap (px) between consecutive slides. Default 16. */
  gap?: number;
  /** Desk padding (px) ABOVE the FIRST slide — the margin a presentation viewer
   *  leaves between the top of the scroll surface and the first slide. Default:
   *  `gap` (uniform desk rhythm — the first slide sits the same distance from the
   *  top as slides sit from each other). Pass `0` for a flush-top layout. */
  paddingTop?: number;
  /** Desk padding (px) BELOW the LAST slide — the margin below the final slide.
   *  Default: `gap`. Pass `0` for a flush-bottom layout. */
  paddingBottom?: number;
  /** Desk gutter (px) to the LEFT of the slides — the horizontal margin between
   *  the left edge of the scroll surface and a slide sitting flush-left (i.e. once
   *  zoomed wide enough that centering no longer applies). Default: `gap` (uniform
   *  desk rhythm — the horizontal gutters match the vertical ones). It also shrinks
   *  the container-derived FIT width so a slide sits inside the gutters at 100%
   *  (an EXPLICIT `opts.width` is the slide's CSS-width contract and is NOT reduced;
   *  the gutters still apply around placement). Pass `0` for a flush-left layout. */
  paddingLeft?: number;
  /** Desk gutter (px) to the RIGHT of the slides. Default: `gap`. Shrinks the
   *  container-derived fit width symmetrically with `paddingLeft`. Pass `0` for a
   *  flush-right layout. */
  paddingRight?: number;
  /** Slides kept mounted beyond the viewport on each side. Default 1. */
  overscan?: number;
  /** Per-slide transparent text-selection overlay. IX6 — works in BOTH render
   *  modes: in worker mode the per-run geometry is collected off-thread and
   *  shipped back beside the slide bitmap, so the overlay is populated identically
   *  to main mode (no more empty overlay / one-time warning). */
  enableTextSelection?: boolean;
  /** Show the built-in read-only comments. Pass options to configure them. Default false. */
  comments?: boolean | PptxCommentsOptions;
  /** Enable read-only slide-element selection with a non-editable outline. Default false. */
  enableElementSelection?: boolean;
  /** Straight-line hit tolerance in CSS pixels. Default 6. */
  elementHitTolerance?: number;
  /** Emits bounded, detached text or element context for read-only AI/MCP use. */
  onSelectionContextChange?: (context: PptxSelectionContext | null) => void;
  /**
   * Called synchronously for a browser `contextmenu` event. The original event
   * can suppress the native menu; `getContext()` resolves the text or element
   * context established at the event target.
   */
  onContextMenu?: (event: ViewerContextMenuEvent<PptxSelectionContext>) => void;
  /** CSS backgrounds for ordinary and active in-document search matches. */
  findHighlightColors?: FindHighlightColors;
  /**
   * Enable interactive audio/video playback. When true, mounted slides render
   * through {@link PptxPresentation.presentSlide} only while they are inside the
   * real viewport plus {@link mediaOverscan}. Other mounted slides retain static
   * canvases and selectable text without allocating media blobs or RAF loops.
   * Default false.
   */
  enableMediaPlayback?: boolean;
  /**
   * Slides beyond the real viewport that may keep interactive media handles.
   * Independent from {@link overscan}, so an integration may mount every text
   * overlay for browser-native Find while media resources stay bounded. Default 1.
   */
  mediaOverscan?: number;
  /** Minimum zoom scale — a DIMENSIONLESS multiplier over the 96-dpi natural
   *  slide size (10% = 0.1), matching `DocxScrollViewer`. A smaller width-fit
   *  base remains reachable as the effective minimum. Default 0.1. */
  zoomMin?: number;
  /** Maximum zoom scale (dimensionless multiplier, 400% = 4). Default 4. */
  zoomMax?: number;
  /** Enable `Ctrl`/`Cmd`+wheel zoom. Default true. */
  enableZoom?: boolean;
  /**
   * Re-fit the presentation to the container width when the container is
   * resized. Default true. Set false to preserve the current absolute scale,
   * including an explicit pre-load `setScale(1)`, independently of the viewport
   * width. Explicit `fitWidth()` and `fitPage()` calls remain available.
   */
  refitOnResize?: boolean;
  /**
   * CSS `background` shorthand for the scroll surface (the "desk") visible
   * behind and between slides — the gray a presentation viewer paints around the
   * slide. Applied to the viewer-owned scroll host. The slides themselves are
   * always drawn on their own white canvas and are unaffected. Default
   * `undefined`: the scroll surface stays transparent so the host container's
   * background shows through (non-breaking).
   */
  background?: string;
  /**
   * CSS `box-shadow` painted on every slide CANVAS (not the wrapper — the
   * text-selection overlay must not cast its own shadow). The soft drop shadow a
   * presentation viewer leaves under each slide.
   *
   * - Default (`undefined`): `'0 1px 3px rgba(0,0,0,0.2)'` — the recipe look, so
   *   the scroll viewer reproduces the Examples appearance with zero config.
   * - `false`: NO shadow (flat slides).
   * - A custom string is applied verbatim. A spread-only ring such as
   *   `'0 0 0 1px #c8ccd0'` gives a crisp 1px BORDER look — and because
   *   `box-shadow` never affects layout (unlike `border`, which would grow the
   *   box and shift every offset), a border and a drop shadow are the SAME knob
   *   here rather than two competing options.
   */
  pageShadow?: string | false;
  /** Fires when the top-most visible slide changes. `topIndex` from
   *  `computeVisibleRange` (the first slide intersecting the viewport top,
   *  EXCLUDING overscan). */
  onVisibleSlideChange?: (topIndex: number, total: number, layoutComplete: boolean) => void;
  /** IX9 — fires whenever the zoom factor actually changes (`1` = 100% = a slide
   *  at its natural EMU→px size): from {@link PptxScrollViewer.setScale},
   *  `zoomIn`/`zoomOut`, `fitWidth`/`fitPage`, a Ctrl/⌘+wheel gesture, or a
   *  container-resize re-fit (when `refitOnResize` is enabled). Named
   *  `onScaleChange` to match the single-canvas viewers so all five share one
   *  notification shape. */
  onScaleChange?: (scale: number) => void;
  /** Receives asynchronous Viewer-managed failures that cannot be observed by
   *  awaiting the method that started them. `load()` failures always reject and
   *  are not also delivered here. Virtualized per-slot render failures (both
   *  main `renderSlide` and worker `renderSlideToBitmap` rejections) and
   *  embedded-media fetch/decode/playback failures invoke it. A failed slide is
   *  left blank rather than crashing the loop.
   *  Without an `onError`, failures are logged via `console.error` so they are
   *  never fully silent. Stable cases can be narrowed with `OoxmlError`,
   *  `OoxmlResourceLimitError`, or `OoxmlDecodedImageLimitError` re-exported by
   *  this package. Other failures remain `Error` values; a `code` of
   *  `parser-crashed` identifies a recognized WASM trap, not a reliably
   *  classified OOM. */
  onError?: (err: Error) => void;
  /**
   * Replacement text for the "Slide N could not be displayed" placeholder of a
   * slide whose part failed to parse (#1974). Static text, or a function of
   * the failure evaluated on the main thread (also in worker mode). `detail:
   * null` hides the detail line. Ignored for healthy slides.
   */
  errorPlaceholder?: import('@silurus/ooxml-core').ErrorPlaceholderOption<import('./part-failures').PptxPartFailure>;
  /**
   * Called once per loaded presentation for each slide whose part failed to
   * parse (#1974). Failures known at load are delivered before `load()`
   * resolves; with `progressiveLayout`, later ones when the slide is reached.
   * Never routed to `onError`; `load()` still resolves. The same records are
   * available from `partFailures`.
   */
  onPartFailure?: (failure: Readonly<import('./part-failures').PptxPartFailure>) => void;
  /**
   * IX1 (design decision — NOT user-confirmed, integrator may veto). Fires on a
   * hyperlink click in any mounted slide's text overlay (requires
   * {@link enableTextSelection}). Default when omitted: external →
   * {@link openExternalHyperlink} (new tab, sanitised, noopener); internal
   * slide-jump → {@link scrollToSlide} once the action resolves to a slide index
   * via {@link PptxPresentation.resolveInternalTarget} (a jump that resolves to
   * no reachable slide is a safe no-op). When provided, the viewer calls this
   * instead and takes NO default action.
   */
  onHyperlinkClick?: (target: HyperlinkTarget) => void;
  /** IX1 — master switch for hyperlink interactivity. Default `true`. When
   *  `false`, the hyperlink machinery is not wired at all: the overlay's link
   *  spans are non-interactive, so there is no pointer cursor, no title tooltip,
   *  no default navigation (external new-tab / internal slide jump), and
   *  `onHyperlinkClick` is never called. Links still render exactly as authored
   *  but are inert, like plain text. */
  enableHyperlinks?: boolean;
}
