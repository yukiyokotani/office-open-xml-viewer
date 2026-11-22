import { clippedBoxVisible, OverlayClipFrames, runPartVisible, runStraddlesClip } from './run-clip.js';
import { countSpacingCharacters } from './character-spacing.js';
import type { DocxTextRunInfo } from './renderer';
import { overlayPercent, type HyperlinkTarget } from '@silurus/ooxml-core';
import { tateChuYokoOverlayScale } from './tate-chu-yoko-overlay';
import { runOverlayMeasure } from './run-overlay-measure';
import { setSelectionData } from '@silurus/ooxml-core/internal/selection-data';

/**
 * Build the transparent text-selection overlay for a rendered docx page: one
 * absolutely-positioned, color-transparent `<span>` per {@link DocxTextRunInfo}
 * (emitted by `renderPage`'s `onTextRun`), so the browser's native selection
 * lands on the drawn glyphs. Extracted verbatim from `DocxViewer._buildTextLayer`
 * so both the pager (DocxViewer) and the continuous-scroll viewer (DocxScrollViewer)
 * share one implementation; also public API for integrators building their own
 * overlay (design §10). IX6 — usable in BOTH render modes: worker mode collects
 * the same `DocxTextRunInfo[]` off-thread and ships it back beside the bitmap, so
 * the overlay is built from identical geometry regardless of thread.
 *
 * Every span is positioned as a PERCENTAGE of `cssWidth`/`cssHeight` (the page's
 * intended CSS-px box), never literal px, and the container's own width/height are
 * left untouched (the caller sizes it `width:100%;height:100%`). This lets the
 * overlay track the canvas's ACTUAL rendered box even when a consumer scales the
 * canvas down with external CSS (`width:100%!important; height:auto`): the
 * `display:inline-block` wrapper shrinks with the canvas, the `100%` container
 * follows, and every `%`-placed span scales with it, so nothing overflows the
 * wrapper into an ancestor's scroll area.
 *
 * @param layer            the overlay div (sized `width:100%;height:100%` by the caller).
 * @param runs             per-run geometry from `renderPage({ onTextRun })`.
 * @param cssWidth         the page's intended CSS width (px, number) — the %
 *                         denominator for the x axis.
 * @param cssHeight        the page's intended CSS height (px, number) — the %
 *                         denominator for the y axis.
 * @param onHyperlinkClick IX1 — invoked when a run carrying a resolved
 *                         {@link HyperlinkTarget} is clicked. A hyperlink run's
 *                         span keeps its transparent glyphs (the visible link
 *                         colour/underline is already drawn on the canvas) but
 *                         gains `cursor:pointer`, a `title` tooltip (the URL or
 *                         bookmark ref) and this click handler. A plain
 *                         `<span>` — not an `<a href>` — is used deliberately so
 *                         the browser's own navigation can never bypass the
 *                         caller's URL sanitisation. When omitted, link runs are
 *                         rendered exactly like plain runs (no click affordance).
 * @param measureForFont   optional width-measurer factory (primed with a run's
 *                         `font`), used ONLY to clamp a §17.3.2.10 縦中横
 *                         (eastAsianVert) span to its drawn one-em cell (#836):
 *                         the span composes a `scaleX(run.w / naturalWidth)` so
 *                         its selection extent matches the compressed glyphs
 *                         instead of the run's natural ~2× width. When omitted,
 *                         a 縦中横 span keeps the bare rotate (no regression for
 *                         callers that do not thread a measurer).
 */
export function buildDocxTextLayer(
  layer: HTMLDivElement,
  runs: DocxTextRunInfo[],
  cssWidth: number,
  cssHeight: number,
  onHyperlinkClick?: (target: HyperlinkTarget) => void,
  measureForFont?: (font: string) => (s: string) => number,
  pageIndex?: number,
): void {
  layer.innerHTML = '';
  setSelectionData(layer, 'ooxmlSelectionSurface', 'docx');
  setSelectionData(layer, 'pageIndex', pageIndex === undefined ? undefined : String(pageIndex));

  const ownerDocument = layer.ownerDocument ?? document;
  const clipFrames = new OverlayClipFrames(layer, cssWidth, cssHeight);
  // Span metrics (font size, line height, tracking) are written in container
  // query units of the layer, so a span keeps its size relative to the page
  // when the page is resized uniformly, as its position and any clip
  // container do. At the intended size they equal the px values.
  layer.style.containerType = 'inline-size';
  // The layer's inline axis is the page width whatever the host's writing
  // mode, so cqw measure the page.
  layer.style.writingMode = 'horizontal-tb';
  const pageLength = (px: number): string => (cssWidth > 0 ? `${(px / cssWidth) * 100}cqw` : `${px}px`);
  const pageFont = (font: string): string => font.replace(
    /(\d+(?:\.\d+)?)px/u, (_match, size: string) => pageLength(Number(size)),
  );
  for (const [runIndex, run] of runs.entries()) {
    // WORD_CELL_CONTENT_CLIP: text its clips hide is not selectable. A run
    // hidden wholly becomes one unselectable span that
    // readDocxTextSelectionContext skips; a run straddling a clip edge is
    // split at the edge, by its painted clusters, into visible and hidden
    // spans. Other runs keep one span.
    // Each piece is fitted to its painted clusters (scaledPieces).
    const pieces = scaledPieces(run, runPieces(run), measureForFont);
    // The `font` shorthand must precede `line-height` because the shorthand
    // resets `line-height` to `normal`. Reset `letter-spacing` so a parent
    // CSS rule cannot drift the trailing edge of the selection. Kerning /
    // ligatures are left at the browser default ('auto') because canvas
    // `measureText` / `fillText` also apply them by default — forcing them
    // off here would make the span wider than the drawn text.
    // ECMA-376 §17.6.20 (tbRl) — a vertical page reports `x`/`y` as the PHYSICAL
    // top-left and carries a `transform` (`rotate(90deg)`); rotate the span about
    // its top-left so the horizontal DOM text lies along the drawn (rotated) glyph
    // run. Horizontal pages carry no orientation transform.
    // ECMA-376 §17.3.2.10 縦中横 (#836): a tate-chu-yoko run is drawn compressed
    // into one em cell (`run.w`), so append a `scaleX(run.w / naturalWidth)` (when
    // a measurer is available) so the span's selectable extent matches the drawn
    // cell instead of the natural ~2× glyph width. Ordered AFTER the rotate so it
    // compresses the span's horizontal (pre-rotation) axis — the along-column
    // direction of the cell. A no-op factor of 1 for every ordinary run (see
    // tate-chu-yoko-overlay.ts).
    // A run projected with `referenceFont` is laid out by the browser at the
    // reference em and scaled onto the run (`scale(k)`, about the same top-left
    // origin, after the orientation): its transparent glyphs then advance
    // exactly like the drawn, size-proportional ones even where the host
    // grid-fits the display size (core reference-advances.ts). Lengths inside
    // the span (line height, tracking, the clipped hit box) are therefore in
    // reference px, i.e. divided by k. A run without it keeps `font` as is.
    const reference = run.referenceFont;
    const k = reference && reference.scale > 0 && Number.isFinite(reference.scale) ? reference.scale : 1;
    // WORD_CELL_CONTENT_CLIP: a clipped run is placed in nested containers
    // of its clips (overflow hidden), which clip its hit box and selection.
    const frame = clipFrames.place(run);
    let transformValue = frame.orientation;
    if (k !== 1) transformValue = `${transformValue ? `${transformValue} ` : ''}scale(${k})`;
    const tcy = measureForFont && run.eastAsianVert
      ? tateChuYokoOverlayScale(run, runOverlayMeasure(run, measureForFont)) : 1;
    if (tcy !== 1) transformValue = `${transformValue ? `${transformValue} ` : ''}scaleX(${tcy})`;
    // IX1 — a run with a resolved hyperlink target becomes a clickable region.
    // Only the cursor changes to a pointer and a title tooltip is added; the
    // glyphs stay `color:transparent` (the link's blue/underline is already
    // painted on the canvas), so a link run's selection behaviour is otherwise
    // identical to a plain run. A non-link run is byte-identical to before.
    const link = onHyperlinkClick ? run.hyperlink : undefined;
    const letterSpacing = run.letterSpacingPx !== undefined
      ? pageLength(run.letterSpacingPx / k)
      : '0';
    // A Word-fitted terminal space shortens the retained run, but the browser
    // still lays out the transparent U+0020 at its natural advance. Clip only
    // that run's hit box so it cannot cover the following run/link. Keep the
    // width percentage-based like its position under responsive canvas scaling.
    // Position the span as a % of the page's intended CSS box, and size its
    // glyph metrics in units of the layer width, so it tracks the canvas's
    // actual rendered size under external CSS scaling. A
    // vertical (tbRl) run's rotate is about `top left`, so the %-placed origin
    // rotates into the column exactly as before at 1× scale.
    for (const piece of pieces) {
      const span = ownerDocument.createElement('span');
      setSelectionData(span, 'ooxmlSelectionRun', 'docx');
      setSelectionData(span, 'runIndex', String(runIndex));
      if (piece.hidden) setSelectionData(span, 'clippedAway', '');
      if (run.paragraphId !== undefined) setSelectionData(span, 'paragraphId', run.paragraphId);
      if (run.source !== undefined) {
        setSelectionData(span, 'sourceStory', run.source.story);
        setSelectionData(span, 'sourceStoryInstance', run.source.storyInstance);
        setSelectionData(span, 'sourcePath', JSON.stringify(run.source.path));
      }
      span.textContent = run.text.slice(piece.start, piece.end);
      const pieceLink = piece.hidden ? undefined : link;
      // A span lies where its clusters are painted (scaled or fitted glyphs
      // included): scaled along the run so its glyph advances, and with them
      // native selection and the hit box, match the paint. Without a measurer
      // the box is sized to the painted width instead.
      const paintScale = piece.scale;
      const pieceTransformValue = paintScale !== undefined && Math.abs(paintScale - 1) > 1e-3
        ? `${transformValue ? `${transformValue} ` : ''}scaleX(${paintScale})` : transformValue;
      const pieceTransform = pieceTransformValue
        ? `transform:${pieceTransformValue};transform-origin:top left;` : '';
      span.style.cssText =
        `position:absolute;` +
        frame.position(piece.offset, 0) +
        `font:${pageFont(k !== 1 && reference ? reference.font : run.font)};line-height:${pageLength(run.h / k)};letter-spacing:${piece.untracked ? '0' : letterSpacing};` +
        (piece.width !== undefined && paintScale === undefined
          ? `width:${frame.width(piece.width / (k * tcy))};overflow:hidden;overflow:clip;`
          : piece.width === undefined && piece.end === run.text.length && run.trailingSpaceCompressionPx !== undefined
            ? `width:${frame.width((run.w - piece.offset) / k)};overflow:hidden;overflow:clip;` : '') +
        pieceTransform +
        `white-space:pre;color:transparent;cursor:${pieceLink ? 'pointer' : 'text'};pointer-events:all;` +
        (piece.hidden ? 'user-select:none;-webkit-user-select:none;' : '');
      if (pieceLink && onHyperlinkClick) {
        const link = pieceLink;
        span.title = link.kind === 'external' ? link.url : link.ref;
        span.addEventListener('click', (event) => {
          // Hyperlink activation owns this click. The Viewer-level element-context
          // listener observes defaultPrevented and must not focus a drawing under
          // link text in the same gesture.
          event.preventDefault();
          onHyperlinkClick(link);
        });
      }
      frame.parent.appendChild(span);
    }
  }
}

interface RunPiece {
  readonly start: number;
  readonly end: number;
  /** Offset of the piece along the run's axis, in CSS px. */
  readonly offset: number;
  /** Painted length of the piece along the run's axis, in CSS px, when the
   * run carries cluster boxes. */
  readonly width?: number;
  readonly hidden: boolean;
  /** The piece's cluster boxes, `[first, last]`, when the run has them. */
  readonly first?: number;
  readonly last?: number;
  /** Horizontal scale fitting the piece's natural text to its painted width. */
  readonly scale?: number;
  /** The piece is laid out without the run's tracking (see scaledPieces). */
  readonly untracked?: true;
}

/** WORD_CELL_CONTENT_CLIP: the visible and hidden pieces of `run`'s text under
 * its clips, in source order. Each of the run's painted clusters
 * (DocxTextRunInfo.clusterBoxes, the layout's own geometry, so right-to-left
 * order, character scaling and fitted text are as painted) is tested against
 * the clips, union gaps included, and consecutive clusters with the same
 * visibility form a piece; no text is measured, and the work is linear. A run
 * that is wholly visible or wholly hidden, or has no cluster boxes, is one
 * piece. */
function runPieces(run: Readonly<DocxTextRunInfo>): readonly RunPiece[] {
  const boxes = run.clusterBoxes;
  // A run with cluster boxes spans its painted extent whole as well (scaled
  // or fitted glyphs), like a split piece.
  let paintedEnd = -Infinity;
  for (const box of boxes ?? []) paintedEnd = Math.max(paintedEnd, box.x + box.width);
  const whole = (hidden: boolean): readonly RunPiece[] => [{
    start: 0, end: run.text.length, offset: 0, hidden,
    ...(paintedEnd > 0 ? { width: paintedEnd, first: 0, last: boxes!.length - 1 } : {}),
  }];
  if (!run.clips) return whole(false);
  if (!clippedBoxVisible(
    { x: run.x, y: run.y, width: run.w, height: run.h, ...(run.transform ? { transform: run.transform } : {}) },
    run.clips,
  )) return whole(true);
  if (!boxes || boxes.length === 0) return whole(false);
  // Without a union, a run its clips do not cut has nothing to split; a union
  // can hide a gap inside the run's extent.
  if (!run.clips.some((clip) => clip.rects) && !runStraddlesClip(run)) return whole(false);
  const groups: Array<{ first: number; last: number; hidden: boolean }> = [];
  for (const [index, box] of boxes.entries()) {
    const previous = groups.at(-1);
    // A cluster without width goes with the one before it.
    const hidden = box.width > 0 ? !runPartVisible(run, box.x, box.width) : previous?.hidden ?? false;
    if (previous && previous.hidden === hidden) previous.last = index;
    else groups.push({ first: index, last: index, hidden });
  }
  if (groups.length <= 1) return whole(groups[0]?.hidden ?? false);
  return groups.map((group, index) => {
    let offset = Infinity;
    let right = -Infinity;
    for (let member = group.first; member <= group.last; member++) {
      offset = Math.min(offset, boxes[member]!.x);
      right = Math.max(right, boxes[member]!.x + boxes[member]!.width);
    }
    return {
      width: Math.max(0, right - offset),
      first: group.first,
      last: group.last,
      start: index === 0 ? 0 : boxes[group.first]!.start,
      end: index === groups.length - 1 ? run.text.length : boxes[groups[index + 1]!.first]!.start,
      offset,
      hidden: group.hidden,
    };
  });
}

/** Relative difference of cluster scales within one span. */
const UNIFORM_SCALE_TOLERANCE = 0.02;

/** Scale of a span whose clusters paint no width: it keeps its text in the
 * DOM but takes (almost) no room. */
const DEGENERATE_SCALE = 1e-3;

/** Fit each piece's spans to where its clusters are painted. A span is
 * scaled along the run so its glyph advances, and with them native
 * selection, the caret and the hit box, match the paint. One scale fits a
 * piece whose clusters are painted at a uniform scale of their natural
 * advances (the common, unscaled case is measured once and left as is).
 * Where they are not (a compressed trailing space, tracking with character
 * scaling, justification), the piece is divided into contiguous groups of
 * uniformly scaled clusters, down to one per cluster, each its own span at
 * its painted position with its own scale. Each cluster is measured once,
 * so the work is linear. Without a measurer, or for tate-chu-yoko text (its
 * own scale), pieces are unchanged. */
function scaledPieces(
  run: Readonly<DocxTextRunInfo>,
  pieces: readonly RunPiece[],
  measureForFont: ((font: string) => (s: string) => number) | undefined,
): readonly RunPiece[] {
  const boxes = run.clusterBoxes;
  if (!measureForFont || run.eastAsianVert || !boxes) return pieces;
  const measure = runOverlayMeasure(run as DocxTextRunInfo, measureForFont);
  const pitch = run.letterSpacingPx ?? 0;
  const natural = (text: string) => measure(text) + countSpacingCharacters(text).pitchBearing * pitch;
  return pieces.flatMap((piece): RunPiece[] => {
    if (piece.width === undefined || piece.first === undefined || piece.last === undefined) return [piece];
    if (!(piece.width > 0)) return [piece];
    const whole = natural(run.text.slice(piece.start, piece.end));
    if (whole > 0 && Math.abs(piece.width / whole - 1) <= 1e-3) return [{ ...piece, scale: 1 }];
    // Per-cluster scales, grouped while they stay uniform. A cluster whose
    // tracked advance is not positive (condensed past its glyph: CSS would
    // lay its text backwards) or that paints no width is fitted untracked,
    // in groups of its own kind.
    const groups: Array<{ first: number; last: number; natural: number; scale: number; untracked: boolean }> = [];
    for (let index = piece.first; index <= piece.last; index++) {
      const box = boxes[index]!;
      const text = run.text.slice(box.start, box.end);
      const tracked = natural(text);
      const untracked = !(tracked > 0 && box.width > 0);
      const own = untracked ? measure(text) : tracked;
      const scale = box.width > 0 && own > 0 ? box.width / own : DEGENERATE_SCALE;
      const previous = groups.at(-1);
      if (previous && previous.untracked === untracked
        && Math.abs(scale / previous.scale - 1) <= UNIFORM_SCALE_TOLERANCE) {
        previous.last = index;
        previous.natural += own;
      } else {
        groups.push({ first: index, last: index, natural: own, scale, untracked });
      }
    }
    return groups.map((group, index) => {
      let offset = Infinity;
      let right = -Infinity;
      for (let member = group.first; member <= group.last; member++) {
        offset = Math.min(offset, boxes[member]!.x);
        right = Math.max(right, boxes[member]!.x + boxes[member]!.width);
      }
      const width = Math.max(0, right - offset);
      return {
        start: index === 0 ? piece.start : boxes[group.first]!.start,
        end: index === groups.length - 1 ? piece.end : boxes[groups[index + 1]!.first]!.start,
        offset,
        width,
        hidden: piece.hidden,
        first: group.first,
        last: group.last,
        scale: group.natural > 0 && width > 0 ? width / group.natural : DEGENERATE_SCALE,
        ...(group.untracked ? { untracked: true as const } : {}),
      };
    });
  });
}
