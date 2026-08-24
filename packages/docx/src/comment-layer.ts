/**
 * DOM builder for the ECMA-376 §17.13.4 comment margin.
 *
 * Follows the find-highlight overlay's architecture (find-highlight-layer.ts):
 * positioned DOM layers over/beside the canvas, geometry from the SAME
 * per-run projection the page was rendered from (`DocxTextRunInfo`), no
 * canvas draw pass. Two layers per page:
 *
 * - the TINT layer sits over the page (`inset:0`, percent-positioned children
 *   so external CSS scaling keeps them on the glyphs) and marks the commented
 *   ranges plus a connector stub from each anchor toward the margin;
 * - the GUTTER layer sits immediately RIGHT of the page (`left:100%`, a fixed
 *   pixel width) and holds the balloons placed by the pure
 *   {@link computeCommentBalloonLayout} model, with a click-to-select
 *   affordance (`pointer-events` stays `none` on the layer, `auto` on each
 *   balloon, so page interactions under the gutter are unaffected).
 *
 * Selecting a thread — by clicking its balloon OR its tinted range on the
 * page — expands the balloon (the pure model raises the selected cap to the
 * page's line capacity) and makes its body scrollable for anything still cut
 * off. The tinted-range click is a HIT-TEST on the page wrapper (one shared
 * listener; the tint boxes themselves keep `pointer-events:none`), so text
 * selection and other page interactions over commented ranges are unaffected.
 *
 * All geometry inputs come from the pure comment-margin model; this module
 * only materializes DOM. Range → run joining uses
 * (`source.path`, `sourceRunIndex`), which both render modes ship.
 */

import { overlayPercent } from '@silurus/ooxml-core';
import type { DocxTextRunInfo } from './renderer';
import {
  computeCommentBalloonLayout,
  type CommentAnchorRange,
  type CommentBalloonRequest,
  type CommentThread,
} from './comment-margin-layout.js';
import type { DocComment } from './types.js';

/** Soft range tint (translucent so drawn glyphs stay legible), and the
 *  stronger emphasis used for the selected thread's range. */
export const DEFAULT_COMMENT_TINT = 'rgba(255, 196, 66, 0.30)';
export const DEFAULT_COMMENT_ACTIVE_TINT = 'rgba(255, 160, 30, 0.45)';
const CONNECTOR_COLOR = '#c19a3f';
const BALLOON_FONT_PX = 11;
const BALLOON_LINE_PX = 14;
const BALLOON_HEADER_PX = 18;
const BALLOON_GAP_PX = 8;
const BALLOON_SIDE_INSET_PX = 8;
const BALLOON_PAD_PX = 6;

export interface DocxCommentLayerModel {
  readonly threads: readonly CommentThread[];
  readonly ranges: readonly CommentAnchorRange[];
}

export interface DocxCommentLayerGeometry {
  /** The page's intended CSS width/height in px — the tint layer's percent
   *  denominators (same contract as the find-highlight layer). */
  readonly cssWidth: number;
  readonly cssHeight: number;
  /** The gutter layer's fixed CSS-px width. */
  readonly gutterWidthPx: number;
}

interface PageAnchor {
  readonly thread: CommentThread;
  readonly run: DocxTextRunInfo;
}

function runKey(run: DocxTextRunInfo): string | null {
  if (!run.source || run.sourceRunIndex === undefined) return null;
  return `${run.source.story}:${run.source.storyInstance}:${run.source.path.join('.')}`;
}

function rangeKey(range: CommentAnchorRange): string {
  return `body:body:${range.paragraphPath.join('.')}`;
}

/** Runs covered by a range on this page, in document order. A zero-length
 *  (reference-only) range snaps to the run at its boundary, else the one
 *  before it. */
function coveredRuns(
  range: CommentAnchorRange,
  runsByParagraph: ReadonlyMap<string, readonly DocxTextRunInfo[]>,
): DocxTextRunInfo[] {
  const paragraphRuns = runsByParagraph.get(rangeKey(range)) ?? [];
  if (range.startRunIndex === range.endRunIndex) {
    const at = paragraphRuns.filter((run) => run.sourceRunIndex === range.startRunIndex);
    if (at.length > 0) return at;
    return paragraphRuns.filter((run) => run.sourceRunIndex === range.startRunIndex - 1);
  }
  return paragraphRuns.filter((run) =>
    run.sourceRunIndex !== undefined
    && run.sourceRunIndex >= range.startRunIndex
    && run.sourceRunIndex < range.endRunIndex);
}

/** Estimate the content lines a comment body needs at the balloon's fixed
 *  typography. Deterministic (no DOM measurement): character budget per line
 *  from the gutter's writable width. */
function commentLines(comment: DocComment, charsPerLine: number): number {
  const paragraphs = comment.paragraphs?.length ? comment.paragraphs : [comment.text];
  return paragraphs.reduce(
    (sum, text) => sum + Math.max(1, Math.ceil(text.length / charsPerLine)),
    0,
  );
}

function threadLines(thread: CommentThread, charsPerLine: number): number {
  // Each reply carries its own one-line author header plus its body.
  return commentLines(thread.root, charsPerLine)
    + thread.replies.reduce(
      (sum, reply) => sum + 1 + commentLines(reply, charsPerLine),
      0,
    );
}

function clearLayer(layer: HTMLElement): void {
  layer.innerHTML = '';
}

// ── Tinted-range click-to-select ────────────────────────────────────────────

/** One clickable commented-range box, in the page's CSS-px space. */
interface TintHitRegion {
  readonly commentId: string;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

interface TintHitState {
  readonly regions: readonly TintHitRegion[];
  readonly cssWidth: number;
  readonly cssHeight: number;
  readonly selectedCommentId: string | null;
  readonly onSelect: (commentId: string | null) => void;
  /** The live tint layer — its bounding box is the hit-test's denominator, so
   *  external CSS scaling of the page keeps clicks on the glyphs. A detached
   *  layer (comments toggled off / slot recycled) reports a zero box, which
   *  disables the stale state without unhooking the wrapper listener. */
  readonly tintLayer: HTMLDivElement;
}

const TINT_HIT_STATE = '__docxCommentTintHits';

type TintHitHost = HTMLElement & { [TINT_HIT_STATE]?: TintHitState };

/** Install (once per wrapper) a click hit-test that selects the thread whose
 *  tinted range was clicked. Rebuilds only swap the state object, so the
 *  wrapper never accumulates listeners; the tint boxes stay
 *  `pointer-events:none` so text selection over commented text still works. */
function installTintHitTest(
  tintLayer: HTMLDivElement,
  regions: readonly TintHitRegion[],
  cssWidth: number,
  cssHeight: number,
  selectedCommentId: string | null,
  onSelect: (commentId: string | null) => void,
): void {
  const host = tintLayer.parentElement as TintHitHost | null;
  if (!host) return;
  const installed = host[TINT_HIT_STATE] !== undefined;
  host[TINT_HIT_STATE] = { regions, cssWidth, cssHeight, selectedCommentId, onSelect, tintLayer };
  if (installed) return;
  host.addEventListener('click', (event: MouseEvent) => {
    const state = host[TINT_HIT_STATE];
    if (!state || state.regions.length === 0) return;
    if (state.cssWidth <= 0 || state.cssHeight <= 0) return;
    const box = state.tintLayer.getBoundingClientRect();
    if (!(box.width > 0) || !(box.height > 0)) return;
    const x = ((event.clientX - box.left) / box.width) * state.cssWidth;
    const y = ((event.clientY - box.top) / box.height) * state.cssHeight;
    const hit = state.regions.find(
      (region) => x >= region.x && x <= region.x + region.w
        && y >= region.y && y <= region.y + region.h,
    );
    if (!hit) return;
    state.onSelect(state.selectedCommentId === hit.commentId ? null : hit.commentId);
  });
}

function div(host: HTMLElement, cssText: string): HTMLDivElement {
  const el = host.ownerDocument
    ? host.ownerDocument.createElement('div')
    : document.createElement('div');
  el.style.cssText = cssText;
  host.appendChild(el);
  return el;
}

/**
 * Populate the comment tint + gutter layers for one page.
 *
 * @param tintLayer   overlay div covering the page (cleared here).
 * @param gutterLayer overlay div beside the page (cleared here).
 * @param runs        the page's projected runs (the render's own geometry).
 * @param model       threads + anchor ranges from the pure comment model.
 * @param geometry    page CSS box + gutter width.
 * @param selectedCommentId  the selected thread's root id, or null.
 * @param onSelect    click handler: balloon click selects (toggles) a thread.
 */
export function buildDocxCommentLayer(
  tintLayer: HTMLDivElement,
  gutterLayer: HTMLDivElement,
  runs: readonly DocxTextRunInfo[],
  model: DocxCommentLayerModel,
  geometry: DocxCommentLayerGeometry,
  selectedCommentId: string | null,
  onSelect: (commentId: string | null) => void,
): void {
  clearLayer(tintLayer);
  clearLayer(gutterLayer);
  const { cssWidth, cssHeight, gutterWidthPx } = geometry;
  // The hit-test state is (re)installed on EVERY build — including the empty
  // ones below — so a rebuild for a new page/document never leaves a previous
  // build's clickable regions behind.
  const hitRegions: TintHitRegion[] = [];
  const commitHitRegions = (): void =>
    installTintHitTest(tintLayer, hitRegions, cssWidth, cssHeight, selectedCommentId, onSelect);
  if (cssWidth <= 0 || cssHeight <= 0 || model.threads.length === 0) {
    commitHitRegions();
    return;
  }

  const runsByParagraph = new Map<string, DocxTextRunInfo[]>();
  for (const run of runs) {
    const key = runKey(run);
    if (key === null) continue;
    const bucket = runsByParagraph.get(key);
    if (bucket) bucket.push(run);
    else runsByParagraph.set(key, [run]);
  }

  const threadIds = new Map(model.threads.map((thread) => [thread.root.id, thread]));
  const anchors = new Map<string, PageAnchor>();

  // Tint every covered run and resolve each thread's anchor (its topmost /
  // leftmost covered run on this page).
  for (const range of model.ranges) {
    const thread = threadIds.get(range.commentId);
    if (!thread) continue; // resolved thread or reply-range: no balloon, no tint
    const covered = coveredRuns(range, runsByParagraph);
    for (const run of covered) {
      const active = selectedCommentId === range.commentId;
      const box = div(
        tintLayer,
        'position:absolute;'
        + `left:${overlayPercent(run.x, cssWidth)};`
        + `top:${overlayPercent(run.y, cssHeight)};`
        + `width:${overlayPercent(run.w, cssWidth)};`
        + `height:${overlayPercent(run.h, cssHeight)};`
        + `background:${active ? DEFAULT_COMMENT_ACTIVE_TINT : DEFAULT_COMMENT_TINT};`
        + 'pointer-events:none;',
      );
      if (run.transform) {
        box.style.transform = run.transform;
        box.style.transformOrigin = 'top left';
      }
      hitRegions.push({ commentId: range.commentId, x: run.x, y: run.y, w: run.w, h: run.h });
      const current = anchors.get(range.commentId);
      if (!current
        || run.y < current.run.y
        || (run.y === current.run.y && run.x < current.run.x)) {
        anchors.set(range.commentId, { thread, run });
      }
    }
  }
  commitHitRegions();
  if (anchors.size === 0) return;

  const charsPerLine = Math.max(
    8,
    Math.floor(
      (gutterWidthPx - 2 * BALLOON_SIDE_INSET_PX - 2 * BALLOON_PAD_PX)
      / (BALLOON_FONT_PX * 0.55),
    ),
  );
  const requests: CommentBalloonRequest[] = [...anchors.values()].map(({ thread, run }) => ({
    commentId: thread.root.id,
    anchorYPx: run.y,
    contentLines: threadLines(thread, charsPerLine),
    selected: selectedCommentId === thread.root.id,
  }));
  const placements = computeCommentBalloonLayout({
    balloons: requests,
    pageHeightPx: cssHeight,
    lineHeightPx: BALLOON_LINE_PX,
    headerHeightPx: BALLOON_HEADER_PX,
    gapPx: BALLOON_GAP_PX,
  });

  for (const placement of placements) {
    const anchor = anchors.get(placement.commentId);
    if (!anchor) continue;
    const thread = anchor.thread;
    const selected = placement.selected;

    // Connector stub: anchor's right edge → the page's right edge at the
    // anchor line (percent-positioned so it tracks CSS scaling), then a leg
    // inside the gutter to the balloon's left edge.
    const anchorLineY = anchor.run.y + anchor.run.h / 2;
    div(
      tintLayer,
      'position:absolute;'
      + `left:${overlayPercent(anchor.run.x + anchor.run.w, cssWidth)};`
      + `top:${overlayPercent(anchorLineY, cssHeight)};`
      + `width:${overlayPercent(Math.max(0, cssWidth - anchor.run.x - anchor.run.w), cssWidth)};`
      + `height:1px;background:${CONNECTOR_COLOR};opacity:${selected ? '0.9' : '0.5'};`
      + 'pointer-events:none;',
    );
    const legX = Math.round(BALLOON_SIDE_INSET_PX / 2);
    const legTop = Math.min(anchorLineY, placement.yPx + BALLOON_HEADER_PX / 2);
    const legBottom = Math.max(anchorLineY, placement.yPx + BALLOON_HEADER_PX / 2);
    div(
      gutterLayer,
      'position:absolute;'
      + `left:0;top:${anchorLineY}px;width:${legX}px;height:1px;`
      + `background:${CONNECTOR_COLOR};opacity:${selected ? '0.9' : '0.5'};pointer-events:none;`,
    );
    if (legBottom - legTop >= 1) {
      div(
        gutterLayer,
        'position:absolute;'
        + `left:${legX}px;top:${legTop}px;width:1px;height:${legBottom - legTop}px;`
        + `background:${CONNECTOR_COLOR};opacity:${selected ? '0.9' : '0.5'};pointer-events:none;`,
      );
    }

    const balloon = div(
      gutterLayer,
      'position:absolute;'
      + `left:${BALLOON_SIDE_INSET_PX}px;`
      + `top:${placement.yPx}px;`
      + `width:${Math.max(0, gutterWidthPx - 2 * BALLOON_SIDE_INSET_PX)}px;`
      + `height:${placement.heightPx}px;`
      // The SELECTED balloon scrolls: its cap already expanded to a page-full
      // (comment-margin-layout), and anything past the placed height — a very
      // long thread, or the estimate under-counting wrapped lines — stays
      // reachable by scrolling instead of being cut off.
      + `box-sizing:border-box;${selected ? 'overflow-y:auto;overflow-x:hidden;' : 'overflow:hidden;'}`
      + `padding:2px ${BALLOON_PAD_PX}px;`
      + 'background:#fffdf5;'
      + `border:1px solid ${selected ? '#b8860b' : '#d8c48a'};`
      + `border-left:3px solid ${selected ? '#b8860b' : '#d8c48a'};`
      + 'border-radius:3px;'
      + `font:${BALLOON_FONT_PX}px/${BALLOON_LINE_PX}px sans-serif;color:#333;`
      + 'cursor:pointer;pointer-events:auto;user-select:none;-webkit-user-select:none;',
    );
    balloon.addEventListener('click', (event) => {
      (event as Event).preventDefault?.();
      onSelect(selected ? null : thread.root.id);
    });

    const header = div(
      balloon,
      `height:${BALLOON_HEADER_PX - 4}px;overflow:hidden;`
      + 'font-weight:bold;white-space:nowrap;text-overflow:ellipsis;',
    );
    header.textContent = [thread.root.author ?? '', thread.root.date ?? '']
      .filter((part) => part.length > 0)
      .join(' · ');
    if (placement.collapsed) continue;
    const appendBody = (comment: DocComment, indentPx: number) => {
      const paragraphs = comment.paragraphs?.length ? comment.paragraphs : [comment.text];
      for (const text of paragraphs) {
        const p = div(balloon, `margin-left:${indentPx}px;`);
        p.textContent = text;
      }
    };
    appendBody(thread.root, 0);
    for (const reply of thread.replies) {
      const replyHeader = div(
        balloon,
        'margin-left:10px;font-weight:bold;white-space:nowrap;'
        + 'overflow:hidden;text-overflow:ellipsis;',
      );
      replyHeader.textContent = [reply.author ?? '', reply.date ?? '']
        .filter((part) => part.length > 0)
        .join(' · ');
      appendBody(reply, 10);
    }
  }
}
