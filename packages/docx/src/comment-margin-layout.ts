/**
 * Pure geometry + threading model for the ECMA-376 §17.13.4 comment margin.
 *
 * Three DOM-free pieces, unit-tested in isolation (the house pattern of
 * xlsx's `comment-popup.ts`: pure module computes, the viewer owns the DOM):
 *
 * 1. {@link buildCommentThreads} — group the flat `doc.comments` list into
 *    top-level threads with ordered replies (`w15:commentEx` paraIdParent
 *    join, already resolved to `parentId` by the parser). Threads whose root
 *    is resolved are dropped entirely (read-only view hides resolved).
 * 2. {@link collectDocumentCommentRanges} — resolve the paragraph-level
 *    anchor marks (`DocParagraph.commentMarks`) into per-paragraph run
 *    intervals, carrying ranges that span paragraphs. Emitted run indexes
 *    are normalized-model indexes (parser-only sidecar runs skipped), so
 *    they join `DocxTextRunInfo.sourceRunIndex` + `source.path` directly.
 * 3. {@link computeCommentBalloonLayout} — place balloons in the right
 *    margin gutter: anchor-ordered, non-overlapping (monotonic push-down),
 *    capped at `maxLines`, truncated back toward each next balloon's anchor,
 *    collapsed to header stubs under page-height pressure, and always giving
 *    the selected balloon its full capped height — where selection RAISES the
 *    cap to the page's own line capacity, so a selected thread expands (the
 *    DOM layer makes anything still cut off scrollable).
 */

import type { BodyElement, DocComment, DocParagraph, DocxCommentMark } from './types.js';

// ── 1. Threading ────────────────────────────────────────────────────────────

export interface CommentThread {
  /** The top-level comment. */
  readonly root: DocComment;
  /** Replies in document (comments-part) order, parent chains flattened. */
  readonly replies: readonly DocComment[];
}

/** Group comments into unresolved top-level threads. A reply whose parent
 * chain cannot be resolved to a known root is dropped (malformed part), and a
 * thread whose ROOT is resolved is dropped with all of its replies. */
export function buildCommentThreads(comments: readonly DocComment[]): CommentThread[] {
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  const rootOf = (comment: DocComment): DocComment | undefined => {
    let current = comment;
    const seen = new Set<string>([comment.id]);
    while (current.parentId !== undefined) {
      const parent = byId.get(current.parentId);
      if (!parent || seen.has(parent.id)) return undefined;
      seen.add(parent.id);
      current = parent;
    }
    return current;
  };
  const threads = new Map<string, { root: DocComment; replies: DocComment[] }>();
  for (const comment of comments) {
    if (comment.parentId === undefined) {
      threads.set(comment.id, { root: comment, replies: [] });
    }
  }
  for (const comment of comments) {
    if (comment.parentId === undefined) continue;
    const root = rootOf(comment);
    if (!root) continue;
    threads.get(root.id)?.replies.push(comment);
  }
  return [...threads.values()]
    .filter((thread) => thread.root.resolved !== true)
    .map((thread) => Object.freeze({
      root: thread.root,
      replies: Object.freeze(thread.replies),
    }));
}

// ── 2. Anchor ranges ────────────────────────────────────────────────────────

export interface CommentAnchorRange {
  readonly commentId: string;
  /** Structural path of the owning paragraph (same scheme as
   *  `DocxTextRunInfo.source.path` for the body story). */
  readonly paragraphPath: readonly number[];
  /** First covered run (normalized-model index, inclusive). */
  readonly startRunIndex: number;
  /** One past the last covered run. `start === end` marks a reference-only
   *  boundary (no `commentRangeStart/End` pair authored). */
  readonly endRunIndex: number;
}

interface ParagraphVisit {
  readonly paragraph: DocParagraph;
  readonly path: readonly number[];
}

function* paragraphsInDocumentOrder(
  elements: readonly BodyElement[],
  prefix: readonly number[] = [],
): Generator<ParagraphVisit> {
  for (const [index, element] of elements.entries()) {
    if (element.type === 'paragraph') {
      yield { paragraph: element as DocParagraph, path: [...prefix, index] };
    } else if (element.type === 'table') {
      const table = element as BodyElement & {
        rows?: readonly { cells: readonly { content: readonly BodyElement[] }[] }[];
      };
      for (const [rowIndex, row] of (table.rows ?? []).entries()) {
        for (const [cellIndex, cell] of row.cells.entries()) {
          yield* paragraphsInDocumentOrder(cell.content, [...prefix, index, rowIndex, cellIndex]);
        }
      }
    }
  }
}

/** Map each parser run boundary to its normalized-model boundary: the layout
 * pipeline strips parser-only sidecar runs (`unavailableDrawing`), so a mark
 * boundary is preceded by only the runs the projection kept. */
function normalizedBoundary(paragraph: DocParagraph, rawRunIndex: number): number {
  const runs = paragraph.runs as readonly Readonly<{ type?: string }>[];
  let publicIndex = 0;
  for (let index = 0; index < Math.min(rawRunIndex, runs.length); index += 1) {
    if (runs[index]?.type !== 'unavailableDrawing') publicIndex += 1;
  }
  return publicIndex;
}

function normalizedRunCount(paragraph: DocParagraph): number {
  return normalizedBoundary(paragraph, paragraph.runs.length);
}

/** Resolve comment anchors to per-paragraph run intervals in document order.
 * Ranges spanning paragraphs emit one interval per covered paragraph
 * (including mark-less middle paragraphs); an unmatched `rangeStart` stays
 * open to the end of the body; a comment with only a `commentReference`
 * yields a zero-length boundary at the reference run. */
export function collectDocumentCommentRanges(
  body: readonly BodyElement[],
): CommentAnchorRange[] {
  const ranges: CommentAnchorRange[] = [];
  /** commentId → start boundary within the CURRENT paragraph (open ranges
   * that started in an earlier paragraph enter at 0). */
  const open = new Map<string, number>();
  const rangedIds = new Set<string>();
  const referenceOnly = new Map<string, { path: readonly number[]; boundary: number }>();
  for (const { paragraph, path } of paragraphsInDocumentOrder(body)) {
    const marks = (paragraph.commentMarks ?? []) as readonly DocxCommentMark[];
    const runCount = normalizedRunCount(paragraph);
    const emitted = new Set<string>();
    for (const mark of marks) {
      const boundary = normalizedBoundary(paragraph, mark.runIndex);
      if (mark.kind === 'rangeStart') {
        rangedIds.add(mark.id);
        if (!open.has(mark.id)) open.set(mark.id, boundary);
      } else if (mark.kind === 'rangeEnd') {
        rangedIds.add(mark.id);
        const start = open.get(mark.id) ?? 0;
        open.delete(mark.id);
        ranges.push(Object.freeze({
          commentId: mark.id,
          paragraphPath: Object.freeze([...path]),
          startRunIndex: start,
          endRunIndex: Math.max(start, boundary),
        }));
        emitted.add(mark.id);
      } else if (mark.kind === 'reference' && !referenceOnly.has(mark.id)) {
        referenceOnly.set(mark.id, { path, boundary });
      }
    }
    // Ranges still open leaving this paragraph cover it to its end.
    for (const [id, start] of open) {
      if (emitted.has(id)) continue;
      ranges.push(Object.freeze({
        commentId: id,
        paragraphPath: Object.freeze([...path]),
        startRunIndex: start,
        endRunIndex: Math.max(start, runCount),
      }));
      open.set(id, 0);
    }
  }
  // §17.13.4.5 reference-only comments (no authored range): a zero-length
  // boundary the overlay snaps to the adjacent run.
  for (const [id, at] of referenceOnly) {
    if (rangedIds.has(id)) continue;
    ranges.push(Object.freeze({
      commentId: id,
      paragraphPath: Object.freeze([...at.path]),
      startRunIndex: at.boundary,
      endRunIndex: at.boundary,
    }));
  }
  return ranges;
}

// ── 3. Balloon layout ───────────────────────────────────────────────────────

export interface CommentBalloonRequest {
  readonly commentId: string;
  /** Anchor line's top edge, in the page's CSS-px coordinate space. */
  readonly anchorYPx: number;
  /** Total content lines the balloon body would need (root + replies, as
   *  measured by the caller). */
  readonly contentLines: number;
  /** Selection wins full height and its anchor position. */
  readonly selected?: boolean;
}

export interface CommentBalloonPlacement {
  readonly commentId: string;
  readonly yPx: number;
  readonly heightPx: number;
  /** Content lines actually visible (0 for a collapsed header-only stub). */
  readonly visibleLines: number;
  /** True when squeezed to a header-only stub. */
  readonly collapsed: boolean;
  readonly selected: boolean;
}

export interface CommentBalloonLayoutInput {
  readonly balloons: readonly CommentBalloonRequest[];
  readonly pageHeightPx: number;
  readonly lineHeightPx: number;
  readonly headerHeightPx: number;
  readonly gapPx: number;
  /** Hard cap on visible content lines per balloon (default 10). The SELECTED
   *  balloon's cap expands to the page's line capacity instead (never below
   *  this value), so selection reveals the thread up to a page-full. */
  readonly maxLines?: number;
}

const DEFAULT_MAX_LINES = 10;

/** Place margin balloons. Invariants (all pinned by tests):
 * - anchor order is kept, balloons never overlap (≥ `gapPx` apart);
 * - no balloon shows more than `maxLines` content lines;
 * - a balloon pushed below its anchor first reclaims space by truncating its
 *   predecessors toward their one-line floor ("truncate to the start of the
 *   next comment"), never by reordering;
 * - when the stack still overflows the page, trailing unselected balloons
 *   collapse to header-only stubs;
 * - the selected balloon always keeps its full capped height, and its cap
 *   expands from `maxLines` to the page's line capacity (never below
 *   `maxLines`), so selecting a long thread reveals it up to a page-full —
 *   content beyond that is the DOM layer's job (scrollable balloon body). */
export function computeCommentBalloonLayout(
  input: CommentBalloonLayoutInput,
): CommentBalloonPlacement[] {
  const maxLines = input.maxLines ?? DEFAULT_MAX_LINES;
  const { lineHeightPx, headerHeightPx, gapPx } = input;
  const selectedMaxLines = Math.max(
    maxLines,
    Math.floor((input.pageHeightPx - headerHeightPx) / lineHeightPx),
  );
  const ordered = [...input.balloons]
    .map((balloon, index) => ({ balloon, index }))
    .sort((a, b) => a.balloon.anchorYPx - b.balloon.anchorYPx || a.index - b.index)
    .map(({ balloon }) => balloon);
  const cappedLines = (balloon: CommentBalloonRequest): number =>
    Math.min(
      Math.max(balloon.contentLines, 1),
      balloon.selected === true ? selectedMaxLines : maxLines,
    );
  const state = ordered.map((balloon) => ({
    balloon,
    heightPx: headerHeightPx + cappedLines(balloon) * lineHeightPx,
    collapsed: false,
  }));
  const floorPx = (entry: (typeof state)[number]): number => (
    entry.balloon.selected
      ? headerHeightPx + cappedLines(entry.balloon) * lineHeightPx
      : headerHeightPx + Math.min(1, cappedLines(entry.balloon)) * lineHeightPx
  );

  const positions = (): number[] => {
    const out: number[] = [];
    let previousBottom = -Infinity;
    for (const entry of state) {
      const y = Math.max(entry.balloon.anchorYPx, previousBottom + gapPx, 0);
      out.push(y);
      previousBottom = y + entry.heightPx;
    }
    return out;
  };

  // Truncate predecessors so each pushed-down balloon can sit at its anchor
  // (bounded by every predecessor's floor).
  for (let index = 1; index < state.length; index += 1) {
    const ys = positions();
    let deficit = ys[index]! - Math.max(state[index]!.balloon.anchorYPx, 0);
    if (deficit <= 0) continue;
    for (let j = index - 1; j >= 0 && deficit > 0; j -= 1) {
      const entry = state[j]!;
      const reducible = Math.max(0, entry.heightPx - floorPx(entry));
      // Whole-line steps: initial heights and floors are line-quantized, so
      // taking line multiples keeps every truncated balloon line-quantized.
      const take = Math.min(reducible, Math.ceil(deficit / lineHeightPx) * lineHeightPx);
      entry.heightPx -= take;
      deficit -= take;
    }
  }

  // Page-height pressure: collapse trailing unselected balloons to stubs.
  const overflows = (): boolean => {
    const ys = positions();
    const last = state.length - 1;
    return last >= 0 && ys[last]! + state[last]!.heightPx > input.pageHeightPx;
  };
  for (let index = state.length - 1; index >= 0 && overflows(); index -= 1) {
    const entry = state[index]!;
    if (entry.balloon.selected || entry.collapsed) continue;
    entry.heightPx = headerHeightPx;
    entry.collapsed = true;
  }

  const ys = positions();
  return state.map((entry, index) => {
    const visibleLines = entry.collapsed
      ? 0
      : Math.max(0, Math.round((entry.heightPx - headerHeightPx) / lineHeightPx));
    return Object.freeze({
      commentId: entry.balloon.commentId,
      yPx: ys[index]!,
      heightPx: entry.heightPx,
      visibleLines,
      collapsed: entry.collapsed,
      selected: entry.balloon.selected === true,
    });
  });
}
