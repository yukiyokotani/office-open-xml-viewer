import type { LayoutSourceStore } from './layout/layout-source-store.js';
import { sourceKey } from './layout/source-key.js';
import { overlayBoxes } from './overlay-geometry.js';
import type { SourceRef } from './layout/types.js';
import { decimalReviewIdKey } from './review-id.js';
import type { DocComment, DocxClipRegion, DocxCommentMark, DocxStorySource, DocxTextRunInfo } from './types.js';
import type { ReviewAnchorProjectionOptions } from './review-projection.js';

export interface CommentAnchorRange {
  readonly commentId: string;
  /** Exact story/path identity used by `DocxTextRunInfo.source`. */
  readonly source: Readonly<DocxStorySource>;
  /** First covered normalized-model run, inclusive. */
  readonly startRunIndex: number;
  /** One past the final covered run. Equal bounds represent a point anchor. */
  readonly endRunIndex: number;
  /** Authored `commentReference` boundary used when the covered runs have no
   * final-state geometry (for example, a comment on deleted content). */
  readonly reference: CommentAnchorPoint;
  /** Nearest final-state text-bearing run in the same story, used only when
   * neither the covered interval nor the authored reference has text geometry.
   * Absent when that story contains no projected text run. */
  readonly geometryFallback?: CommentAnchorGeometryFallback;
  /** Set when a covered run has geometry somewhere in the laid-out document
   * (known only to anchors collected with its run index). The range is then
   * shown where those runs are: a page that holds none of them, such as an
   * earlier page of the same paragraph, shows nothing of it and does not
   * resolve it from the reference boundary. */
  readonly coveredGeometry?: true;
  /** Set on the one range of a comment with several ranges that carries the
   * comment's margin card and marker: the first, in document order, of which
   * a page shows something (assignCommentCardOwners). Absent for a comment
   * with one range, and for ranges not decided against a laid-out document;
   * the card then stands with the first range a page shows. */
  readonly cardOwner?: true;
}

export interface CommentAnchorPoint {
  readonly source: Readonly<DocxStorySource>;
  readonly runIndex: number;
  /** Preferred adjacent run at the authored boundary. Consumers may fall back
   * to the opposite side when that run is absent from final-state geometry. */
  readonly affinity: 'following' | 'preceding';
}

export interface CommentAnchorGeometryFallback {
  readonly source: Readonly<DocxStorySource>;
  readonly sourceRunIndex: number;
}

export type DocxCommentAnchorKind = 'range' | 'point' | 'fallback';

/** UI-neutral rectangle for one continuous highlighted line segment: text
 * painted without a gap on one line. */
export interface DocxCommentHighlightRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly transform?: string;
  /** The clips of the highlighted text (DocxTextRunInfo.clips). The
   *  rectangle is where the text's glyphs are painted inside them (across
   *  the line, the text's font box); text that the clips hide wholly yields
   *  no rectangle. For a clip turned against the text the rectangle bounds
   *  the visible part, and a host cuts the drawn highlight to the clips, for
   *  example with nested `overflow:hidden` containers. */
  readonly clips?: readonly Readonly<DocxClipRegion>[];
}

/** One authored anchor resolved against the text geometry supplied for a page. */
export interface ResolvedDocxCommentAnchor {
  readonly anchor: Readonly<CommentAnchorRange>;
  readonly kind: DocxCommentAnchorKind;
  readonly rects: readonly Readonly<DocxCommentHighlightRect>[];
}

/** A top-level comment and its replies, limited to anchors visible in the
 * supplied page runs. This shape carries no DOM or Viewer-owned UI state. */
export interface ResolvedDocxCommentThread {
  readonly root: Readonly<DocComment>;
  readonly replies: readonly Readonly<DocComment>[];
  readonly anchors: readonly Readonly<ResolvedDocxCommentAnchor>[];
}

export interface ResolveDocxCommentThreadsOptions {
  /** Include threads whose root is resolved. Default `true`. */
  readonly includeResolved?: boolean;
}

interface ParagraphVisit {
  readonly paragraph: CommentParagraph;
  readonly source: SourceRef;
}

interface CommentParagraph {
  readonly runs: readonly Readonly<{
    type?: string;
    text?: string;
    noteRef?: unknown;
    revision?: Readonly<{ kind?: string }>;
  }>[];
  readonly commentMarks?: readonly DocxCommentMark[];
}

interface MarkLocation {
  readonly paragraphIndex: number;
  readonly source: SourceRef;
  readonly boundary: number;
  readonly runCount: number;
}

function frozenSource(source: SourceRef): Readonly<DocxStorySource> {
  return Object.freeze({ ...source, path: Object.freeze([...source.path]) });
}

function frozenPoint(location: MarkLocation): CommentAnchorPoint {
  return Object.freeze({
    source: frozenSource(location.source),
    runIndex: location.boundary,
    affinity: location.boundary < location.runCount ? 'following' : 'preceding',
  });
}

function normalizedBoundary(paragraph: CommentParagraph, rawRunIndex: number): number {
  const runs = paragraph.runs;
  let normalized = 0;
  for (let index = 0; index < Math.min(rawRunIndex, runs.length); index += 1) {
    if (runs[index]?.type !== 'unavailableDrawing') normalized += 1;
  }
  return normalized;
}

function frozenRange(
  commentId: string,
  location: MarkLocation,
  reference: MarkLocation,
  geometryFallback: MarkLocation | undefined,
  endRunIndex = location.boundary,
  coveredGeometry = false,
): CommentAnchorRange {
  return Object.freeze({
    commentId,
    source: frozenSource(location.source),
    startRunIndex: location.boundary,
    endRunIndex,
    reference: frozenPoint(reference),
    ...(geometryFallback === undefined ? {} : {
      geometryFallback: Object.freeze({
        source: frozenSource(geometryFallback.source),
        sourceRunIndex: geometryFallback.boundary,
      }),
    }),
    ...(coveredGeometry ? { coveredGeometry: true as const } : {}),
  });
}

type RenderedRunIndex = ReadonlyMap<string, ReadonlySet<number>>;

function indexRenderedRuns(runs: readonly Readonly<DocxTextRunInfo>[]): RenderedRunIndex {
  const index = new Map<string, Set<number>>();
  for (const run of runs) {
    if (run.source === undefined || run.sourceRunIndex === undefined || run.text.length === 0) continue;
    const key = sourceKey(run.source);
    const runIndices = index.get(key) ?? new Set<number>();
    if (!index.has(key)) index.set(key, runIndices);
    runIndices.add(run.sourceRunIndex);
  }
  return index;
}

function renderedTextLocations(
  paragraphs: readonly ParagraphVisit[],
  renderedRunIndex: RenderedRunIndex,
): ProjectedTextIndex {
  const locations = paragraphs.flatMap(({ paragraph, source }, paragraphIndex) => {
    const runCount = normalizedBoundary(paragraph, paragraph.runs.length);
    return [...(renderedRunIndex.get(sourceKey(source)) ?? [])]
      .sort((left, right) => left - right)
      .map((boundary) => ({ paragraphIndex, source, boundary, runCount }));
  });
  return {
    locations,
    paragraphIndices: new Set(locations.map(({ paragraphIndex }) => paragraphIndex)),
  };
}

interface ProjectedTextIndex {
  readonly locations: readonly MarkLocation[];
  readonly paragraphIndices: ReadonlySet<number>;
}

function compareMarkLocations(left: MarkLocation, right: MarkLocation): number {
  return left.paragraphIndex - right.paragraphIndex || left.boundary - right.boundary;
}

/** First projected location at or after `reference`, using the document-order
 * index produced above. Exported only for focused complexity tests; it is not
 * part of the package entrypoint. */
function lowerBoundProjectedTextLocation(
  locations: readonly MarkLocation[],
  reference: MarkLocation,
): number {
  let low = 0;
  let high = locations.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    if (compareMarkLocations(locations[middle]!, reference) < 0) low = middle + 1;
    else high = middle;
  }
  return low;
}

function nearestProjectedText(
  reference: MarkLocation,
  index: ProjectedTextIndex,
): MarkLocation | undefined {
  // Same-paragraph gaps are resolved from the authored reference boundary by
  // the consumer helper below. Publish a cross-paragraph fallback only when
  // the entire paragraph has no final-state text geometry.
  if (index.paragraphIndices.has(reference.paragraphIndex)) return undefined;
  const followingIndex = lowerBoundProjectedTextLocation(index.locations, reference);
  const following = index.locations[followingIndex];
  const preceding = followingIndex > 0 ? index.locations[followingIndex - 1] : undefined;
  return frozenPoint(reference).affinity === 'following'
    ? following ?? preceding
    : preceding ?? following;
}

/** Decide, once for a laid-out document, which range of each comment with
 * several ranges carries its margin card: the first range, in document
 * order, of which some page shows something (its runs have geometry inside
 * their clips there). Every page then draws the card with that range and no
 * other, so a comment whose first paragraph is clipped away, on this page or
 * on another, still has one card. `sourcePages` names the pages that hold
 * runs of each paragraph source and `runsForPage` projects a page; only the
 * pages of undecided ranges are projected, each once. Ranges are returned in
 * their order, the owners marked (`cardOwner`). */
export function assignCommentCardOwners(
  anchors: readonly CommentAnchorRange[],
  sourcePages: ReadonlyMap<string, readonly number[]>,
  runsForPage: (pageIndex: number) => readonly Readonly<DocxTextRunInfo>[],
): readonly CommentAnchorRange[] {
  const byComment = new Map<string, number[]>();
  anchors.forEach((anchor, index) => {
    const indices = byComment.get(anchor.commentId);
    if (indices) indices.push(index);
    else byComment.set(anchor.commentId, [index]);
  });
  const projected = new Map<number, readonly Readonly<DocxTextRunInfo>[]>();
  const owners = new Set<number>();
  for (const indices of byComment.values()) {
    if (indices.length < 2) continue;
    owner: for (const index of indices) {
      const anchor = anchors[index]!;
      const pages = new Set<number>();
      for (const source of [anchor.source, anchor.reference.source, anchor.geometryFallback?.source]) {
        for (const page of source ? sourcePages.get(sourceKey(source)) ?? [] : []) pages.add(page);
      }
      for (const page of [...pages].sort((left, right) => left - right)) {
        let runs = projected.get(page);
        if (!runs) projected.set(page, runs = runsForPage(page));
        if (overlayBoxes(resolveCommentAnchorRuns(anchor, runs)).length > 0) {
          owners.add(index);
          break owner;
        }
      }
    }
  }
  if (owners.size === 0) return anchors;
  return Object.freeze(anchors.map((anchor, index) => (
    owners.has(index) ? Object.freeze({ ...anchor, cardOwner: true as const }) : anchor)));
}

/** Resolve one document story's §17.13.4 marks. Kept module-public for focused
 * tests; the package API exposes `DocxDocument.commentAnchorRanges()` instead.
 * Per §17.13.4.3/.4, a lone rangeStart or rangeEnd is a single anchor point. */
export function collectStoryCommentRanges(
  paragraphs: readonly ParagraphVisit[],
  validCommentIds: ReadonlySet<string>,
  renderedRuns: readonly Readonly<DocxTextRunInfo>[] = [],
): CommentAnchorRange[] {
  return collectIndexedStoryCommentRanges(
    paragraphs,
    validCommentIds,
    indexRenderedRuns(renderedRuns),
  );
}

function collectIndexedStoryCommentRanges(
  paragraphs: readonly ParagraphVisit[],
  validCommentIds: ReadonlySet<string>,
  renderedRunIndex: RenderedRunIndex,
): CommentAnchorRange[] {
  const validByValue = new Map<string, string>();
  const ambiguousValues = new Set<string>();
  for (const id of validCommentIds) {
    const key = decimalReviewIdKey(id);
    if (key === undefined || ambiguousValues.has(key)) continue;
    if (validByValue.has(key)) {
      validByValue.delete(key);
      ambiguousValues.add(key);
    } else {
      validByValue.set(key, id);
    }
  }
  const marksById = new Map<string, {
    starts: MarkLocation[];
    ends: MarkLocation[];
    references: MarkLocation[];
  }>();
  for (const [paragraphIndex, { paragraph, source }] of paragraphs.entries()) {
    for (const mark of paragraph.commentMarks ?? []) {
      const key = decimalReviewIdKey(mark.id);
      if (key === undefined) continue;
      const byKind = marksById.get(key) ?? { starts: [], ends: [], references: [] };
      marksById.set(key, byKind);
      const location = {
        paragraphIndex,
        source,
        boundary: normalizedBoundary(paragraph, mark.runIndex),
        runCount: normalizedBoundary(paragraph, paragraph.runs.length),
      };
      if (mark.kind === 'rangeStart') byKind.starts.push(location);
      else if (mark.kind === 'rangeEnd') byKind.ends.push(location);
      else if (mark.kind === 'reference') byKind.references.push(location);
    }
  }

  const ranges: CommentAnchorRange[] = [];
  const projectedTextIndex = renderedTextLocations(paragraphs, renderedRunIndex);
  for (const [valueKey, marks] of marksById) {
    const id = validByValue.get(valueKey);
    if (id === undefined || marks.references.length !== 1) continue;
    const reference = marks.references[0]!;
    // Duplicate or reversed boundaries are non-conformant and ambiguous. Do
    // not guess a range from them merely because a reference exists.
    if (marks.starts.length > 1 || marks.ends.length > 1) continue;
    const start = marks.starts.length === 1 ? marks.starts[0] : undefined;
    const end = marks.ends.length === 1 ? marks.ends[0] : undefined;
    const orderedPair = start && end && (
      start.paragraphIndex < end.paragraphIndex
      || (start.paragraphIndex === end.paragraphIndex && start.boundary <= end.boundary)
    );
    if (orderedPair) {
      const fallbackOrigin = start.paragraphIndex === end.paragraphIndex
        && start.boundary === end.boundary ? start : reference;
      const geometryFallback = nearestProjectedText(fallbackOrigin, projectedTextIndex);
      // One segment per paragraph of the range. A segment that covers no run
      // (the range starts at the end of a paragraph, ends at the start of
      // one, or crosses an empty one) marks nothing beside the segments that
      // do: kept, it would resolve as a point to the run before it, outside
      // the range. A range with no covered run at all stays one point.
      const segments: CommentAnchorRange[] = [];
      for (let index = start.paragraphIndex; index <= end.paragraphIndex; index += 1) {
        const visit = paragraphs[index]!;
        const from = index === start.paragraphIndex ? start.boundary : 0;
        const to = index === end.paragraphIndex
          ? end.boundary
          : normalizedBoundary(visit.paragraph, visit.paragraph.runs.length);
        // Whether a covered run is laid out anywhere in the document: the
        // range then belongs to the pages that hold those runs.
        const rendered = renderedRunIndex.get(sourceKey(visit.source));
        let coveredGeometry = false;
        if (rendered) {
          for (let runIndex = from; runIndex < to && !coveredGeometry; runIndex += 1) {
            coveredGeometry = rendered.has(runIndex);
          }
        }
        segments.push(frozenRange(id, {
          paragraphIndex: index,
          source: visit.source,
          boundary: from,
          runCount: normalizedBoundary(visit.paragraph, visit.paragraph.runs.length),
        }, reference, geometryFallback, Math.max(from, to), coveredGeometry));
      }
      const covering = segments.filter((segment) => segment.endRunIndex > segment.startRunIndex);
      ranges.push(...(covering.length > 0 ? covering : segments.slice(0, 1)));
      continue;
    }

    const loneBoundary = start && marks.ends.length === 0
      ? start
      : end && marks.starts.length === 0 ? end : undefined;
    // A reversed start/end pair is not a lone boundary.
    if (start && end) continue;
    const point = loneBoundary ?? reference;
    const geometryFallback = nearestProjectedText(point, projectedTextIndex);
    ranges.push(frozenRange(id, point, reference, geometryFallback));
  }
  return ranges;
}

/** Resolve comment ranges in every retained story (body, headers, footers,
 * notes, and text boxes) using the same source identities as rendered runs. */
export function collectLayoutSourceCommentRanges(
  comments: readonly Readonly<{ id: string }>[],
  source: LayoutSourceStore,
  renderedRunIndex: RenderedRunIndex = new Map(),
): CommentAnchorRange[] {
  const validCommentIds = new Set(comments.map(({ id }) => id));
  const byStory = new Map<string, ParagraphVisit[]>();
  for (const blockSource of source.blocks.sources) {
    const block = source.blocks.resolve(blockSource);
    if (block.type !== 'paragraph') continue;
    const key = `${blockSource.story}\u0000${blockSource.storyInstance}`;
    const visits = byStory.get(key) ?? [];
    if (!byStory.has(key)) byStory.set(key, visits);
    visits.push({ paragraph: block, source: blockSource });
  }
  return [...byStory.values()].flatMap((paragraphs) =>
    collectIndexedStoryCommentRanges(paragraphs, validCommentIds, renderedRunIndex));
}

/** Avoid any story walk for the common comment-free document. */
export function collectLayoutSourceCommentRangesIfPresent(
  comments: readonly Readonly<{ id: string }>[] | undefined,
  source: LayoutSourceStore,
  renderedRunIndex: RenderedRunIndex = new Map(),
  options: ReviewAnchorProjectionOptions = {},
): CommentAnchorRange[] {
  if ((comments?.length ?? 0) === 0) return [];
  const collected = collectLayoutSourceCommentRanges(comments ?? [], source, renderedRunIndex);
  const completed = options.completedSourceKeys;
  const ranges = completed === undefined ? collected : collected.filter((range) => {
    const sourceIndices = renderedRunIndex.get(sourceKey(range.source)) ?? new Set<number>();
    const hasCoveredRun = range.startRunIndex < range.endRunIndex
      ? [...sourceIndices].some((index) =>
          index >= range.startRunIndex && index < range.endRunIndex)
      : sourceIndices.has(
          range.reference.affinity === 'following'
            ? range.startRunIndex
            : range.startRunIndex - 1,
        );
    const referenceIndices = renderedRunIndex.get(sourceKey(range.reference.source));
    const hasReferenceRun = referenceIndices?.has(
      range.reference.affinity === 'following'
        ? range.reference.runIndex
        : range.reference.runIndex - 1,
    ) === true;
    return hasCoveredRun
      || hasReferenceRun
      || completed.has(sourceKey(range.source))
      || completed.has(sourceKey(range.reference.source));
  });
  const owners = options.cardOwners;
  return owners ? [...assignCommentCardOwners(ranges, owners.sourcePages, owners.runsForPage)] : ranges;
}

function sameStorySource(
  left: Readonly<DocxStorySource> | undefined,
  right: Readonly<DocxStorySource>,
): boolean {
  return left !== undefined
    && left.story === right.story
    && left.storyInstance === right.storyInstance
    && left.path.length === right.path.length
    && left.path.every((value, index) => value === right.path[index]);
}

/** Join one structural comment anchor to projected page text. Covered runs win;
 * a point/gapped anchor then selects the nearest authored-boundary run, and a
 * paragraph with no final-state text uses its deterministic adjacent-paragraph
 * fallback. Split geometry for the selected source run remains intact. */
export function resolveCommentAnchorRuns(
  anchor: Readonly<CommentAnchorRange>,
  runs: readonly Readonly<DocxTextRunInfo>[],
): readonly Readonly<DocxTextRunInfo>[] {
  const addressable = runs.filter((run) =>
    run.sourceRunIndex !== undefined && sameStorySource(run.source, anchor.source));
  const covered = addressable.filter((run) =>
    (run.sourceRunIndex as number) >= anchor.startRunIndex
    && (run.sourceRunIndex as number) < anchor.endRunIndex);
  if (covered.length > 0) return covered;
  // The covered runs are laid out on another page: nothing of the range is
  // here, and its reference boundary does not stand in for it.
  if (anchor.coveredGeometry && anchor.startRunIndex < anchor.endRunIndex) return [];

  // A lone rangeStart/rangeEnd or reference-only comment is a point anchor at
  // the range's own source/boundary (§17.13.4.3/.4). The separately-authored
  // commentReference may be elsewhere and must not displace that point.
  if (anchor.startRunIndex === anchor.endRunIndex && addressable.length > 0) {
    const followingIndex = addressable.reduce<number | undefined>((nearest, run) => {
      const index = run.sourceRunIndex as number;
      if (index < anchor.startRunIndex) return nearest;
      return nearest === undefined || index < nearest ? index : nearest;
    }, undefined);
    const precedingIndex = addressable.reduce<number | undefined>((nearest, run) => {
      const index = run.sourceRunIndex as number;
      if (index >= anchor.startRunIndex) return nearest;
      return nearest === undefined || index > nearest ? index : nearest;
    }, undefined);
    const selectedIndex = followingIndex ?? precedingIndex;
    if (selectedIndex !== undefined) {
      return addressable.filter((run) => run.sourceRunIndex === selectedIndex);
    }
  }

  const referenceRuns = runs.filter((run) =>
    run.sourceRunIndex !== undefined && sameStorySource(run.source, anchor.reference.source));
  const followingIndex = referenceRuns.reduce<number | undefined>((nearest, run) => {
    const index = run.sourceRunIndex as number;
    if (index < anchor.reference.runIndex) return nearest;
    return nearest === undefined || index < nearest ? index : nearest;
  }, undefined);
  const precedingIndex = referenceRuns.reduce<number | undefined>((nearest, run) => {
    const index = run.sourceRunIndex as number;
    if (index >= anchor.reference.runIndex) return nearest;
    return nearest === undefined || index > nearest ? index : nearest;
  }, undefined);
  const selectedIndex = anchor.reference.affinity === 'following'
    ? followingIndex ?? precedingIndex
    : precedingIndex ?? followingIndex;
  if (selectedIndex !== undefined) {
    return referenceRuns.filter((run) => run.sourceRunIndex === selectedIndex);
  }

  const fallback = anchor.geometryFallback;
  if (fallback === undefined) return [];
  return runs.filter((run) =>
    run.sourceRunIndex === fallback.sourceRunIndex
    && sameStorySource(run.source, fallback.source));
}

function resolvedAnchorKind(
  anchor: Readonly<CommentAnchorRange>,
  runs: readonly Readonly<DocxTextRunInfo>[],
): DocxCommentAnchorKind {
  const covered = runs.some((run) =>
    run.sourceRunIndex !== undefined
    && sameStorySource(run.source, anchor.source)
    && run.sourceRunIndex >= anchor.startRunIndex
    && run.sourceRunIndex < anchor.endRunIndex);
  if (covered) return 'range';
  const authoredPoint = runs.some((run) =>
    sameStorySource(run.source, anchor.source)
    || sameStorySource(run.source, anchor.reference.source));
  return authoredPoint ? 'point' : 'fallback';
}

function highlightRects(
  runs: readonly Readonly<DocxTextRunInfo>[],
): readonly Readonly<DocxCommentHighlightRect>[] {
  // The same boxes as the viewer's tint (overlay-geometry.ts): the glyphs of
  // the anchored runs inside their clips, joined only where they touch.
  return Object.freeze(overlayBoxes(runs)
    .sort((left, right) => left.y - right.y || left.x - right.x)
    .map((box) => Object.freeze(box)));
}

/** Resolve page-visible DOCX comments into UI-neutral threads and continuous
 * highlight rectangles. This absorbs thread ancestry, anchor joining, Word's
 * text-highlight bounds, and the joining of touching highlights without owning any DOM,
 * selection state, styling, or framework lifecycle. Invalid orphan/cyclic
 * replies are not promoted to roots. */
export function resolveDocxCommentThreads(
  comments: readonly Readonly<DocComment>[],
  anchors: readonly Readonly<CommentAnchorRange>[],
  runs: readonly Readonly<DocxTextRunInfo>[],
  options: ResolveDocxCommentThreadsOptions = {},
): readonly Readonly<ResolvedDocxCommentThread>[] {
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  const roots = comments.filter((comment) => comment.parentId === undefined);
  const rootById = new Map(roots.map((root) => [root.id, root]));
  const replies = new Map<string, Readonly<DocComment>[]>();
  const rootIdByCommentId = new Map(roots.map((root) => [root.id, root.id]));
  for (const comment of comments) {
    if (comment.parentId === undefined) continue;
    const seen = new Set<string>([comment.id]);
    let current: Readonly<DocComment> = comment;
    while (current.parentId !== undefined) {
      const parent = byId.get(current.parentId);
      if (!parent || seen.has(parent.id)) {
        current = comment;
        break;
      }
      seen.add(parent.id);
      current = parent;
    }
    if (!rootById.has(current.id) || current === comment) continue;
    rootIdByCommentId.set(comment.id, current.id);
    const group = replies.get(current.id) ?? [];
    if (!replies.has(current.id)) replies.set(current.id, group);
    group.push(comment);
  }

  const pageSources = new Set<string>();
  for (const run of runs) {
    if (run.source) pageSources.add(sourceKey(run.source));
  }
  const resolvedByRoot = new Map<string, ResolvedDocxCommentAnchor[]>();
  for (const anchor of anchors) {
    const rootId = rootIdByCommentId.get(anchor.commentId);
    if (rootId === undefined) continue;
    const mayResolve = pageSources.has(sourceKey(anchor.source)) ||
      pageSources.has(sourceKey(anchor.reference.source)) ||
      (anchor.geometryFallback !== undefined &&
        pageSources.has(sourceKey(anchor.geometryFallback.source)));
    if (!mayResolve) continue;
    const resolvedRuns = resolveCommentAnchorRuns(anchor, runs);
    if (resolvedRuns.length === 0) continue;
    // An anchor whose content the page hides whole (clipped away) is not
    // visible on it: it has no highlight, and is no anchor of the page.
    const rects = highlightRects(resolvedRuns);
    if (rects.length === 0) continue;
    const resolved = resolvedByRoot.get(rootId) ?? [];
    if (!resolvedByRoot.has(rootId)) resolvedByRoot.set(rootId, resolved);
    resolved.push(Object.freeze({
      anchor,
      kind: resolvedAnchorKind(anchor, resolvedRuns),
      rects,
    }));
  }

  return Object.freeze(roots.flatMap((root): ResolvedDocxCommentThread[] => {
    if (options.includeResolved === false && root.resolved === true) return [];
    const resolved = resolvedByRoot.get(root.id);
    if (!resolved?.length) return [];
    return [Object.freeze({
      root,
      replies: Object.freeze([...(replies.get(root.id) ?? [])]),
      anchors: Object.freeze([...resolved]),
    })];
  }));
}
