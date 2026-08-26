import type { LayoutSourceStore } from './layout/layout-source-store.js';
import { sourceKey } from './layout/source-key.js';
import type { SourceRef } from './layout/types.js';
import { decimalReviewIdKey } from './review-id.js';
import type { DocComment, DocxCommentMark, DocxStorySource, DocxTextRunInfo } from './types.js';

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

/** UI-neutral rectangle for one continuous highlighted line segment. */
export interface DocxCommentHighlightRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly transform?: string;
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
export function lowerBoundProjectedTextLocation(
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
      for (let index = start.paragraphIndex; index <= end.paragraphIndex; index += 1) {
        const visit = paragraphs[index]!;
        const from = index === start.paragraphIndex ? start.boundary : 0;
        const to = index === end.paragraphIndex
          ? end.boundary
          : normalizedBoundary(visit.paragraph, visit.paragraph.runs.length);
        ranges.push(frozenRange(id, {
          paragraphIndex: index,
          source: visit.source,
          boundary: from,
          runCount: normalizedBoundary(visit.paragraph, visit.paragraph.runs.length),
        }, reference, geometryFallback, Math.max(from, to)));
      }
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
): CommentAnchorRange[] {
  if ((comments?.length ?? 0) === 0) return [];
  return collectLayoutSourceCommentRanges(comments ?? [], source, renderedRunIndex);
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
  const rects = runs
    .map((run): DocxCommentHighlightRect => {
      const bounds = run.highlightBounds;
      return {
        x: bounds?.x ?? run.x,
        y: bounds?.y ?? run.y,
        width: bounds?.width ?? run.w,
        height: bounds?.height ?? run.h,
        ...(run.transform ? { transform: run.transform } : {}),
      };
    })
    .filter(({ width, height }) => width > 0 && height > 0)
    .sort((left, right) => left.y - right.y || left.x - right.x);
  const merged: DocxCommentHighlightRect[] = [];
  for (const rect of rects) {
    const previous = merged.at(-1);
    if (
      previous
      && previous.y === rect.y
      && previous.height === rect.height
      && previous.transform === rect.transform
    ) {
      const left = Math.min(previous.x, rect.x);
      const right = Math.max(previous.x + previous.width, rect.x + rect.width);
      merged[merged.length - 1] = Object.freeze({
        ...previous,
        x: left,
        width: right - left,
      });
    } else {
      merged.push(Object.freeze({ ...rect }));
    }
  }
  return Object.freeze(merged);
}

/** Resolve page-visible DOCX comments into UI-neutral threads and continuous
 * highlight rectangles. This absorbs thread ancestry, anchor joining, Word's
 * text-highlight bounds, and same-line run merging without owning any DOM,
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
    const resolved = resolvedByRoot.get(rootId) ?? [];
    if (!resolvedByRoot.has(rootId)) resolvedByRoot.set(rootId, resolved);
    resolved.push(Object.freeze({
      anchor,
      kind: resolvedAnchorKind(anchor, resolvedRuns),
      rects: highlightRects(resolvedRuns),
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
