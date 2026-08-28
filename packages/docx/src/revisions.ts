import type { LayoutSourceStore } from './layout/layout-source-store.js';
import { sourceKey } from './layout/source-key.js';
import type { SourceRef } from './layout/types.js';
import { decimalReviewIdKey } from './review-id.js';
import type { DocRevision, DocxStorySource, DocxTextRunInfo } from './types.js';
import type { ReviewAnchorProjectionOptions } from './review-projection.js';

export interface RevisionAnchorRange {
  /** Index into `DocxDocument.revisions`, preserving authored document order. */
  readonly revisionIndex: number;
  /** Exact story/path identity used by `DocxTextRunInfo.source`. */
  readonly source: Readonly<DocxStorySource>;
  /** First normalized-model run covered by the revision, inclusive. */
  readonly startRunIndex: number;
  /** One past the last normalized-model run covered by the revision. */
  readonly endRunIndex: number;
  /** Nearest final-state text-bearing run when the revised content itself has
   * no geometry, as for `w:del` and `w:moveFrom`. */
  readonly geometryFallback?: RevisionAnchorGeometryFallback;
}

export interface RevisionAnchorGeometryFallback {
  readonly source: Readonly<DocxStorySource>;
  readonly sourceRunIndex: number;
}

type RenderedRunIndex = ReadonlyMap<string, ReadonlySet<number>>;
type SortedRenderedRunIndex = ReadonlyMap<string, readonly number[]>;

interface RevisionParagraph {
  readonly source: SourceRef;
  readonly runs: readonly Readonly<{
    type?: string;
    revision?: Readonly<{ kind?: string; id?: string }>;
  }>[];
}

function revisionKey(revision: Readonly<{ kind: string; id?: string }>): string | undefined {
  const id = decimalReviewIdKey(revision.id);
  return id === undefined ? undefined : `${revision.kind}\u0000${id}`;
}

function frozenSource(source: SourceRef): Readonly<DocxStorySource> {
  return Object.freeze({ ...source, path: Object.freeze([...source.path]) });
}

function sameSource(
  left: Readonly<DocxStorySource> | undefined,
  right: Readonly<DocxStorySource>,
): boolean {
  return left !== undefined
    && left.story === right.story
    && left.storyInstance === right.storyInstance
    && left.path.length === right.path.length
    && left.path.every((value, index) => value === right.path[index]);
}

function lowerBound(values: readonly number[], target: number): number {
  let low = 0;
  let high = values.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    if (values[middle]! < target) low = middle + 1;
    else high = middle;
  }
  return low;
}

interface ParagraphGeometryFallbacks {
  readonly following: readonly (RevisionAnchorGeometryFallback | undefined)[];
  readonly preceding: readonly (RevisionAnchorGeometryFallback | undefined)[];
}

function paragraphGeometryFallbacks(
  paragraphs: readonly RevisionParagraph[],
  renderedRunIndex: SortedRenderedRunIndex,
): ParagraphGeometryFallbacks {
  const following: (RevisionAnchorGeometryFallback | undefined)[] = new Array(paragraphs.length);
  let next: RevisionAnchorGeometryFallback | undefined;
  for (let index = paragraphs.length - 1; index >= 0; index -= 1) {
    following[index] = next;
    const paragraph = paragraphs[index]!;
    const first = renderedRunIndex.get(sourceKey(paragraph.source))?.[0];
    if (first !== undefined) {
      next = Object.freeze({ source: frozenSource(paragraph.source), sourceRunIndex: first });
    }
  }

  const preceding: (RevisionAnchorGeometryFallback | undefined)[] = new Array(paragraphs.length);
  let previous: RevisionAnchorGeometryFallback | undefined;
  for (let index = 0; index < paragraphs.length; index += 1) {
    preceding[index] = previous;
    const paragraph = paragraphs[index]!;
    const runs = renderedRunIndex.get(sourceKey(paragraph.source));
    const last = runs?.at(-1);
    if (last !== undefined) {
      previous = Object.freeze({ source: frozenSource(paragraph.source), sourceRunIndex: last });
    }
  }
  return { following, preceding };
}

function nearestRenderedRun(
  paragraphIndex: number,
  startRunIndex: number,
  endRunIndex: number,
  paragraphs: readonly RevisionParagraph[],
  renderedRunIndex: SortedRenderedRunIndex,
  paragraphFallbacks: ParagraphGeometryFallbacks,
): RevisionAnchorGeometryFallback | undefined {
  const paragraph = paragraphs[paragraphIndex]!;
  const sameParagraph = renderedRunIndex.get(sourceKey(paragraph.source)) ?? [];
  const followingOffset = lowerBound(sameParagraph, endRunIndex);
  const following = sameParagraph[followingOffset];
  const precedingOffset = lowerBound(sameParagraph, startRunIndex) - 1;
  const preceding = sameParagraph[precedingOffset];
  const sameParagraphIndex = following ?? preceding;
  if (sameParagraphIndex !== undefined) {
    return Object.freeze({
      source: frozenSource(paragraph.source),
      sourceRunIndex: sameParagraphIndex,
    });
  }

  return paragraphFallbacks.following[paragraphIndex]
    ?? paragraphFallbacks.preceding[paragraphIndex];
}

/** Resolve body revision containers to normalized source-run ranges. Valid
 * `w:id` values are the normative identity; duplicate or malformed ids are
 * ambiguous and deliberately remain available only through `revisions`. */
export function collectLayoutSourceRevisionRanges(
  revisions: readonly Readonly<DocRevision>[],
  source: LayoutSourceStore,
  renderedRunIndex: RenderedRunIndex = new Map(),
): RevisionAnchorRange[] {
  const revisionIndexByKey = new Map<string, number>();
  const ambiguousKeys = new Set<string>();
  revisions.forEach((revision, index) => {
    const key = revisionKey(revision);
    if (key === undefined || ambiguousKeys.has(key)) return;
    if (revisionIndexByKey.has(key)) {
      revisionIndexByKey.delete(key);
      ambiguousKeys.add(key);
    } else {
      revisionIndexByKey.set(key, index);
    }
  });

  const paragraphs: RevisionParagraph[] = source.blocks.sources.flatMap((blockSource) => {
    if (blockSource.story !== 'body') return [];
    const block = source.blocks.resolve(blockSource);
    return block.type === 'paragraph' ? [{ source: blockSource, runs: block.runs }] : [];
  });
  const sortedRenderedRunIndex: SortedRenderedRunIndex = new Map(
    [...renderedRunIndex].map(([key, indices]) => [
      key,
      Object.freeze([...indices].sort((left, right) => left - right)),
    ]),
  );
  const paragraphFallbacks = paragraphGeometryFallbacks(paragraphs, sortedRenderedRunIndex);
  const ranges: RevisionAnchorRange[] = [];
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    let runIndex = 0;
    while (runIndex < paragraph.runs.length) {
      const run = paragraph.runs[runIndex]!;
      const key = run.revision?.kind
        ? revisionKey({ kind: run.revision.kind, id: run.revision.id })
        : undefined;
      const revisionIndex = key === undefined ? undefined : revisionIndexByKey.get(key);
      if (revisionIndex === undefined) {
        runIndex += 1;
        continue;
      }
      const startRunIndex = runIndex;
      runIndex += 1;
      while (runIndex < paragraph.runs.length) {
        const next = paragraph.runs[runIndex]!.revision;
        if (!next?.kind || revisionKey({ kind: next.kind, id: next.id }) !== key) break;
        runIndex += 1;
      }
      const endRunIndex = runIndex;
      const rendered = sortedRenderedRunIndex.get(sourceKey(paragraph.source)) ?? [];
      const firstCoveredOffset = lowerBound(rendered, startRunIndex);
      const hasOwnGeometry = (rendered[firstCoveredOffset] ?? endRunIndex) < endRunIndex;
      const geometryFallback = hasOwnGeometry
        ? undefined
        : nearestRenderedRun(
            paragraphIndex,
            startRunIndex,
            endRunIndex,
            paragraphs,
            sortedRenderedRunIndex,
            paragraphFallbacks,
          );
      ranges.push(Object.freeze({
        revisionIndex,
        source: frozenSource(paragraph.source),
        startRunIndex,
        endRunIndex,
        ...(geometryFallback === undefined ? {} : { geometryFallback }),
      }));
    }
  }
  return ranges;
}

export function collectLayoutSourceRevisionRangesIfPresent(
  revisions: readonly Readonly<DocRevision>[] | undefined,
  source: LayoutSourceStore,
  renderedRunIndex: RenderedRunIndex = new Map(),
  options: ReviewAnchorProjectionOptions = {},
): RevisionAnchorRange[] {
  if ((revisions?.length ?? 0) === 0) return [];
  const ranges = collectLayoutSourceRevisionRanges(revisions ?? [], source, renderedRunIndex);
  const completed = options.completedSourceKeys;
  if (completed === undefined) return ranges;
  return ranges.filter((range) => {
    const indices = renderedRunIndex.get(sourceKey(range.source));
    const hasCoveredRun = indices !== undefined && [...indices].some((index) =>
      index >= range.startRunIndex && index < range.endRunIndex);
    return hasCoveredRun || completed.has(sourceKey(range.source));
  });
}

/** Join one revision range to projected final-state text. Insertions and move
 * destinations resolve to their own geometry; deletions and move sources use
 * the deterministic adjacent final-state run published with the range. */
export function resolveRevisionAnchorRuns(
  anchor: Readonly<RevisionAnchorRange>,
  runs: readonly Readonly<DocxTextRunInfo>[],
): readonly Readonly<DocxTextRunInfo>[] {
  const covered = runs.filter((run) => run.sourceRunIndex !== undefined
    && sameSource(run.source, anchor.source)
    && (run.sourceRunIndex as number) >= anchor.startRunIndex
    && (run.sourceRunIndex as number) < anchor.endRunIndex);
  if (covered.length > 0) return covered;
  const fallback = anchor.geometryFallback;
  if (fallback === undefined) return [];
  return runs.filter((run) => run.sourceRunIndex === fallback.sourceRunIndex
    && sameSource(run.source, fallback.source));
}
