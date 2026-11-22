import { isFloatingPayloadRun } from './anchor-classification.js';
import { runIsOmitted } from './revision-visibility.js';

/**
 * WORD_ANCHOR_HOST_NO_LINE_METRIC, by construction.
 *
 * Under that rule a paragraph is sized and classified from its
 * revision-visible inline content only: a floating drawing (ECMA-376
 * §20.4.2.3 wp:anchor and its VML/legacy/ODF equivalents) and the invisible
 * character that anchors it contribute no line metric, and content the
 * active revision view omits (§17.13.5.14 deletions, §17.13.5.22 moved-away
 * content) emits nothing.
 *
 * Layout therefore reads a paragraph through two separate channels:
 *
 * - the METRIC-VISIBLE runs (this projection): every decision about line
 *   metrics, leading rules, grid class, content facts, mark-only
 *   classification and page-edge admission reads these, so a paragraph with a
 *   floating anchor is decided exactly as the same paragraph without it;
 * - the ANCHORING runs: carried in the source run array (source run indices
 *   stay stable) and in the segment stream as position-only items
 *   (`isAnchorPlacementSegment`), used for placement only.
 *
 * A decision that reads `paragraph.runs` directly instead of this projection
 * lets an anchor influence it. Do that only where the anchor itself is the
 * subject (placement, wrap exclusions, resource acquisition, source maps).
 */

type RunFacts = Readonly<{
  type: string;
  anchor?: boolean;
  inline?: boolean;
  listLabel?: unknown;
  __listLabel?: unknown;
  anchorAcquisitionInput?: unknown;
  __anchorAcquisition?: unknown;
  revision?: Readonly<{ kind?: string }> | null;
}>;

/** A floating drawing payload or the anchor-host run that marks its anchoring
 * character. A picture list label's metric run (`listLabel`) has the
 * anchor-host type but is the label font's line metric, not an anchor. */
export function isAnchoringRun(run: RunFacts): boolean {
  if (run.type === 'anchorHost') return run.listLabel === undefined && run.__listLabel === undefined;
  return isFloatingPayloadRun(run);
}

type Projections = { anchorFree?: readonly unknown[]; final?: readonly unknown[]; markup?: readonly unknown[] };
const projections = new WeakMap<readonly unknown[], Projections>();

/** The retained projections of a run array. Only a frozen array (every
 * retained layout source) can keep them: a caller-owned array may change
 * between two layouts, so its projections are computed on each use. */
function projectionsOf(runs: readonly unknown[]): Projections {
  if (!Object.isFrozen(runs)) return {};
  let cached = projections.get(runs);
  if (!cached) projections.set(runs, cached = {});
  return cached;
}

/** `runs` itself when the predicate removes nothing, so an ordinary paragraph
 * allocates nothing and keeps its array identity. */
function without<R>(runs: readonly R[], removed: (run: R) => boolean): readonly R[] {
  let kept: R[] | undefined;
  for (let index = 0; index < runs.length; index += 1) {
    const run = runs[index]!;
    if (removed(run)) kept ??= runs.slice(0, index);
    else kept?.push(run);
  }
  return kept ? Object.freeze(kept) : runs;
}

/** The runs without floating drawings and their anchoring characters, in
 * every revision view. For source-level classifications that are decided once
 * for all views. Computed once per retained (frozen) run array. */
export function anchorFreeRuns<R extends RunFacts>(runs: readonly R[]): readonly R[] {
  const cached = projectionsOf(runs);
  return (cached.anchorFree ??= without(runs, isAnchoringRun)) as readonly R[];
}

/** The metric-visible runs of a paragraph in the given revision view: no
 * floating drawing, no anchoring character, nothing the view omits. Computed
 * once per retained (frozen) run array and view. */
export function metricVisibleRuns<R extends RunFacts>(
  runs: readonly R[],
  showTrackedChanges: boolean | undefined,
): readonly R[] {
  const cached = projectionsOf(runs);
  const view = showTrackedChanges === true ? 'markup' : 'final';
  return (cached[view] ??= without(anchorFreeRuns(runs),
    (run) => runIsOmitted(run, showTrackedChanges))) as readonly R[];
}

/** Whether the paragraph anchors a floating drawing the view displays. */
export function hasVisibleAnchoringRun(
  runs: readonly RunFacts[],
  showTrackedChanges: boolean | undefined,
): boolean {
  return anchorFreeRuns(runs) !== runs
    && runs.some((run) => isAnchoringRun(run) && !runIsOmitted(run, showTrackedChanges));
}

declare const metricVisible: unique symbol;
/** A paragraph whose `runs` are its metric-visible runs. Only
 * {@link metricVisibleParagraph} produces one, so a function that takes this
 * type cannot be handed a paragraph that still carries anchors. */
export type MetricVisibleParagraph<P> = P & { readonly [metricVisible]: true };

const paragraphs = new WeakMap<object, { final?: object; markup?: object }>();

/** The paragraph as line metrics see it: the same mark and properties over
 * its metric-visible runs. The paragraph itself when nothing is removed. */
export function metricVisibleParagraph<P extends Readonly<{ runs: readonly RunFacts[] }>>(
  paragraph: P,
  showTrackedChanges: boolean | undefined,
): MetricVisibleParagraph<P> {
  const runs = metricVisibleRuns(paragraph.runs, showTrackedChanges);
  if (runs === paragraph.runs) return paragraph as MetricVisibleParagraph<P>;
  if (!Object.isFrozen(paragraph)) return { ...paragraph, runs } as MetricVisibleParagraph<P>;
  let cached = paragraphs.get(paragraph);
  if (!cached) paragraphs.set(paragraph, cached = {});
  const view = showTrackedChanges === true ? 'markup' : 'final';
  return (cached[view] ??= Object.freeze({ ...paragraph, runs })) as MetricVisibleParagraph<P>;
}
