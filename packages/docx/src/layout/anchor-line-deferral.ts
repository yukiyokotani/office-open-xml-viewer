import type { PageAnchorPrescanInput } from './body-layout-kernel.js';
import type { ParagraphLayout } from './types.js';
import { WORD_HOST_FLOAT_BOTTOM_MARGIN_DEFERRAL } from './anchor-compatibility.js';

/**
 * Anchor-line deferral for page-owned drawings (issue #1615), governed by the
 * compatibility rule `word-page-anchor-line-deferral` (anchor-compatibility.ts).
 *
 * A page-owned drawing (positionH/V relativeFrom page or margin, ECMA-376
 * §20.4.3.4-5) is registered on the page of its anchor, so text laid out
 * earlier on that page wraps around it (§20.4.2.17 wrapSquare, §20.4.2.20
 * wrapTopAndBottom). When that wrap would push the anchor line L off page N,
 * the rule keeps page N laid out with only the anchors accepted before L,
 * ends the page just above L, and places the drawing on the page L reaches.
 *
 * Body pagination resolves page-owned anchors by exact-state convergence over
 * whole passes, so the counterfactual "page N with the accepted anchors plus
 * the anchors on L" is only observed in a pass that registered exactly that
 * set on N. A proof records that pass's carry reads up to the end of page N. A
 * later pass applies the deferral only after checking that the same carry
 * reads lead to page N with the accepted anchors registered, so it reproduces
 * exactly the proving pass's page-N context minus the deferred anchors.
 */

type PageStartAnchors = PageAnchorPrescanInput['anchors'];

/** One read of anchor-convergence state by a body pagination pass. Everything
 * else a pass reads is identical in every pass of one convergence run. */
export type PageAnchorInputEvent = Readonly<{
  kind: 'prescan';
  pageIndex: number;
  flowDomainId: string;
  anchors: PageStartAnchors;
}> | Readonly<{
  kind: 'page-owned-table';
  pageIndex: number;
  key: string;
  floor: number | undefined;
}> | Readonly<{
  /** A proven anchor-line deferral took effect: the line holding `keys`
   * starts a page after `pageIndex`. Checks that do not apply change nothing
   * and are not reads. */
  kind: 'anchor-line-deferral';
  pageIndex: number;
  keys: readonly string[];
}>;

export type AnchorLineDeferralProof = Readonly<{
  /** Occurrence keys of the page-owned drawings anchored on line L. */
  keys: readonly string[];
  /** Page N, on which registering `keys` pushed L to a later page. */
  pageIndex: number;
  /** Serialized reads of the proving pass (shared by its proofs). */
  prefix: readonly string[];
  /** Index in `prefix` of the prescan that registered `keys` on page N. */
  registrationIndex: number;
  /** That prescan without `keys`: what a deferring pass reads there. */
  expectedRegistration: string;
  /** Serialized later carry reads of the proving pass on page N, in order. */
  inPage: readonly string[];
  identity: string;
  /** A settled floor (`anchorLineFloor`), not a proof of a page context. */
  floor?: true;
}>;

/** Proofs by `anchorLineDeferralKey(anchor, page)`. A proof for one page
 * never replaces another page's proof of the same anchor. */
export type AnchorLineDeferrals = ReadonlyMap<string, AnchorLineDeferralProof>;

export function anchorLineDeferralKey(occurrenceKey: string, pageIndex: number): string {
  return `${pageIndex}\u0000${occurrenceKey}`;
}

/** Key of the settled floor of one anchor (see `anchorLineFloor`). */
export function anchorLineFloorKey(occurrenceKey: string): string {
  return `floor\u0000${occurrenceKey}`;
}

/**
 * A floor settled by the anchor solver's cycle recovery
 * (LIBRARY_ANCHOR_LINE_CYCLE_RECOVERY, library policy): the line anchoring
 * `occurrenceKey` starts no page before `pageIndex`. Unlike a proof it holds
 * in every page context, so it ends any earlier page above that line, unless
 * the line would start that page (the first line of a page stays). Stored
 * under `anchorLineFloorKey` beside the proofs and carried like them.
 */
export function anchorLineFloor(occurrenceKey: string, pageIndex: number): AnchorLineDeferralProof {
  return Object.freeze({
    keys: Object.freeze([occurrenceKey]),
    pageIndex,
    prefix: Object.freeze([]),
    registrationIndex: -1,
    expectedRegistration: '',
    inPage: Object.freeze([]),
    identity: `floor:${occurrenceKey}@${pageIndex}`,
    floor: true,
  });
}

export function anchorKeysId(keys: readonly string[]): string {
  return [...keys].sort().join('\n');
}

export function serializeAnchorInput(
  event: PageAnchorInputEvent,
  anchorsIdentity: (anchors: PageStartAnchors) => string,
): string {
  switch (event.kind) {
    case 'prescan':
      return `P|${event.pageIndex}|${event.flowDomainId}|${anchorsIdentity(event.anchors)}`;
    case 'page-owned-table':
      return `T|${event.pageIndex}|${event.key}|${event.floor ?? ''}`;
    case 'anchor-line-deferral':
      return `D|${event.pageIndex}|${anchorKeysId(event.keys)}`;
  }
}

export function createAnchorLineDeferralProof(
  keys: readonly string[],
  pageIndex: number,
  events: readonly PageAnchorInputEvent[],
  serialized: readonly string[],
  registrationIndex: number,
  anchorsIdentity: (anchors: PageStartAnchors) => string,
): AnchorLineDeferralProof {
  const registration = events[registrationIndex];
  if (registration?.kind !== 'prescan' || registration.pageIndex !== pageIndex) {
    throw new Error('Anchor-line deferral proof must start at its page registration');
  }
  const keySet = new Set(keys);
  const expectedRegistration = serializeAnchorInput(Object.freeze({
    ...registration,
    anchors: Object.freeze(registration.anchors.filter((anchor) => !keySet.has(anchor.occurrenceId))),
  }), anchorsIdentity);
  const inPage: string[] = [];
  for (let index = registrationIndex + 1; index < events.length; index += 1) {
    const event = events[index]!;
    if (event.pageIndex > pageIndex) break;
    if (event.pageIndex === pageIndex) inPage.push(serialized[index]!);
  }
  const sortedKeys = Object.freeze([...keys].sort());
  return Object.freeze({
    keys: sortedKeys,
    pageIndex,
    prefix: serialized,
    registrationIndex,
    expectedRegistration,
    inPage: Object.freeze(inPage),
    // State identity for cycle detection. The prefix can hold one read per
    // page, so it enters as a digest; application still compares it exactly.
    identity: `${sortedKeys.join(',')}@${pageIndex}:${registrationIndex}:${
      digest(serialized, registrationIndex)}\u0002${expectedRegistration}\u0002${
      inPage.join('\u0001')}`,
  });
}

/** Two independent 32-bit FNV-1a lanes over the first `length` reads. */
function digest(values: readonly string[], length: number): string {
  let left = 0x811c9dc5;
  let right = 0x01000193 ^ 0x5bd1e995;
  for (let index = 0; index < length; index += 1) {
    const value = values[index]!;
    for (let offset = 0; offset < value.length; offset += 1) {
      const code = value.charCodeAt(offset);
      left = Math.imul(left ^ code, 0x01000193);
      right = Math.imul(right ^ code, 0x5bd1e995) ^ (right >>> 15);
    }
    left = Math.imul(left ^ 0xff, 0x01000193);
    right = Math.imul(right ^ 0xff, 0x5bd1e995) ^ (right >>> 15);
  }
  return `${(left >>> 0).toString(16)}${(right >>> 0).toString(16)}`;
}

/** What a deferring pass knows about its own reads when it reaches line L. */
export interface AnchorLineDeferralContext {
  /** Serialized reads of this pass so far, in order. */
  readonly reads: readonly string[];
}

/** Whether the proof's counterfactual is exactly this pass's page-N context
 * plus the deferred anchors: identical reads before N's registration, the
 * same registration without them, and after it exactly the proving pass's
 * later reads on page N — none missing, none extra. A different read (for
 * example a page-owned table or section region reached on N only because the
 * anchors are absent) means the proof does not describe this page, so the
 * line is laid out normally and the next pass retests it. */
export function anchorLineDeferralApplies(
  proof: AnchorLineDeferralProof,
  keys: readonly string[],
  pageIndex: number,
  context: AnchorLineDeferralContext,
): boolean {
  if (proof.floor || proof.pageIndex !== pageIndex || anchorKeysId(proof.keys) !== anchorKeysId(keys)) return false;
  const { reads } = context;
  if (reads.length !== proof.registrationIndex + 1 + proof.inPage.length) return false;
  for (let index = 0; index < proof.registrationIndex; index += 1) {
    if (reads[index] !== proof.prefix[index]) return false;
  }
  if (reads[proof.registrationIndex] !== proof.expectedRegistration) return false;
  for (let index = 0; index < proof.inPage.length; index += 1) {
    if (reads[proof.registrationIndex + 1 + index] !== proof.inPage[index]) return false;
  }
  return true;
}

export function anchorLineDeferralsIdentity(deferrals: AnchorLineDeferrals): string {
  return [...new Set(deferrals.values())]
    .map((proof) => proof.identity)
    .sort()
    .join('\u0003');
}

/** Occurrence keys of the page-owned drawings anchored on each line. */
export function pageOwnedAnchorKeysByLine(layout: ParagraphLayout): readonly (readonly string[])[] {
  if (layout.drawings.length === 0) return layout.lines.map(() => []);
  const drawings = new Map(layout.drawings.map((drawing) => [drawing.id, drawing]));
  return layout.lines.map((line) => line.placements.flatMap((placement) => {
    if (placement.kind !== 'drawing') return [];
    const anchor = drawings.get(placement.drawingId)?.anchorLayer;
    if (!anchor || anchor.horizontalOwnership !== 'page' || anchor.verticalOwnership !== 'page') {
      return [];
    }
    return [anchor.acquisitionOccurrenceId ?? anchor.occurrenceId];
  }));
}

/** One host-owned (paragraph- or line-relative) wrapping float anchored on a
 * line: its acquisition key, the top and bottom of its drawing box widened by
 * the authored §20.4.2.6 effect extent, and the bottom of its anchor line box
 * (the whole cell of a §17.6.5 grid line). */
export type HostFloatAnchor = Readonly<{ key: string; topPt: number; bottomPt: number; lineBottomPt: number }>;

/** WORD_HOST_FLOAT_BOTTOM_MARGIN_DEFERRAL: for each line of `layout`, the
 * host-owned square, topAndBottom, tight or through floats anchored on it.
 * The box is the drawing extent (§20.4.2.7) widened by its effect extent
 * (§20.4.2.6), not its wrap exclusion: wrap distances (§20.4.2.3 distB) are
 * left out. */
export function hostWrappingFloatsByLine(layout: ParagraphLayout): readonly (readonly HostFloatAnchor[])[] {
  void WORD_HOST_FLOAT_BOTTOM_MARGIN_DEFERRAL;
  if (layout.drawings.length === 0 || layout.exclusions.length === 0) return layout.lines.map(() => []);
  const wrapping = new Set(layout.exclusions
    .filter((exclusion) => exclusion.verticalOwnership === 'host'
      && (exclusion.wrap === 'square' || exclusion.wrap === 'topAndBottom'
        || exclusion.wrap === 'tight' || exclusion.wrap === 'through'))
    .map((exclusion) => exclusion.anchorOccurrenceId)
    .filter((id): id is string => id !== undefined));
  const effects = new Map((layout.anchorFrames ?? []).flatMap((frame) => (
    frame.status === 'resolved' ? [[frame.occurrenceId, frame.geometry.parentEffectExtent] as const] : []
  )));
  const drawings = new Map(layout.drawings.map((drawing) => [drawing.id, drawing]));
  return layout.lines.map((line) => line.placements.flatMap((placement): HostFloatAnchor[] => {
    if (placement.kind !== 'drawing') return [];
    const drawing = drawings.get(placement.drawingId);
    const anchor = drawing?.anchorLayer;
    if (!drawing || !anchor || anchor.verticalOwnership !== 'host') return [];
    if (!wrapping.has(anchor.occurrenceId)
      && !(anchor.acquisitionOccurrenceId !== undefined && wrapping.has(anchor.acquisitionOccurrenceId))) return [];
    const effect = (anchor.acquisitionOccurrenceId === undefined ? undefined : effects.get(anchor.acquisitionOccurrenceId))
      ?? effects.get(anchor.occurrenceId);
    return [Object.freeze({
      key: anchor.acquisitionOccurrenceId ?? anchor.occurrenceId,
      topPt: drawing.flowBounds.yPt - (effect?.topPt ?? 0),
      bottomPt: drawing.flowBounds.yPt + drawing.flowBounds.heightPt + (effect?.bottomPt ?? 0),
      lineBottomPt: line.bounds.yPt + line.bounds.heightPt,
    })];
  }));
}

/** What WORD_HOST_FLOAT_BOTTOM_MARGIN_DEFERRAL does with the host floats of
 * one acquisition: crossing floats that can rise to end at the body bottom
 * without reaching above their anchor line's bottom are lifted by
 * `lifts` (pt); the first line with a crossing float that cannot rise moves
 * on (`line`). Lifts are given only for lines before that line. */
export type HostFloatBottomMarginPlan = Readonly<{
  lifts: readonly (readonly [key: string, liftPt: number])[];
  line?: Readonly<{
    index: number;
    /** Every crossing float on the line. */
    keys: readonly string[];
    /** Crossing floats on the line that cannot rise. */
    fixedKeys: readonly string[];
    /** Lifts of the line's other crossing floats, used when it stays. */
    lifts: readonly (readonly [key: string, liftPt: number])[];
  }>;
}>;

/** {@link HostFloatBottomMarginPlan} once the paragraph is moved down by
 * `offsetPt`, over its first `lineCount` lines. */
export function hostFloatBottomMarginPlan(
  byLine: readonly (readonly HostFloatAnchor[])[],
  bodyBottomPt: number,
  offsetPt = 0,
  lineCount = byLine.length,
): HostFloatBottomMarginPlan {
  const lifts: (readonly [string, number])[] = [];
  for (let index = 0; index < Math.min(lineCount, byLine.length); index += 1) {
    const crossing = byLine[index]!.filter((float) => float.bottomPt + offsetPt > bodyBottomPt + CROSSING_EPSILON_PT);
    if (crossing.length === 0) continue;
    const rises = (float: HostFloatAnchor): boolean =>
      bodyBottomPt - (float.bottomPt - float.topPt) >= float.lineBottomPt + offsetPt - CROSSING_EPSILON_PT;
    const lineLifts = crossing.filter(rises)
      .map((float) => Object.freeze([float.key, float.bottomPt + offsetPt - bodyBottomPt] as const));
    const fixedKeys = crossing.filter((float) => !rises(float)).map((float) => float.key);
    if (fixedKeys.length > 0) {
      return Object.freeze({
        lifts: Object.freeze(lifts),
        line: Object.freeze({
          index,
          keys: Object.freeze(crossing.map((float) => float.key)),
          fixedKeys: Object.freeze(fixedKeys),
          lifts: Object.freeze(lineLifts),
        }),
      });
    }
    lifts.push(...lineLifts);
  }
  return Object.freeze({ lifts: Object.freeze(lifts) });
}

/** The first line of `layout` that anchors any drawing keyed in `keys`. */
export function anchorLineIndexOf(layout: ParagraphLayout, keys: ReadonlySet<string>): number | undefined {
  const drawings = new Map(layout.drawings.map((drawing) => [drawing.id, drawing]));
  const index = layout.lines.findIndex((line) => line.placements.some((placement) => {
    if (placement.kind !== 'drawing') return false;
    const anchor = drawings.get(placement.drawingId)?.anchorLayer;
    return anchor !== undefined && keys.has(anchor.acquisitionOccurrenceId ?? anchor.occurrenceId);
  }));
  return index < 0 ? undefined : index;
}

/** Geometry noise, far below one twip; not a layout tolerance. */
const CROSSING_EPSILON_PT = 1e-6;
