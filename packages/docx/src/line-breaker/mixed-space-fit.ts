/**
 * Line-breaker projection of WORD_COMPRESSED_SPACE_LINE_FIT (issue #1660).
 *
 * The rule applies only to mixed East Asian / Latin lines in compatibility
 * modes below 15 with a compressing characterSpacingControl. Latin-only lines
 * keep WORD_LATIN_INTERWORD_XAVG_FLOOR unchanged: a line enters this
 * projection only once it (or its candidate) holds East Asian text, which that
 * Latin projection never compresses.
 *
 * Every decision is a pure function of the candidate line's content: while a
 * line is built `currentWidth` stays the natural advance of its committed
 * items, and the line facts (shrinkable space count and floor, East Asian
 * presence, unsupported content, visible-text tail) come from a reversible
 * running summary that mirrors `currentLine` item for item. The summary is
 * updated O(1) per commit through the same helpers that mutate the line, and
 * tests assert it equals a full recomputation. The reduction is materialized
 * once, when the line is finalized, onto each space-bearing segment.
 */
import { calcEffectiveFontPx, EAST_ASIAN_RE } from '../layout/text.js';
import {
  wordCompressedAutoSpaceFloorPt,
  wordCompressedSpaceEastAsianOverflowLimit,
  wordCompressedSpaceFloor,
} from '../layout/line-compatibility.js';
import { COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION } from './text-runs.js';
import type { LayoutSeg, LayoutTextSeg } from './model.js';
import type { PassOperationState } from './pass-operations.js';
import { leadingAutoSpacePx } from './advance.js';
import { sameRouteFaces } from './font-routes.js';

/** One committed line item's contribution to the running summary. */
interface Contribution {
  readonly item: LayoutSeg;
  readonly inkless: boolean;
  /** The item is text holding a non-whitespace character. */
  readonly visibleText: boolean;
  readonly eastAsian: boolean;
  /** Content this projection was not observed with (ends it for the line). */
  readonly invalid: boolean;
  /** Shrinkable trailing U+0020 carried by the item (0 when none). */
  readonly gapCount: number;
  readonly capacity: number;
  /** The gap differs in face or floor from the line's first gap. */
  readonly mismatched: boolean;
  /** Shrinkable room of the item's WORD_EAST_ASIAN_AUTO_SPACE gap (0 when none). */
  readonly autoSpaceCapacity: number;
  /** Content with its own advance geometry, which ends automatic-gap
   * compression for the line (see {@link autoSpaceNeutral}). */
  readonly autoSpaceInvalid: boolean;
  /** Terminal-cluster facts of the committed line ending with this item,
   * derived lazily (only East Asian lines ever ask) and then cached. */
  tail?: LineTail;
}

/**
 * The terminal-cluster class, one definition for committed items and
 * candidates: trailing U+0020, then trailing closing punctuation, are
 * excluded from the core whose natural overflow the East Asian limit caps.
 * `tailPx` is the natural advance after the core; `coreLast` the core's last
 * character. A non-text item ends the cluster (no core character).
 */
interface LineTail {
  readonly tailPx: number;
  readonly coreLast: string | undefined;
}

const EMPTY_TAIL: LineTail = { tailPx: 0, coreLast: undefined };

/** Natural advance of a substring of a committed or candidate segment. */
export type MixedSpaceMeasure = (segment: LayoutTextSeg, text: string) => number;

/** Split `text` into its core and terminal cluster (U+0020, then closing marks). */
function clusterCore(text: string): string {
  let core = text.replace(/ +$/u, '');
  while (core.length > 0 && COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(core.at(-1)!)) {
    core = core.slice(0, -1);
  }
  return core;
}


/**
 * Running summary of the committed line, O(1) per commit. It is maintained
 * only through commitMixedLineItem / popMixedLineItem / replaceLastMixedLineItem,
 * which also perform the matching `currentLine` mutation, and is reset with
 * the line. `contributions` mirrors `currentLine` item for item when the
 * paragraph can use the rule (and stays empty otherwise); the totals equal a
 * full recomputation by construction, which tests assert.
 */
export interface MixedSpaceState {
  /** The current line was admitted by shrinking its U+0020; its natural
   * advance may exceed the band until the line is finalized. */
  compressed: boolean;
  contributions: Contribution[];
  eastAsianItems: number;
  invalidItems: number;
  mismatchedGaps: number;
  spaceCount: number;
  /** Index into `contributions` of the line's first gap, or -1. */
  firstGap: number;
  /** Count of contributions with visible text. */
  visibleItems: number;
  /** Total shrinkable room of the committed automatic East Asian gaps. */
  autoSpaceCapacity: number;
  autoSpaceInvalidItems: number;
}

export function createMixedSpaceState(): MixedSpaceState {
  return {
    compressed: false,
    contributions: [],
    eastAsianItems: 0,
    invalidItems: 0,
    mismatchedGaps: 0,
    spaceCount: 0,
    firstGap: -1,
    visibleItems: 0,
    autoSpaceCapacity: 0,
    autoSpaceInvalidItems: 0,
  };
}

type BreakerState = PassOperationState['breakerState'];
type LineItem = BreakerState['currentLine'][number];

function tailAfter(previous: LineTail, item: LayoutSeg, measure: MixedSpaceMeasure): LineTail {
  if (inkless(item)) return previous;
  if (!('text' in item)) return EMPTY_TAIL;
  const core = clusterCore(item.text);
  if (core.length === 0) return { tailPx: previous.tailPx + item.measuredWidth, coreLast: previous.coreLast };
  if (core.length === item.text.length) return { tailPx: 0, coreLast: core.at(-1) };
  // A plain U+0020 tail reuses the advance ordinary fitting already derived.
  if (
    item.mixedNaturalTrailingSpacePx !== undefined &&
    item.text.replace(/ +$/u, '').length === core.length
  ) {
    return { tailPx: item.mixedNaturalTrailingSpacePx, coreLast: core.at(-1) };
  }
  return { tailPx: item.measuredWidth - measure(item, core), coreLast: core.at(-1) };
}

function contributionOf(state: MixedSpaceState, item: LayoutSeg, scale: number): Contribution {
  if (inkless(item)) {
    return {
      item, inkless: true,
      // The visible-content predicate counts every text item, as before.
      visibleText: 'text' in item && /\S/u.test(item.text), eastAsian: false,
      invalid: false, gapCount: 0, capacity: 0, mismatched: false, autoSpaceCapacity: 0,
      autoSpaceInvalid: false,
    };
  }
  if (!('text' in item)) {
    return {
      item, inkless: false, visibleText: false, eastAsian: false,
      invalid: true, gapCount: 0, capacity: 0, mismatched: false, autoSpaceCapacity: 0,
      autoSpaceInvalid: true,
    };
  }
  const eastAsian = EAST_ASIAN_RE.test(item.text);
  const visibleText = /\S/u.test(item.text);
  const spaces = trailingSpaceCount(item.text);
  const autoSpaceCapacity = autoSpaceRoom(item, item.autoSpaceLeadingPx ?? 0, scale);
  const autoSpaceInvalid = !autoSpaceNeutral(item);
  if (
    item.mixedNaturalTrailingSpacePx !== undefined &&
    spaces > 0 &&
    spaces === item.mixedNaturalTrailingSpaceCount &&
    !item.text.slice(0, item.text.length - spaces).includes(' ')
  ) {
    const capacity = perSpaceCapacity(item, spaces, scale);
    const first = state.firstGap >= 0 ? state.contributions[state.firstGap]! : undefined;
    const mismatched = first !== undefined && (
      !sameSpaceFace(item, first.item as LayoutTextSeg) || Math.abs(capacity - first.capacity) > 1e-6
    );
    return {
      item, inkless: false, visibleText, eastAsian, invalid: false,
      gapCount: spaces, capacity, mismatched, autoSpaceCapacity, autoSpaceInvalid,
    };
  }
  return {
    item, inkless: false, visibleText, eastAsian, invalid: !neutral(item, item.text),
    gapCount: 0, capacity: 0, mismatched: false, autoSpaceCapacity, autoSpaceInvalid,
  };
}

function addContribution(state: MixedSpaceState, contribution: Contribution): void {
  if (contribution.gapCount > 0 && state.firstGap < 0) state.firstGap = state.contributions.length;
  state.contributions.push(contribution);
  if (contribution.visibleText) state.visibleItems += 1;
  if (contribution.eastAsian) state.eastAsianItems += 1;
  if (contribution.invalid) state.invalidItems += 1;
  if (contribution.mismatched) state.mismatchedGaps += 1;
  state.spaceCount += contribution.gapCount;
  state.autoSpaceCapacity += contribution.autoSpaceCapacity;
  if (contribution.autoSpaceInvalid) state.autoSpaceInvalidItems += 1;
}

function removeContribution(state: MixedSpaceState): void {
  const contribution = state.contributions.pop();
  if (!contribution) return;
  if (contribution.visibleText) state.visibleItems -= 1;
  if (contribution.eastAsian) state.eastAsianItems -= 1;
  if (contribution.invalid) state.invalidItems -= 1;
  if (contribution.mismatched) state.mismatchedGaps -= 1;
  state.spaceCount -= contribution.gapCount;
  state.autoSpaceCapacity -= contribution.autoSpaceCapacity;
  if (contribution.autoSpaceInvalid) state.autoSpaceInvalidItems -= 1;
  if (state.firstGap === state.contributions.length) state.firstGap = -1;
}

/** Append a committed item to the line and, when the paragraph can use the
 * rule, to its summary. */
export function commitMixedLineItem(
  breakerState: BreakerState,
  item: LineItem,
  scale: number,
): void {
  breakerState.currentLine.push(item);
  if (!breakerState.mixedSpaceEnabled) return;
  addContribution(breakerState.mixedSpace, contributionOf(breakerState.mixedSpace, item, scale));
}

/** Remove the last committed item from the line and its summary. */
export function popMixedLineItem(breakerState: BreakerState): void {
  breakerState.currentLine.pop();
  if (breakerState.mixedSpaceEnabled) removeContribution(breakerState.mixedSpace);
}

/** Replace the last committed item (a kinsoku-retracted head) in both. */
export function replaceLastMixedLineItem(
  breakerState: BreakerState,
  item: LineItem,
  scale: number,
): void {
  popMixedLineItem(breakerState);
  commitMixedLineItem(breakerState, item, scale);
}

/** Terminal-cluster facts of the committed line; amortized O(1): each
 * contribution's tail is derived once from its predecessor and cached. */
function lineTail(state: MixedSpaceState, measure: MixedSpaceMeasure): LineTail {
  const contributions = state.contributions;
  let start = contributions.length - 1;
  while (start >= 0 && contributions[start]!.tail === undefined) start -= 1;
  let tail = start >= 0 ? contributions[start]!.tail! : EMPTY_TAIL;
  for (let index = start + 1; index < contributions.length; index += 1) {
    tail = tailAfter(tail, contributions[index]!.item, measure);
    contributions[index]!.tail = tail;
  }
  return tail;
}

/** The committed line holds visible (non-whitespace) text; O(1). Only asked
 * for segments that carry the rule's eligibility. */
export function mixedLineHasVisibleText(breakerState: BreakerState): boolean {
  return breakerState.mixedSpace.visibleItems > 0;
}

/** The current line could shrink spaces for some candidate: O(1) scope gate
 * evaluated before any candidate text is built or measured. */
export function mixedLineMayShrink(breakerState: BreakerState): boolean {
  const state = breakerState.mixedSpace;
  return breakerState.mixedSpaceEnabled && (
    (state.spaceCount > 0 && state.invalidItems === 0 && state.mismatchedGaps === 0)
    || (state.autoSpaceCapacity > 0 && state.autoSpaceInvalidItems === 0));
}

/** Scope gate for a candidate holding `text`: the line can shrink — through
 * its committed gaps, or through the automatic East Asian gap that the
 * candidate `segment` brings after the line's text — and the candidate or the
 * committed line holds East Asian text. Latin-only lines stop here before any
 * candidate is built or measured. */
export function mixedCandidateMayShrink(
  breakerState: BreakerState,
  text: string,
  segment?: LayoutTextSeg,
): boolean {
  const state = breakerState.mixedSpace;
  const last = breakerState.currentLine.at(-1);
  const candidateGap = segment?.autoSpaceCompressible === true && breakerState.mixedSpaceEnabled
    && state.autoSpaceInvalidItems === 0 && last !== undefined && 'text' in last;
  return (mixedLineMayShrink(breakerState) || candidateGap)
    && (state.eastAsianItems > 0 || EAST_ASIAN_RE.test(text));
}

/** U+0020 compression was not observed on a linesAndChars grid and keeps
 * natural spaces there; WORD_EAST_ASIAN_AUTO_SPACE gaps shrink on it too. */
function spacesShrinkOnGrid(gridType: string | null | undefined): boolean {
  return gridType !== 'linesAndChars';
}

/** Room a WORD_EAST_ASIAN_AUTO_SPACE gap of `appliedPx` can give up on a line
 * of WORD_COMPRESSED_SPACE_LINE_FIT: down to its compressed floor. */
function autoSpaceRoom(segment: LayoutTextSeg, appliedPx: number, scale: number): number {
  if (segment.autoSpaceCompressible !== true || segment.autoSpaceBeforePt === undefined) return 0;
  return Math.max(0, appliedPx - wordCompressedAutoSpaceFloorPt(segment.autoSpaceBeforePt) * scale);
}

/** Line facts from the running summary; O(1). */
function summaryLine(state: MixedSpaceState): {
  valid: boolean;
  eastAsian: boolean;
  face: LayoutTextSeg | undefined;
  perSpaceCapacity: number;
  count: number;
} {
  const first = state.firstGap >= 0 ? state.contributions[state.firstGap]! : undefined;
  const valid = state.invalidItems === 0 && state.mismatchedGaps === 0;
  return {
    valid,
    eastAsian: state.eastAsianItems > 0,
    face: valid ? first?.item as LayoutTextSeg | undefined : undefined,
    perSpaceCapacity: valid ? first?.capacity ?? 0 : 0,
    count: valid ? state.spaceCount : 0,
  };
}

/** One candidate unit: consecutive pieces of text that must share a line
 * (a segment, a prefix of one, or a joined group across source runs). */
export interface MixedSpaceCandidate {
  readonly pieces: readonly Readonly<{ segment: LayoutTextSeg; text: string }>[];
  /** Fit width of the whole unit, as used by the ordinary natural fit test. */
  readonly fitWidth: number;
}

function sameSpaceFace(candidate: LayoutTextSeg, reference: LayoutTextSeg): boolean {
  return (
    candidate.mixedSpaceAverageWidthRatio === reference.mixedSpaceAverageWidthRatio &&
    sameRouteFaces(candidate.fontRoute, reference.fontRoute) &&
    candidate.fontFamily === reference.fontFamily &&
    candidate.fontSize === reference.fontSize &&
    candidate.bold === reference.bold &&
    candidate.italic === reference.italic
  );
}

/** Content that keeps its natural advance beside shrinking automatic East
 * Asian gaps: anything but vertical, tate-chu-yoko, right-to-left or fitText
 * geometry. Authored spaces keep their advance, and the
 * balanceSingleByteDoubleByteWidth projection does not touch the gap. */
function autoSpaceNeutral(segment: LayoutTextSeg): boolean {
  return !segment.verticalRun && !segment.tateChuYoko && !segment.rtl
    && segment.fitTextRegionIndex === undefined;
}

/** Text without U+0020 that keeps its natural advance on a mixed line. */
function neutral(segment: LayoutTextSeg, text: string): boolean {
  return (
    !text.includes(' ') &&
    !segment.verticalRun &&
    !segment.tateChuYoko &&
    !segment.rtl &&
    segment.fitTextRegionIndex === undefined &&
    segment.widthBalanceGridDeltaFactor === undefined
  );
}

function inkless(segment: LayoutSeg): boolean {
  return ('text' in segment && segment.metricOnly === true)
    || ('imagePath' in segment && Boolean(segment.anchor));
}

function trailingSpaceCount(text: string): number {
  return text.length - text.replace(/ +$/u, '').length;
}

function perSpaceCapacity(segment: LayoutTextSeg, count: number, scale: number): number {
  return Math.max(
    0,
    segment.mixedNaturalTrailingSpacePx! / count -
      wordCompressedSpaceFloor(calcEffectiveFontPx(segment, scale), segment.mixedSpaceAverageWidthRatio!),
  );
}

/** Total U+0020 reduction the line needs to append `candidate`, or undefined
 * when the candidate belongs on the next line. Pure; O(1) in the committed
 * line plus the candidate's own text. */
export function performMixedSpaceRequirement(
  operationState: PassOperationState,
  candidate: MixedSpaceCandidate,
): number | undefined {
  const {
    breakerState,
    availW,
    fitsMeasuredWidth,
    characterGrid,
    baseRtl,
    widthPolicy,
    strNaturalAdvance,
    scale,
  } = operationState;
  const { pieces } = candidate;
  if (
    pieces.length === 0 ||
    !mixedCandidateMayShrink(
      breakerState, pieces.map((piece) => piece.text).join(''), pieces[0]!.segment,
    ) ||
    baseRtl ||
    widthPolicy !== 'bounded' ||
    characterGrid?.type === 'snapToChars' ||
    breakerState.currentLine.length === 0
  )
    return undefined;
  const line = summaryLine(breakerState.mixedSpace);
  // Only mixed East Asian / Latin lines; Latin-only lines keep their own rule.
  if (!line.eastAsian && !pieces.some((piece) => EAST_ASIAN_RE.test(piece.text))) return undefined;
  // U+0020 shrink only on a line and candidate inside their observed scope;
  // automatic East Asian gaps only beside content of natural advance.
  let spacesValid = line.valid && spacesShrinkOnGrid(characterGrid?.type);
  let autoSpaceValid = breakerState.mixedSpace.autoSpaceInvalidItems === 0;
  let candidateAutoSpaceCapacity = 0;
  for (const [index, piece] of pieces.entries()) {
    const spaces = trailingSpaceCount(piece.text);
    const visible = piece.text.slice(0, piece.text.length - spaces);
    if (!piece.segment.fontRoute) return undefined;
    if (!autoSpaceNeutral(piece.segment)) autoSpaceValid = false;
    if (!neutral(piece.segment, visible)) spacesValid = false;
    // Trailing spaces become gaps of the line; they must share its face.
    if (
      spaces > 0 &&
      (index !== pieces.length - 1 ||
        piece.segment.mixedSpaceAverageWidthRatio === undefined ||
        (line.face !== undefined && !sameSpaceFace(piece.segment, line.face)))
    ) {
      spacesValid = false;
    }
    // Every piece starts its segment and follows text on this line, so it
    // brings its same-line automatic gap.
    candidateAutoSpaceCapacity += autoSpaceRoom(
      piece.segment, leadingAutoSpacePx(operationState, piece.segment), scale,
    );
  }
  const spaceCapacity = spacesValid && line.face ? line.perSpaceCapacity * line.count : 0;
  const autoSpaceCapacity = autoSpaceValid
    ? breakerState.mixedSpace.autoSpaceCapacity + candidateAutoSpaceCapacity : 0;
  const capacity = spaceCapacity + autoSpaceCapacity;
  if (!(capacity > 0)) return undefined;
  const naturalWidth = breakerState.currentWidth;
  const fitWidth = candidate.fitWidth;
  const required = Math.max(0, naturalWidth + fitWidth - availW());
  if (required > capacity || !fitsMeasuredWidth(naturalWidth + fitWidth - required, availW())) {
    return undefined;
  }
  // An East Asian final core character may overflow by at most half its font
  // size, measured naturally up to the end of the line-plus-candidate core:
  // the terminal cluster (U+0020, then closing punctuation, which keeps its
  // own line-end compression) is excluded whether it was committed earlier or
  // arrives with the candidate, so a source-run seam cannot change it.
  const cores = pieces.map((piece) => piece.text.replace(/ +$/u, ''));
  let last = cores.length - 1;
  while (last >= 0) {
    cores[last] = clusterCore(cores[last]!);
    if (cores[last]!.length > 0) break;
    last -= 1;
  }
  const tail = last >= 0 ? EMPTY_TAIL : lineTail(breakerState.mixedSpace, strNaturalAdvance);
  const lastCharacter = last >= 0 ? cores[last]!.at(-1) : tail.coreLast;
  const limitSegment = last >= 0 ? pieces[last]!.segment : pieces[0]!.segment;
  const limit = wordCompressedSpaceEastAsianOverflowLimit(
    lastCharacter !== undefined && EAST_ASIAN_RE.test(lastCharacter),
    calcEffectiveFontPx(limitSegment, scale),
  );
  if (limit !== undefined) {
    let coreEnd = naturalWidth - (last >= 0 ? 0 : tail.tailPx);
    for (let index = 0; index <= last; index += 1) {
      if (cores[index]!.length > 0) {
        coreEnd += strNaturalAdvance(pieces[index]!.segment, cores[index]!);
      }
    }
    if (coreEnd - availW() > limit + 1e-9) return undefined;
  }
  return required;
}

/** Record that the line was admitted by shrinking its spaces. */
export function performMarkMixedSpacesCompressed(operationState: PassOperationState): void {
  operationState.breakerState.mixedSpace.compressed = true;
}

/** Finalize the line: shrink its spaces by exactly what its committed natural
 * advance still exceeds the band (a kinsoku retraction may have shortened it),
 * write each space-bearing segment once, and reset for the next line. */
export function performSettleMixedSpaces(operationState: PassOperationState): void {
  const { breakerState, availW, characterGrid } = operationState;
  if (breakerState.mixedSpace.compressed) {
    const line = summaryLine(breakerState.mixedSpace);
    // Finalization walks the line once; every gap is written exactly once.
    const gaps = breakerState.mixedSpace.contributions.flatMap((contribution) => {
      return contribution.gapCount > 0
        ? [{ segment: contribution.item as LayoutTextSeg, count: contribution.gapCount }]
        : [];
    });
    // Spaces ending the line are not gaps: like every fit decision, the line
    // end is judged without them (they keep their natural advance).
    let lineEndSpacePx = 0;
    for (let index = breakerState.currentLine.length - 1; index >= 0; index -= 1) {
      const item = breakerState.currentLine[index]!;
      if (inkless(item)) continue;
      const gap = gaps.at(-1);
      if (gap?.segment !== item) break;
      gaps.pop();
      lineEndSpacePx += gap.segment.mixedNaturalTrailingSpacePx!;
      if (gap.segment.text.trim().length > 0) break;
    }
    const count = gaps.reduce((sum, gap) => sum + gap.count, 0);
    const needed = Math.max(0, breakerState.currentWidth - lineEndSpacePx - availW());
    const spaceCapacity = line.valid && count > 0 && spacesShrinkOnGrid(characterGrid?.type)
      ? line.perSpaceCapacity * count : 0;
    const autoSpaceGaps = breakerState.mixedSpace.autoSpaceInvalidItems === 0
      ? breakerState.mixedSpace.contributions.filter((contribution) => contribution.autoSpaceCapacity > 0)
      : [];
    const autoSpaceCapacity = autoSpaceGaps.reduce((sum, gap) => sum + gap.autoSpaceCapacity, 0);
    const reduction = Math.min(needed, spaceCapacity + autoSpaceCapacity);
    if (reduction > 0) {
      // Every gap gives up the same share of its own room; with U+0020 alone
      // that is the same amount per space.
      const spaceReduction = autoSpaceCapacity > 0
        ? reduction * spaceCapacity / (spaceCapacity + autoSpaceCapacity)
        : reduction;
      if (spaceReduction > 0) {
        const perSpace = spaceReduction / count;
        for (const gap of gaps) {
          const gapReduction = perSpace * gap.count;
          gap.segment.measuredWidth -= gapReduction;
          gap.segment.latinSpaceCompressionPx = gapReduction;
        }
      }
      if (autoSpaceCapacity > 0) {
        const share = (reduction - spaceReduction) / autoSpaceCapacity;
        for (const gap of autoSpaceGaps) {
          const segment = gap.item as LayoutTextSeg;
          const gapReduction = share * gap.autoSpaceCapacity;
          segment.measuredWidth -= gapReduction;
          segment.autoSpaceLeadingPx = (segment.autoSpaceLeadingPx ?? 0) - gapReduction;
        }
      }
      breakerState.currentWidth -= reduction;
    }
  }
  breakerState.mixedSpace = createMixedSpaceState();
}
