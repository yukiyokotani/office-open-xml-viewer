import {
  ExactConvergenceError,
  convergeExactState,
} from './convergence.js';
import { LayoutInvariantError } from './diagnostics.js';

interface LineWrapStateSegment {
  readonly src?: unknown;
  readonly text?: string;
}

export interface LineWrapStateLine {
  readonly physicalLineIndex?: number;
  readonly consumedEnd?: unknown;
  readonly topY?: number;
  readonly xOffset?: number;
  readonly availWidth?: number;
  readonly segments: readonly LineWrapStateSegment[];
}

export class LineWrapNonConvergenceError extends LayoutInvariantError {
  readonly reason: 'cycle' | 'limit';
  readonly states: readonly string[];

  constructor(reason: 'cycle' | 'limit', states: readonly string[]) {
    super(
      'NON_CONVERGENCE',
      reason === 'cycle'
        ? `line wrap measure/resolve cycle did not converge (${states.length} states)`
        : `line wrap measure/resolve pass limit did not converge (${states.length} states)`,
    );
    this.name = 'LineWrapNonConvergenceError';
    this.reason = reason;
    this.states = Object.freeze([...states]);
  }
}

export function cloneSegmentsForLinePass<T extends object>(segments: readonly T[]): T[] {
  return segments.map((segment) => ({ ...segment }));
}

function samePhysicalLine(left: LineWrapStateLine | undefined, right: LineWrapStateLine | undefined): boolean {
  if (!left || !right) return false;
  if (left.physicalLineIndex !== undefined || right.physicalLineIndex !== undefined) {
    return left.physicalLineIndex === right.physicalLineIndex;
  }
  return left.topY !== undefined && left.topY === right.topY;
}

function lineWrapState(
  lines: readonly LineWrapStateLine[],
  probeHeights: readonly number[],
  probeFloors: readonly number[],
): string {
  let physicalIndex = -1;
  return JSON.stringify(lines.map((line, index) => {
    if (!samePhysicalLine(lines[index - 1], line)) physicalIndex += 1;
    return {
      physicalLineIndex: line.physicalLineIndex,
      end: line.consumedEnd, topY: line.topY, xOffset: line.xOffset,
      availableWidth: line.availWidth, probeHeight: probeHeights[physicalIndex],
      probeFloor: probeFloors[physicalIndex],
      segments: line.segments.map(segment => ({ source: segment.src,
        ...(segment.text === undefined ? {} : { text: segment.text }),
      })),
    };
  }));
}

const MAX_LINE_WRAP_PASSES = 16;

/**
 * An adjacent exact-state repeat is the fixed point; a non-adjacent repeat is
 * a real cycle. The pass budget is a fail-closed resource guard for a
 * deterministic orbit whose geometric state cardinality has no useful small
 * bound; exhaustion never accepts stale line geometry. Probes are indexed by
 * physical lines, never gaps. For fixed metrics and one physical line, all G
 * gaps are discovered in one resolving pass and confirmed in the next (three
 * passes including initial measurement), irrespective of G. Variable-height
 * reflow retains the same 16-state fail-closed budget.
 *
 * Two monotone rules make the height feedback a finite lattice walk:
 * - Exclusion probes are monotone per physical line: a line is probed with the
 *   largest allocation it has had (`probeFloors`), while its advance stays the
 *   latest resolved allocation. A taller band only removes gap content, so a
 *   short probe admitting a tall unit and a tall probe excluding it cannot
 *   alternate.
 * - A probe band that leaves out trailing grid leading (`trailingLeading`) is
 *   held at its running maximum per physical line too, never above the floor,
 *   so a leading that grows with a taller advance cannot shrink the band.
 * - If the plain iteration still revisits a state (an advance feedback, e.g. a
 *   paragraph-wide ruby/grid allocation that depends on the partition), it
 *   restarts from the largest advances seen and keeps every physical line's
 *   advance and probe at the running maximum. Lines may then be spaced wider
 *   than their final allocation, never narrower, so no band overlaps.
 * Both maxima range over the finite set of allocations the content can
 * produce, so they rise finitely often; once stable, the next pass repeats.
 * Fixed metrics converge in three passes; otherwise the bound is the number
 * of rises plus two, and the unchanged 16-state budget fails closed past it.
 */
export function convergeLineWrap<TLine extends LineWrapStateLine>(
  measure: (
    probeHeights: readonly number[] | null,
    probeFloors: readonly number[] | null,
    probeLeadings: readonly number[] | null,
  ) => TLine[],
  lineBoxHeight: (line: TLine) => number,
  resolveLineAdvances?: (lines: readonly TLine[]) => readonly number[],
  /** Trailing part of a physical line's advance that its exclusion probe
   * leaves out (WORD_GRID_LINE_CONTENT_FIT_ADMISSION mode-15 band). */
  trailingLeading?: (line: TLine, advance: number) => number,
): TLine[] {
  type Pass = Readonly<{
    lines: TLine[];
    probeHeights: readonly number[];
    probeFloors: readonly number[];
    /** Running maximum of each physical line's leading-reduced probe. */
    probeBands: readonly number[] | null;
    state: string;
  }>;
  const leadings = (lines: readonly TLine[], advances: readonly number[]): readonly number[] | null => {
    if (!trailingLeading) return null;
    let physical = 0;
    const out: number[] = [];
    lines.forEach((line, index) => {
      if (samePhysicalLine(line, lines[index + 1])) return;
      out.push(Math.max(0, trailingLeading(line, advances[physical] ?? 0)));
      physical += 1;
    });
    return Object.freeze(out);
  };
  // The leading-reduced band (probe minus trailing leading) is held at its
  // running maximum per physical line, like the probe floors: a taller advance
  // can raise the leading while the floor stays, and an unheld band could then
  // shrink and readmit content the larger band excluded.
  const bands = (
    floors: readonly number[],
    lineLeadings: readonly number[] | null,
    previous: readonly number[] | null | undefined,
  ): readonly number[] | null => lineLeadings === null ? null
    : Object.freeze(floors.map((floor, index) => Math.min(floor, Math.max(
      floor - (lineLeadings[index] ?? 0),
      previous?.[index] ?? Number.NEGATIVE_INFINITY,
    ))));
  // The probe leadings a pass with `floors` receives so its band is `held`.
  const leadingsFor = (
    floors: readonly number[] | null,
    held: readonly number[] | null | undefined,
  ): readonly number[] | null => !floors || !held ? null
    : Object.freeze(floors.map((floor, index) =>
      held[index] === undefined ? 0 : Math.max(0, floor - held[index]!)));
  const bandState = (held: readonly number[] | null): string =>
    held ? `|${held.join(',')}` : '';
  const resolved = (lines: readonly TLine[]): readonly number[] => {
    const advances = resolveLineAdvances?.(lines);
    return Object.freeze(lines.flatMap((line, index) =>
      samePhysicalLine(line, lines[index + 1])
        ? [] : [advances?.[index] ?? lineBoxHeight(line)]));
  };
  const maximum = (left: readonly number[], right: readonly number[]): readonly number[] =>
    Object.freeze(Array.from({ length: Math.max(left.length, right.length) }, (_, index) =>
      Math.max(left[index] ?? Number.NEGATIVE_INFINITY, right[index] ?? Number.NEGATIVE_INFINITY)));
  // Running maxima of every resolved allocation, used if a cycle appears.
  let seen: readonly number[] = [];
  let seenBands: readonly number[] | null = null;
  let passes = 0;
  try {
    try {
      return convergeExactState<Pass>({
        step: (previous) => {
          passes += 1;
          const lines = measure(previous?.probeHeights ?? null, previous?.probeFloors ?? null,
            leadingsFor(previous?.probeFloors ?? null, previous?.probeBands));
          const probeHeights = resolved(lines);
          seen = maximum(seen, probeHeights);
          const probeFloors = previous ? maximum(probeHeights, previous.probeFloors)
            .slice(0, probeHeights.length) : probeHeights;
          const probeBands = bands(probeFloors, leadings(lines, probeHeights), previous?.probeBands);
          seenBands = probeBands && (seenBands ? maximum(seenBands, probeBands) : probeBands);
          return Object.freeze({
            lines, probeHeights, probeFloors: Object.freeze(probeFloors), probeBands,
            state: lineWrapState(lines, probeHeights, probeFloors) + bandState(probeBands),
          });
        },
        stateOf: (pass) => pass.state,
        limit: MAX_LINE_WRAP_PASSES,
      }).value.lines;
    } catch (error) {
      if (!(error instanceof ExactConvergenceError) || error.reason !== 'cycle'
        || MAX_LINE_WRAP_PASSES - passes < 2) throw error;
    }
    // Monotone advances: inputs only rise, so the walk is finite.
    let held = seen;
    let heldBands: readonly number[] | null = seenBands;
    return convergeExactState<Pass>({
      step: (previous) => {
        const input = previous ? maximum(held, previous.probeHeights) : held;
        held = input;
        const lines = measure(input, input, leadingsFor(input, heldBands));
        const probeHeights = resolved(lines);
        heldBands = bands(input, leadings(lines, probeHeights), heldBands);
        return Object.freeze({
          lines, probeHeights, probeFloors: input, probeBands: heldBands,
          state: lineWrapState(lines, input, input) + bandState(heldBands),
        });
      },
      stateOf: (pass) => pass.state,
      limit: MAX_LINE_WRAP_PASSES - passes,
    }).value.lines;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LineWrapNonConvergenceError(error.reason, error.states);
    }
    throw error;
  }
}
