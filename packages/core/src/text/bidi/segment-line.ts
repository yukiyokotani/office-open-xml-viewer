import { getDefaultBidiEngine } from './engine.js';
import { hasStrongRtl, OBJECT_PLACEHOLDER, buildVisualOrder } from './line-order.js';
import type { BidiClass } from './types.js';

export interface SegmentLineBidiOptions {
  /** DrawingML and Word tabs use Bidi_Class S; ordinary cell objects do not. */
  isTab?: (segment: unknown) => boolean;
  /** Word's run-level rtl mark is an HL1 ambiguous-class override. */
  isRtlMarked?: (segment: unknown) => boolean;
  /** Word's bidi-language digits are classified AN. */
  digitsAsAN?: (segment: unknown) => boolean;
  isRtlAmbiguous?: (character: string) => boolean;
  /** Word paints a whole word slice RTL if any non-space unit is odd-level. */
  directionFromAnyOdd?: boolean;
  /**
   * Direction of the span a segment belongs to, or undefined for a segment
   * outside any span (tabs, objects). A maximal run of consecutive segments
   * whose span direction differs from the base direction is wrapped in a
   * directional isolate (UAX#9 LRI/RLI ... PDI), so it is ordered as one unit
   * and resolved internally in its own direction.
   */
  isolateDirection?: (segment: unknown) => 'ltr' | 'rtl' | undefined;
  /** With directionFromAnyOdd: a block of slices whose closing U+0020
   *  resolves to the other direction is painted in that direction. */
  paintDirectionFromTrailingSpace?: boolean;
}

const LRI = '\u2066';
const RLI = '\u2067';
const PDI = '\u2069';

export interface SegmentLineVisualOrder {
  order: number[];
  rtl: boolean[];
}

const textOf = (segment: unknown): string | undefined => {
  const value = (segment as { text?: unknown }).text;
  return typeof value === 'string' ? value : undefined;
};

/** Shared UAX#9 gate, including Word's run-level rtl mark when requested. */
export function segmentLineHasRtl(
  segments: readonly unknown[], options: SegmentLineBidiOptions = {},
): boolean {
  return segments.some((segment) => options.isRtlMarked?.(segment)
    || hasStrongRtl(textOf(segment) ?? ''));
}

/**
 * UAX#9 L2 ordering of laid-out segments. All three renderers paint complete
 * segments with Canvas; this helper supplies their visual order and per-segment
 * direction. Host hooks express Word's HL1 run overrides and PowerPoint's tab
 * separator without copying the algorithm. The segment-level approximation
 * cannot split a single styled run at an internal embedding-level boundary.
 */
export function computeSegmentLineVisualOrder(
  segments: readonly unknown[], baseRtl: boolean,
  options: SegmentLineBidiOptions = {},
  lineEndNeutral = false,
): SegmentLineVisualOrder {
  const count = segments.length;
  if (count === 0) return { order: [], rtl: [] };
  let full = '';
  const start = new Array<number>(count);
  const end = new Array<number>(count);
  // The UAX#9 engine treats a missing entry as no override, so only authored
  // HL1 positions need storage; the array can remain sparse.
  let override: (BidiClass | undefined)[] | undefined;
  const isolates = options.isolateDirection
    ? segments.map((segment) => {
      const direction = options.isolateDirection!(segment);
      return direction !== undefined && (direction === 'rtl') !== baseRtl ? direction : undefined;
    })
    : undefined;
  for (let i = 0; i < count; i++) {
    const segment = segments[i];
    const value = textOf(segment) ?? '';
    const isolate = isolates?.[i];
    if (isolate !== undefined && isolates![i - 1] !== isolate) {
      full += isolate === 'rtl' ? RLI : LRI;
    }
    const digitsAsAN = options.digitsAsAN?.(segment) === true;
    const rtlMarked = options.isRtlMarked?.(segment) === true;
    start[i] = full.length;
    full += value.length > 0 ? value : OBJECT_PLACEHOLDER;
    end[i] = full.length;
    if (options.isTab?.(segment)) {
      (override ??= [])[start[i]] = 'S';
    } else if (value.length > 0 && (digitsAsAN || rtlMarked)) {
      const classes = override ??= [];
      for (let offset = start[i]; offset < end[i]; offset++) {
        const code = full.charCodeAt(offset);
        if (digitsAsAN && code >= 0x30 && code <= 0x39) {
          classes[offset] = 'AN';
        } else if (rtlMarked && options.isRtlAmbiguous?.(full[offset])) {
          classes[offset] = 'R';
        }
      }
    }
    if (isolate !== undefined && isolates![i + 1] !== isolate) full += PDI;
  }
  // A final neutral keeps the line's trailing whitespace out of UAX #9 L1,
  // for hosts whose line-end spaces stay with their runs.
  if (lineEndNeutral) full += OBJECT_PLACEHOLDER;
  const { levels, paragraphLevel } = getDefaultBidiEngine().computeLevels(
    full, baseRtl ? 'rtl' : 'ltr', override,
  );
  if (!options.directionFromAnyOdd) {
    const { order, segLevels } = buildVisualOrder(levels, paragraphLevel, start);
    return { order, rtl: Array.from(segLevels, (level) => (level & 1) === 1) };
  }

  // Word's whitespace-delimited slices use a content-level paint direction.
  // Exclude trailing spaces, then choose the first anchor with that parity so
  // L2 and the Canvas direction agree for a leading-neutral RTL slice.
  const rtl = new Array<boolean>(count);
  const anchor = new Array<number>(count);
  const spaceStart = new Array<number>(count);
  for (let i = 0; i < count; i++) {
    let scanEnd = end[i];
    while (scanEnd > start[i] && full[scanEnd - 1] === ' ') scanEnd--;
    let firstOdd = -1;
    let firstEven = -1;
    for (let k = start[i]; k < scanEnd; k++) {
      const level = levels[k];
      if (level === 255) continue;
      if ((level & 1) === 1 && firstOdd < 0) firstOdd = k;
      else if ((level & 1) === 0 && firstEven < 0) firstEven = k;
    }
    anchor[i] = firstOdd >= 0 ? firstOdd : firstEven >= 0 ? firstEven : start[i];
    rtl[i] = firstOdd >= 0;
    spaceStart[i] = scanEnd;
  }
  if (options.paintDirectionFromTrailingSpace) {
    // A run of consecutive slices at one ordering level is drawn as a block.
    // When the U+0020 that ends the block resolves to the other direction,
    // UAX #9 puts it outside the block, on that direction's side; painting
    // every slice of the block in that direction moves each slice's own
    // trailing space to the same side, so the gaps fall between the words.
    let runStart = 0;
    for (let i = 1; i <= count; i++) {
      if (i < count && textOf(segments[i]) !== undefined && textOf(segments[i - 1]) !== undefined
        && levels[anchor[i]!] === levels[anchor[runStart]!]) continue;
      const last = i - 1;
      const space = spaceStart[last]!;
      if (textOf(segments[last]) !== undefined && space < end[last]! && levels[space] !== 255) {
        const spaceRtl = (levels[space]! & 1) === 1;
        if (spaceRtl !== ((levels[anchor[last]!]! & 1) === 1)) {
          for (let k = runStart; k <= last; k++) rtl[k] = spaceRtl;
        }
      }
      runStart = i;
    }
  }
  return { order: buildVisualOrder(levels, paragraphLevel, anchor).order, rtl };
}
