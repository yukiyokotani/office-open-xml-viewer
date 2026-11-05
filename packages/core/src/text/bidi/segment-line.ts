export { getDefaultBidiEngine, getBidiEngineLifecycleToken } from './engine.js';
import { getDefaultBidiEngine } from './engine.js';
import { hasStrongRtl, OBJECT_PLACEHOLDER, buildVisualOrder } from './line-order.js';
import type { BidiClass } from './types.js';
import { bidiClass, bracket } from './char-data.js';
import { isUnmodifiedNativeBidiEngine } from './uax9/index.js';

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
  /** Host-authored direction of a text span. Consecutive opposite-direction
   * segments resolve together inside a UAX #9 directional isolate. Objects
   * and tabs return undefined and terminate the span. */
  isolateDirection?: (segment: unknown) => 'ltr' | 'rtl' | undefined;
  /** Preserve legacy ordering when one retained paint slice crosses a resolved
   * odd/even level boundary that the host cannot project into separate paint. */
  isolateRequiresUniformSlices?: boolean;
}

export interface SegmentLineVisualOrder {
  order: number[];
  rtl: boolean[];
}

const textOf = (segment: unknown): string | undefined => {
  const value = (segment as { text?: unknown }).text;
  return typeof value === 'string' ? value : undefined;
};

/** HL1 precedence is shared with certification. `character` is one UTF-16
 * code unit, exactly as canonical override construction supplies it; the native
 * adapter reads the leading unit of each code point, not an astral string. */
function classOverrideForUnit(
  character: string, digitsAsAN: boolean, rtlMarked: boolean,
  options: SegmentLineBidiOptions,
): BidiClass | undefined {
  const code = character.charCodeAt(0);
  if (digitsAsAN && code >= 0x30 && code <= 0x39) return 'AN';
  return rtlMarked && options.isRtlAmbiguous?.(character) ? 'R' : undefined;
}

/** Exact native UAX#9 Unicode 17 (revision 51, X1-X10/W2-W7/N0-N2/I2/L1-L2)
 * execution certificate, not an alternate bidi policy.
 * No embeddings/isolates/overrides or L/NSM are admitted. W2-W7 introduce no L;
 * N0 has no brackets, N1/N2 resolve NI to R, and I2/L1 leave every guaranteed
 * R/AL content anchor at level 1. L2 therefore reverses all retained anchors.
 * ASCII-space-only slices use a level-1 fallback anchor but rtl=false because
 * the canonical content scan trims them completely; tabs have rtl=true.
 * Unknown shapes/classes or synthetic isolates require the full resolver.
 * Only synchronous acquisition memo misses opt in; no result cache lives here.
 */
export function certifyReverseSegmentLine(
  segments: readonly unknown[], baseRtl: boolean,
  options: SegmentLineBidiOptions,
): SegmentLineVisualOrder | undefined {
  if (!baseRtl || !options.directionFromAnyOdd
    || !isUnmodifiedNativeBidiEngine(getDefaultBidiEngine())) return undefined;
  const rtl: boolean[] = [];
  for (const segment of segments) {
    if (typeof segment !== 'object' || segment === null) return undefined;
    const isolate = options.isolateDirection?.(segment);
    if (isolate !== undefined && isolate !== 'rtl') return undefined;
    const text = textOf(segment);
    if (options.isTab?.(segment) === true) {
      if (text !== undefined && text.length !== 0) return undefined;
      rtl.push(true);
      continue;
    }
    if (text === undefined || text.length === 0) return undefined;
    let contentEnd = text.length;
    while (contentEnd > 0 && text[contentEnd - 1] === ' ') contentEnd--;
    const digitsAsAN = options.digitsAsAN?.(segment) === true;
    const rtlMarked = options.isRtlMarked?.(segment) === true;
    let hasR = false;
    for (let offset = 0; offset < text.length;) {
      const cp = text.codePointAt(offset)!;
      const original = bidiClass(cp);
      // Reject real controls even if a host override might replace their type.
      switch (original) {
        case 'R': case 'AL': case 'EN': case 'AN': case 'ES': case 'ET':
        case 'CS': case 'ON': case 'WS': break;
        default: return undefined;
      }
      if (bracket(cp)) return undefined;
      const effective = classOverrideForUnit(text[offset], digitsAsAN, rtlMarked, options) ?? original;
      if (offset < contentEnd && (effective === 'R' || effective === 'AL')) hasR = true;
      offset += cp > 0xffff ? 2 : 1;
    }
    if (contentEnd !== 0 && !hasR) return undefined;
    rtl.push(contentEnd !== 0);
  }
  const order: number[] = [];
  for (let index = segments.length - 1; index >= 0; index--) order.push(index);
  return { order, rtl };
}

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
): SegmentLineVisualOrder {
  const count = segments.length;
  if (count === 0) return { order: [], rtl: [] };
  let full = '';
  const start = new Array<number>(count);
  const end = new Array<number>(count);
  // The UAX#9 engine treats a missing entry as no override, so only authored
  // HL1 positions need storage; the array can remain sparse.
  let override: (BidiClass | undefined)[] | undefined;
  let hasIsolate = false;
  const isolates = options.isolateDirection
    ? segments.map(segment => {
      const direction = options.isolateDirection!(segment);
      if (direction !== undefined && (direction === 'rtl') !== baseRtl) {
        hasIsolate = true;
        return direction;
      }
      return undefined;
    })
    : undefined;
  for (let i = 0; i < count; i++) {
    const segment = segments[i];
    const value = textOf(segment) ?? '';
    const isolate = isolates?.[i];
    if (isolate !== undefined && isolates![i - 1] !== isolate) {
      full += isolate === 'rtl' ? '\u2067' : '\u2066';
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
        const value = classOverrideForUnit(full[offset], digitsAsAN, rtlMarked, options);
        if (value !== undefined) classes[offset] = value;
      }
    }
    if (isolate !== undefined && isolates![i + 1] !== isolate) full += '\u2069';
  }
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
    if (hasIsolate && options.isolateRequiresUniformSlices
      && firstOdd >= 0 && firstEven >= 0) {
      // One anchor cannot preserve both islands of this already-shaped slice.
      // Reuse the previous model for the whole line; splitting paint here would
      // invent geometry. The second pass has no isolate hook, so work is at
      // most two linear UAX resolutions with line-local storage (#2076).
      return computeSegmentLineVisualOrder(segments, baseRtl, {
        ...options, isolateDirection: undefined,
      });
    }
    rtl[i] = firstOdd >= 0;
    anchor[i] = firstOdd >= 0 ? firstOdd : firstEven >= 0 ? firstEven : start[i];
  }
  return { order: buildVisualOrder(levels, paragraphLevel, anchor).order, rtl };
}
