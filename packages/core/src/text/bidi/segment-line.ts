export { getDefaultBidiEngine, getBidiEngineLifecycleToken } from './engine.js';
export { unicodePunctuationOrSymbol, bidiClass } from './char-data.js';
import { getDefaultBidiEngine } from './engine.js';
import { hasStrongRtl, OBJECT_PLACEHOLDER, buildVisualOrder } from './line-order.js';
import type { BidiClass } from './types.js';
import { bidiClass, bracket, unicodePunctuationOrSymbol } from './char-data.js';
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

/** Internal acquired-cell view. One resolved anchor per caller-owned source
 * cell, never a public per-character payload. In particular a space-only
 * cell's content rtl flag is false while its actual anchor can be odd. */
export interface SegmentLineResolvedOrder extends SegmentLineVisualOrder {
  anchorLevels: Uint8Array;
}

/** Caller-owned UTF-16 cells partition ORIGINAL parent text. Empty/textless
 * parents retain their single original placeholder; no scope or tab is added. */
export interface SegmentLineSourceCell {
  readonly parentIndex: number;
  readonly startUTF16: number;
  readonly endUTF16: number;
  readonly spaceOnly: boolean;
}
export interface SegmentLineSourceCellOrder extends SegmentLineResolvedOrder {
  readonly view: 'cells' | 'parents';
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
 * Explicit formatting, BN/B/S text classes, brackets/NSM and unknown shapes
 * require full resolution; ordinary strong direction marks keep their UCD class.
 * R/AL-bearing pieces outside opposite spans keep their level-1 odd anchor.
 * Pure L pieces must start L, end L/EN before trailing U+0020, and contain no
 * R/AL/AN: W2/W7 reset at their first L, N1/N2 bound interior NI by L, and their
 * content stays level 2 both at outer base 1 and inside a synthetic LRI span.
 * Every opposite span member must be such an L piece. Thus even a global
 * mixed-slice fallback that removes synthetic isolates preserves these anchors.
 * Space-only slices have rtl=false but can have level 2: with L pieces present,
 * certify their level-1 anchor only at host start, after an R piece, or where
 * L1 resets them before an actual S tab/end. Tabs are barriers, not strong R.
 * Other space contexts require full resolution. L2 runs on the exact short
 * anchor vector; the all-R case retains its direct descending order.
 * Only synchronous acquisition memo misses opt in; no result cache lives here.
 */
export function certifyNativeSegmentLineOrder(
  segments: readonly unknown[], baseRtl: boolean,
  options: SegmentLineBidiOptions,
): SegmentLineVisualOrder | undefined {
  if (!baseRtl || !options.directionFromAnyOdd) return undefined;
  const engine = getDefaultBidiEngine();
  if (!isUnmodifiedNativeBidiEngine(engine)) return undefined;
  // Private tags preserve real tab barriers until space certification. Reuse
  // this backing store as the level vector only after those checks finish.
  const SPACE = 0, RIGHT = 1, LEFT = 2, TAB = 3;
  const kinds = new Uint8Array(segments.length);
  const rtl: boolean[] = [];
  let hasL = false;
  // Only an immutable own reference to this closed, stateless primitive can
  // skip repeated character callbacks. Mutable/accessor/unknown policies keep
  // the full scan; never freeze or register a caller's policy as trusted.
  let reuseAdjacent = false;
  try {
    const descriptor = Object.getOwnPropertyDescriptor(options, 'isRtlAmbiguous');
    reuseAdjacent = descriptor?.value === unicodePunctuationOrSymbol
      && descriptor.writable === false && descriptor.configurable === false;
  } catch { /* Unknown policy proxies keep the existing full scan. */ }
  let previousText: string | undefined;
  let previousRtl = false;
  let previousDigits = false;
  let previousIsolate: 'ltr' | 'rtl' | undefined;
  let previousKind: number | undefined;

  const classifyPiece = (text: string, digitsAsAN: boolean, rtlMarked: boolean): number | undefined => {
    let contentEnd = text.length;
    while (contentEnd > 0 && text[contentEnd - 1] === ' ') contentEnd--;
    let hasR = false;
    let rightEligible = true;
    let leftEligible = true;
    let first: BidiClass | undefined;
    let lastContent: BidiClass | undefined;
    for (let offset = 0; offset < text.length;) {
      const cp = text.codePointAt(offset)!;
      const original = bidiClass(cp);
      // Reject real controls even if a host override might replace their type.
      switch (original) {
        case 'L': case 'R': case 'AL': case 'EN': case 'AN': case 'ES': case 'ET':
        case 'CS': case 'ON': case 'WS': break;
        default: return undefined;
      }
      if (bracket(cp)) return undefined;
      const effective = classOverrideForUnit(text[offset], digitsAsAN, rtlMarked, options) ?? original;
      first ??= effective;
      if (offset < contentEnd) lastContent = effective;
      if (original === 'L' || effective === 'L') rightEligible = false;
      if (effective === 'R' || effective === 'AL' || effective === 'AN') leftEligible = false;
      if (offset < contentEnd && (effective === 'R' || effective === 'AL')) hasR = true;
      offset += cp > 0xffff ? 2 : 1;
    }
    if (contentEnd === 0) return SPACE;
    if (rightEligible && hasR) return RIGHT;
    if (leftEligible && first === 'L' && (lastContent === 'L' || lastContent === 'EN')) return LEFT;
    return undefined;
  };
  for (const segment of segments) {
    if (typeof segment !== 'object' || segment === null) return undefined;
    const isolate = options.isolateDirection?.(segment);
    const text = textOf(segment);
    if (options.isTab?.(segment) === true) {
      if ((isolate !== undefined && isolate !== 'rtl')
        || (text !== undefined && text.length !== 0)) return undefined;
      kinds[rtl.length] = TAB;
      rtl.push(true);
      previousText = undefined;
      previousKind = undefined;
      continue;
    }
    if (text === undefined || text.length === 0) return undefined;
    const digitsAsAN = options.digitsAsAN?.(segment) === true;
    const rtlMarked = options.isRtlMarked?.(segment) === true;
    // Per-segment hooks above remain fresh even on a hit. This constant local
    // workspace reuses only intrinsic kind, never space context or geometry.
    const adjacent = reuseAdjacent && previousKind !== undefined && text === previousText
      && rtlMarked === previousRtl && digitsAsAN === previousDigits && isolate === previousIsolate;
    const kind = adjacent ? previousKind : classifyPiece(text, digitsAsAN, rtlMarked);
    if (kind === undefined || (kind !== LEFT && isolate !== undefined && isolate !== 'rtl')) return undefined;
    previousText = text;
    previousRtl = rtlMarked;
    previousDigits = digitsAsAN;
    previousIsolate = isolate;
    previousKind = kind;
    kinds[rtl.length] = kind;
    rtl.push(kind === RIGHT);
    if (kind === LEFT) hasL = true;
  }
  if (hasL) {
    for (let start = 0; start < kinds.length;) {
      if (kinds[start] !== SPACE) { start++; continue; }
      let end = start + 1;
      while (end < kinds.length && kinds[end] === SPACE) end++;
      if (start !== 0 && kinds[start - 1] !== RIGHT
        && end !== kinds.length && kinds[end] !== TAB) return undefined;
      start = end;
    }
    for (let index = 0; index < kinds.length; index++) {
      if (kinds[index] !== LEFT) kinds[index] = RIGHT;
    }
    return { order: engine.reorderVisual(kinds, 0, kinds.length), rtl };
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
  return computeSegmentLineOrder(segments, baseRtl, options, false);
}

/** Source-local native shaping boundaries from the SAME authored whole-host
 * assembly and UAX #9 resolution used below. Unicode 17 / rev.51 §3.5 requires
 * shaping within resolved level runs. No font, width, occurrence identity or
 * alternate classification is introduced. Callers still own grapheme-safe
 * eligibility and may not treat a SPACE level as a visible ink direction. */
export interface SegmentNativeLevelUnit {
  readonly parentIndex: number;
  readonly startUTF16: number;
  readonly endUTF16: number;
  readonly level: number;
}
export function resolveSegmentNativeLevelUnits(
  segments: readonly unknown[], baseRtl: boolean, options: SegmentLineBidiOptions,
): readonly SegmentNativeLevelUnit[] | undefined {
  if (!isUnmodifiedNativeBidiEngine(getDefaultBidiEngine())) return undefined;
  let units: readonly SegmentNativeLevelUnit[] = [];
  computeSegmentLineOrder(segments, baseRtl, options, false, undefined,
    resolved => { units = resolved; });
  return units;
}

/** Final acquired-cell projection only. Author scopes are assembled from the
 * ORIGINAL parents. A mixed visible island cannot be split by paint, so its
 * whole line returns the exact legacy parent model (#2076). Default format
 * consumers do not use this entry and retain their existing algorithm/API. */
export function computeSegmentLineSourceCellOrder(
  parents: readonly unknown[], cells: readonly SegmentLineSourceCell[],
  baseRtl: boolean, options: SegmentLineBidiOptions,
  legacyOptions: SegmentLineBidiOptions,
): SegmentLineSourceCellOrder {
  const legacy = (): SegmentLineSourceCellOrder => ({
    ...computeSegmentLineOrder(parents, baseRtl, legacyOptions, true) as SegmentLineResolvedOrder,
    view: 'parents',
  });
  // Reusing levels/falling directly to the existing no-isolate pass requires
  // the same classification hooks and either the same authored scope policy
  // with its established uniform-slice fallback, or no legacy isolates.
  if (!options.directionFromAnyOdd
    || options.isTab !== legacyOptions.isTab || options.isRtlMarked !== legacyOptions.isRtlMarked
    || options.digitsAsAN !== legacyOptions.digitsAsAN
    || options.isRtlAmbiguous !== legacyOptions.isRtlAmbiguous
    || options.directionFromAnyOdd !== legacyOptions.directionFromAnyOdd
    || (legacyOptions.isolateDirection !== undefined
      && (legacyOptions.isolateDirection !== options.isolateDirection
        || !legacyOptions.isolateRequiresUniformSlices))) return legacy();
  let ci = 0;
  for (let pi = 0; pi < parents.length; pi++) {
    const text = textOf(parents[pi]) ?? '';
    let end = 0, count = 0;
    while (cells[ci]?.parentIndex === pi) {
      const cell = cells[ci++]!;
      if (!Number.isInteger(cell.startUTF16) || !Number.isInteger(cell.endUTF16)
        || cell.startUTF16 !== end || cell.endUTF16 < end || cell.endUTF16 > text.length
        || (text.length > 0 && cell.endUTF16 === end)
        || cell.spaceOnly !== (text.slice(cell.startUTF16, cell.endUTF16) === ' ')) return legacy();
      end = cell.endUTF16; count++;
    }
    if (count === 0 || end !== text.length || (text.length === 0 && count !== 1)) return legacy();
  }
  if (ci !== cells.length) return legacy();
  return computeSegmentLineOrder(parents, baseRtl, options, true, { cells, legacyOptions }) as SegmentLineSourceCellOrder;
}

function computeSegmentLineOrder(
  segments: readonly unknown[], baseRtl: boolean,
  options: SegmentLineBidiOptions, retainAnchorLevels: boolean,
  projection?: Readonly<{ cells: readonly SegmentLineSourceCell[]; legacyOptions: SegmentLineBidiOptions }>,
  retainNativeUnits?: (units: readonly SegmentNativeLevelUnit[]) => void,
): SegmentLineVisualOrder | SegmentLineResolvedOrder | SegmentLineSourceCellOrder {
  const count = segments.length;
  if (count === 0 && projection) return { view: 'cells', order: [], rtl: [], anchorLevels: new Uint8Array(0) };
  if (count === 0) return retainAnchorLevels
    ? { order: [], rtl: [], anchorLevels: new Uint8Array(0) } : { order: [], rtl: [] };
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
  if (retainNativeUnits) {
    const units: SegmentNativeLevelUnit[] = [];
    for (let parentIndex = 0; parentIndex < count; parentIndex++) {
      const value = textOf(segments[parentIndex]) ?? '';
      if (value.length === 0) continue;
      let first = 0;
      for (let offset = 1; offset <= value.length; offset++) {
        if (offset !== value.length && levels[start[parentIndex]! + offset] === levels[start[parentIndex]! + first]) continue;
        units.push({ parentIndex, startUTF16: first, endUTF16: offset,
          level: levels[start[parentIndex]! + first]! });
        first = offset;
      }
    }
    retainNativeUnits(units);
    return { order: [], rtl: [] };
  }
  if (projection) {
    const cellStart = projection.cells.map(cell => start[cell.parentIndex]! + cell.startUTF16);
    const cellEnd = projection.cells.map(cell =>
      end[cell.parentIndex]! - start[cell.parentIndex]! === 1
        && (textOf(segments[cell.parentIndex]) ?? '').length === 0
        ? end[cell.parentIndex]! : start[cell.parentIndex]! + cell.endUTF16);
    const fine = projectResolvedLine(full, levels, paragraphLevel, cellStart, cellEnd, options,
      projection.cells, true) as SegmentLineResolvedOrder | undefined;
    if (fine) return { ...fine, view: 'cells' };
    if (!hasIsolate) {
      // Identical host/classification: project original parent anchors from
      // the SAME resolved levels, without another UAX pass.
      const coarse = projectResolvedLine(full, levels, paragraphLevel, start, end,
        projection.legacyOptions, false, true) as SegmentLineResolvedOrder;
      return { ...coarse, view: 'parents' };
    }
    // The mixed visible interval belongs to its original parent too. If the
    // legacy model used these isolates, its existing mixed-parent fallback
    // necessarily chooses no isolates; line-level legacy already uses none.
    // Go directly there: at most the established two whole-host resolutions,
    // never an extra authored pass or a fine-paint/coarse-budget mixture.
    const coarse = computeSegmentLineOrder(segments, baseRtl, {
      ...projection.legacyOptions, isolateDirection: undefined,
    }, true) as SegmentLineResolvedOrder;
    return { ...coarse, view: 'parents' };
  }
  const result = projectResolvedLine(full, levels, paragraphLevel, start, end, options,
    Boolean(hasIsolate && options.isolateRequiresUniformSlices), retainAnchorLevels);
  if (!result) {
    // Existing whole-line mixed-slice fallback; the second pass has no isolate
    // hook, so work remains at most two linear UAX resolutions (#2076).
    return computeSegmentLineOrder(segments, baseRtl, {
      ...options, isolateDirection: undefined,
    }, retainAnchorLevels);
  }
  return result;
}

function projectResolvedLine(
  full: string, levels: Uint8Array, paragraphLevel: number,
  start: readonly number[], end: readonly number[], options: SegmentLineBidiOptions,
  rejectMixed: boolean | readonly SegmentLineSourceCell[], retainAnchorLevels: boolean,
): SegmentLineVisualOrder | SegmentLineResolvedOrder | undefined {
  const count = start.length;
  if (!options.directionFromAnyOdd) {
    const { order, segLevels } = buildVisualOrder(levels, paragraphLevel, start);
    const rtl = Array.from(segLevels, (level) => (level & 1) === 1);
    return retainAnchorLevels ? { order, rtl, anchorLevels: segLevels } : { order, rtl };
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
    if ((typeof rejectMixed === 'boolean' ? rejectMixed : !rejectMixed[i]!.spaceOnly)
      && firstOdd >= 0 && firstEven >= 0) return undefined;
    rtl[i] = firstOdd >= 0;
    anchor[i] = firstOdd >= 0 ? firstOdd : firstEven >= 0 ? firstEven : start[i];
  }
  const { order, segLevels } = buildVisualOrder(levels, paragraphLevel, anchor);
  return retainAnchorLevels ? { order, rtl, anchorLevels: segLevels } : { order, rtl };
}
