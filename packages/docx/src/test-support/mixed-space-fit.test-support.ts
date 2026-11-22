import { calcEffectiveFontPx, EAST_ASIAN_RE } from '../layout/text.js';
import { wordCompressedAutoSpaceFloorPt, wordCompressedSpaceFloor } from '../layout/line-compatibility.js';
import { COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION } from '../line-breaker/text-runs.js';
import type { LayoutSeg, LayoutTextSeg } from '../line-breaker/model.js';
import type { PassOperationState } from '../line-breaker/pass-operations.js';
import { sameRouteFaces } from '../line-breaker/font-routes.js';

interface LineTail { tailPx: number; coreLast: string | undefined }
const EMPTY_TAIL: LineTail = { tailPx: 0, coreLast: undefined };
interface LineSpaces {
  valid: boolean; eastAsian: boolean; face: LayoutTextSeg | undefined;
  perSpaceCapacity: number; gaps: { segment: LayoutTextSeg; count: number }[]; count: number;
  autoSpaceCapacity: number;
}
type MixedSpaceMeasure = (segment: LayoutTextSeg, text: string) => number;

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

function autoSpaceRoom(segment: LayoutTextSeg, scale: number): number {
  if (segment.autoSpaceCompressible !== true || segment.autoSpaceBeforePt === undefined) return 0;
  return Math.max(0, (segment.autoSpaceLeadingPx ?? 0)
    - wordCompressedAutoSpaceFloorPt(segment.autoSpaceBeforePt) * scale);
}

function scanLine(line: readonly LayoutSeg[], scale: number): LineSpaces {
  let eastAsian = false;
  const autoSpaceCapacity = line.reduce((sum, item) =>
    sum + ('text' in item && !inkless(item) ? autoSpaceRoom(item, scale) : 0), 0);
  let face: LayoutTextSeg | undefined;
  let capacity = 0;
  let count = 0;
  const gaps: { segment: LayoutTextSeg; count: number }[] = [];
  const invalid = (): LineSpaces => ({
    valid: false, eastAsian, face, perSpaceCapacity: 0, gaps: [], count: 0, autoSpaceCapacity,
  });
  for (const item of line) {
    if (inkless(item)) continue;
    if (!('text' in item)) return invalid();
    if (EAST_ASIAN_RE.test(item.text)) eastAsian = true;
    const spaces = trailingSpaceCount(item.text);
    if (
      item.mixedNaturalTrailingSpacePx !== undefined &&
      spaces > 0 &&
      spaces === item.mixedNaturalTrailingSpaceCount &&
      !item.text.slice(0, item.text.length - spaces).includes(' ')
    ) {
      const itemCapacity = perSpaceCapacity(item, spaces, scale);
      if (face && (!sameSpaceFace(item, face) || Math.abs(itemCapacity - capacity) > 1e-6)) {
        return invalid();
      }
      face ??= item;
      capacity = itemCapacity;
      gaps.push({ segment: item, count: spaces });
      count += spaces;
    } else if (!neutral(item, item.text)) {
      return invalid();
    }
  }
  return { valid: true, eastAsian, face, perSpaceCapacity: capacity, gaps, count, autoSpaceCapacity };
}

function clusterCore(text: string): string {
  let core = text.replace(/ +$/u, '');
  while (core.length > 0 && COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(core.at(-1)!)) {
    core = core.slice(0, -1);
  }
  return core;
}

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

/** Independent full scan checks the incremental summary used by production. */
function assertSummary({ breakerState, scale, strNaturalAdvance }: PassOperationState): void {
  if (!breakerState.mixedSpaceEnabled) return;
  const state = breakerState.mixedSpace;
  const line = breakerState.currentLine;
  const full = scanLine(line, scale);
  const first = state.contributions[state.firstGap];
  const valid = state.invalidItems === 0 && state.mismatchedGaps === 0;
  let tail = EMPTY_TAIL;
  for (const item of line) tail = tailAfter(tail, item, strNaturalAdvance);
  const cached = state.contributions.at(-1)?.tail;
  const same = state.contributions.length === line.length
    && state.contributions.every((c, i) => c.item === line[i])
    && state.visibleItems === line.filter(item => 'text' in item && /\S/u.test(item.text)).length
    && (!cached || (Math.abs(cached.tailPx - tail.tailPx) <= 1e-9 && cached.coreLast === tail.coreLast))
    && valid === full.valid
    && Math.abs(state.autoSpaceCapacity - full.autoSpaceCapacity) <= 1e-9
    && state.autoSpaceInvalidItems === line.filter((item) => !inkless(item) && (!('text' in item)
      || item.verticalRun === true || item.tateChuYoko === true || item.rtl === true
      || item.fitTextRegionIndex !== undefined)).length
    && (!valid || ((state.eastAsianItems > 0) === full.eastAsian
      && first?.item === full.face && state.spaceCount === full.count
      && Math.abs((first?.capacity ?? 0) - full.perSpaceCapacity) <= 1e-9));
  if (!same) throw new Error('WORD_COMPRESSED_SPACE_LINE_FIT line summary drifted from its committed items');
}

let assertions = true;
let work = 0;
const observed = new WeakSet<object>();
function observeReads<T>(array: T[]): T[] {
  if (observed.has(array)) return array;
  const proxy = new Proxy(array, {
    get(target, key, receiver) {
      if (typeof key === 'string' && /^\d+$/.test(key)) work++;
      return Reflect.get(target, key, receiver);
    },
  });
  observed.add(proxy);
  return proxy;
}
export function setMixedSpaceSummaryAssertions(enabled: boolean): void { assertions = enabled; }
export function mixedSpaceSummaryWork(reset = false): number {
  const value = work;
  if (reset) work = 0;
  return value;
}

export function instrumentMixedSpace(mixedSpaceFit: typeof import('../line-breaker/mixed-space-fit.js')) {
  const wrap = <K extends 'performMixedSpaceRequirement' | 'performSettleMixedSpaces'>(name: K) =>
    (...args: Parameters<(typeof mixedSpaceFit)[K]>) => {
      const operation = args[0];
      if (operation.breakerState.mixedSpaceEnabled) {
        operation.breakerState.currentLine = observeReads(operation.breakerState.currentLine);
        operation.breakerState.mixedSpace.contributions = observeReads(operation.breakerState.mixedSpace.contributions);
      }
      if (assertions) assertSummary(operation);
      // The two functions have different arities; Reflect preserves the argument tuple.
      const result = Reflect.apply(mixedSpaceFit[name], undefined, args);
      if (assertions && name === 'performMixedSpaceRequirement') assertSummary(operation);
      return result;
    };
  return { ...mixedSpaceFit,
    performMixedSpaceRequirement: wrap('performMixedSpaceRequirement'),
    performSettleMixedSpaces: wrap('performSettleMixedSpaces'),
  };
}
