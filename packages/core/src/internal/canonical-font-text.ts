import { canonicalCombiningClass } from './canonical-combining-class.js';
import { ASSIGNED_RANGES, COMPOSITIONS, DECOMPOSITIONS, IGNORABLE_RANGES, MARK_RANGES } from './canonical-font-data.js';

function inRanges(ranges: readonly number[], cp: number): boolean {
  let lo = 0, hi = ranges.length / 2 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1, at = mid * 2;
    if (cp < ranges[at]) hi = mid - 1;
    else if (cp > ranges[at + 1]) lo = mid + 1;
    else return true;
  }
  return false;
}
export const canonicalUnicodeAssigned = (cp: number): boolean => inRanges(ASSIGNED_RANGES, cp);
export const canonicalUnicodeMark = (cp: number): boolean => inRanges(MARK_RANGES, cp);
export const canonicalUnicodeIgnorable = (cp: number): boolean => inRanges(IGNORABLE_RANGES, cp);
export function canonicalDecomposition(cp: number): readonly [number, number] | undefined {
  // Unicode Standard §3.12 algorithmic Hangul decomposition, direct LV/T edge.
  if (cp >= 0xac00 && cp <= 0xd7a3) {
    const index = cp - 0xac00, tail = index % 28;
    return tail ? [cp - tail, 0x11a7 + tail] : [0x1100 + Math.floor(index / 588), 0x1161 + Math.floor(index % 588 / 28)];
  }
  let lo = 0, hi = DECOMPOSITIONS.length / 3 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1, at = mid * 3;
    if (cp < DECOMPOSITIONS[at]) hi = mid - 1;
    else if (cp > DECOMPOSITIONS[at]) lo = mid + 1;
    else return [DECOMPOSITIONS[at + 1], DECOMPOSITIONS[at + 2]];
  }
  return undefined;
}
export function canonicalComposition(a: number, b: number): number | undefined {
  if (a >= 0x1100 && a <= 0x1112 && b >= 0x1161 && b <= 0x1175) return 0xac00 + (a - 0x1100) * 588 + (b - 0x1161) * 28;
  if (a >= 0xac00 && a <= 0xd7a3 && (a - 0xac00) % 28 === 0 && b >= 0x11a8 && b <= 0x11c2) return a + b - 0x11a7;
  let lo = 0, hi = COMPOSITIONS.length / 3 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1, at = mid * 3;
    const x = COMPOSITIONS[at], y = COMPOSITIONS[at + 1];
    if (a < x || (a === x && b < y)) hi = mid - 1;
    else if (a > x || (a === x && b > y)) lo = mid + 1;
    else return COMPOSITIONS[at + 2];
  }
  return undefined;
}
/** Stable CCC buckets give O(n + 256) work per disordered sequence, avoiding
 * insertion-sort quadratic work for adversarial long graphemes. Class-zero
 * marks are boundaries, not exceptions to canonical ordering. */
export function canonicalOrder(points: readonly number[]): number[] {
  const result: number[] = [];
  let at = 0;
  while (at < points.length) {
    if (canonicalCombiningClass(points[at]) === 0) result.push(points[at++]);
    const start = at;
    let previous = 0, ordered = true;
    while (at < points.length && canonicalCombiningClass(points[at]) !== 0) {
      const cls = canonicalCombiningClass(points[at++]);
      if (cls < previous) ordered = false;
      previous = cls;
    }
    if (ordered) for (let i = start; i < at; i++) result.push(points[i]);
    else {
      const buckets: number[][] = [];
      for (let i = start; i < at; i++) (buckets[canonicalCombiningClass(points[i])] ??= []).push(points[i]);
      for (const bucket of buckets) if (bucket) for (const cp of bucket) result.push(cp);
    }
  }
  return result;
}
export function canonicalCompose(points: readonly number[], permitted: (cp: number) => boolean = () => true, marksOnly = false): number[] {
  const result: number[] = [];
  let starter = -1, previousClass = 0;
  for (const cp of points) {
    const cls = canonicalCombiningClass(cp);
    const composed = starter >= 0 && (previousClass === 0 || previousClass < cls)
      && (!marksOnly || canonicalUnicodeMark(cp)) ? canonicalComposition(result[starter], cp) : undefined;
    if (composed !== undefined && permitted(composed)) { result[starter] = composed; continue; }
    if (cls === 0) starter = result.length;
    result.push(cp); previousClass = cls;
  }
  return result;
}
/** Pinned NFC display policy shared by support, measurement and painting.
 * Logical OOXML remains unchanged. CSS Fonts §5.4 does not mandate NFC input.
 * Using the same Unicode release for display/decomposition/CCC avoids native
 * normalize()'s otherwise unrecorded Unicode-version boundary. */
export function canonicalFontClusterText(text: string): string {
  const decomposed: number[] = [];
  const visit = (cp: number) => {
    const direct = canonicalDecomposition(cp);
    if (!direct) decomposed.push(cp);
    else { visit(direct[0]); if (direct[1]) visit(direct[1]); }
  };
  for (const ch of text) visit(ch.codePointAt(0) as number);
  return canonicalCompose(canonicalOrder(decomposed)).map((cp) => String.fromCodePoint(cp)).join('');
}
