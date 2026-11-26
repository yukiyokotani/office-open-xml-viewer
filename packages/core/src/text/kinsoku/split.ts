import type { KinsokuRules } from './rules.js';

/**
 * Adjust a CJK line-break position so it does not violate kinsoku.
 *
 * Given a line being split into `head = chars[0..splitAt)` (stays on the
 * current line) and `tail = chars[splitAt..]` (overflows to the next),
 * return the largest legal `splitAt' <= splitAt` such that:
 *   1. `tail'[0]` is not line-START-forbidden (行頭禁則 追い出し — the
 *      offending char and any preceding forbidden chars are pulled down
 *      onto the next line), AND
 *   2. `head'[last]` is not line-END-forbidden (push a dangling opener
 *      to the next line).
 *
 * Retraction is bounded: we never retract below `minSplit` (default 1,
 * so at least one code point always stays on a non-empty line and we
 * keep forward progress). If no legal split exists within that bound
 * (pathological run of forbidden chars), the original `splitAt` is
 * returned unchanged — Word likewise lets an over-long forbidden run
 * overflow rather than loop forever.
 *
 * `chars` must be an array of single code points (e.g. `[...text]`).
 */
export function kinsokuAdjustedSplit(
  chars: string[],
  splitAt: number,
  rules: KinsokuRules,
  minSplit = 1,
): number {
  if (!rules.enabled) return splitAt;
  if (splitAt <= 0 || splitAt >= chars.length) return splitAt;

  const startForbidden = (i: number): boolean =>
    i < chars.length && rules.lineStartForbidden.has(chars[i].codePointAt(0)!);
  const endForbidden = (i: number): boolean =>
    i >= 0 && rules.lineEndForbidden.has(chars[i].codePointAt(0)!);

  let s = splitAt;
  // Retract while the tail begins with a start-forbidden char OR the head
  // ends with an end-forbidden char. Each retraction moves one code point
  // from the head onto the tail (追い出し). Bounded by minSplit.
  while (s > minSplit && (startForbidden(s) || endForbidden(s - 1))) {
    s--;
  }
  // If we hit the floor and it is still illegal, no legal break exists in
  // range — fall back to the unrestricted split (never empty, never hang).
  if (s <= minSplit && (startForbidden(s) || endForbidden(s - 1))) {
    return splitAt;
  }
  return s;
}
