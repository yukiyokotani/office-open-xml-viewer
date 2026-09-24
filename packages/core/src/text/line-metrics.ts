/**
 * Compatibility exports for the former family-keyed metric table.
 *
 * A family name identifies neither the face selected by Canvas nor the font
 * version used by Office. The old table combined numbers from several font
 * files and Office observations, then applied them to any matching name. That
 * could change line height even when the selected font had different metrics.
 *
 * These signatures were part of the public core API, so they remain callable.
 * They intentionally grant no metric authority to a family name. Format-owned
 * layout may use a metadata-only reference catalog as an explicitly bounded
 * approximation; exact geometry requires selected resource identity and
 * coverage. See issue #1525.
 */

/** @deprecated A family name cannot determine the selected font's line height. */
export function fontWinLineHeightRatio(
  _family: string | null | undefined,
  _eastAsian = false,
): number | null {
  return null;
}

/** @deprecated The family-keyed line-height floor has been removed. */
export function intendedSingleLinePx(
  _family: string | null | undefined,
  _emPx: number,
  _eastAsian = false,
): number {
  return 0;
}

/** @deprecated A family name cannot correct measured Canvas ascent/descent. */
export function correctLineMetrics(
  _family: string | null | undefined,
  _emPx: number,
  ascentPx: number,
  descentPx: number,
  _eastAsian = false,
): { ascent: number; descent: number } {
  return { ascent: ascentPx, descent: descentPx };
}
