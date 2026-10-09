import type { Worksheet } from '../types.js';

/**
 * Private, projection-owned canonical CSS widths for ACTIVE view-only column
 * resizes.
 *
 * Settled view-only policy: a user drag captures a logical CSS pixel size, so
 * the drag's intent is that pixel size, not the stored-width number written
 * beside it. A later MDW change (retained font rebind) must keep a user-resized
 * column at its dragged CSS px, while authored stored widths keep decoding via
 * the current MDW. Raw `colWidths` values stay as they are (the drag still
 * writes `pxToColWidth`), and this sidecar is never a public Worksheet field.
 * `null` removes an override; 0 means hidden. No DPR/device-unit math lives
 * here. The drag captures integer CSS px; existing per-band display-scale
 * rounding is unchanged, so this is not fractional Office-width support.
 * Worksheets without an override allocate nothing.
 */
const cssWidthsByWorksheet = new WeakMap<Worksheet, Map<number, number>>();

/** The worksheet's CSS overrides, or undefined when it has none. */
export function columnCssWidths(ws: Worksheet): ReadonlyMap<number, number> | undefined {
  return cssWidthsByWorksheet.get(ws);
}

export function getColumnCssWidth(ws: Worksheet, index: number): number | undefined {
  return cssWidthsByWorksheet.get(ws)?.get(index);
}

/** Set (number) or remove (null / invalid) one override; returns whether the
 * projection changed so callers can invalidate geometry and bump revisions. */
export function setColumnCssWidth(ws: Worksheet, index: number, cssPx: number | null): boolean {
  const map = cssWidthsByWorksheet.get(ws);
  if (cssPx === null || !Number.isFinite(cssPx) || cssPx < 0) {
    if (!map?.delete(index)) return false;
    if (map.size === 0) cssWidthsByWorksheet.delete(ws);
    return true;
  }
  if (map?.get(index) === cssPx) return false;
  if (map) map.set(index, cssPx);
  else cssWidthsByWorksheet.set(ws, new Map([[index, cssPx]]));
  return true;
}

/** Give a freshly cloned projection its own copy of the source's overrides,
 * so later edits of either never leak into the other. */
export function inheritColumnCssWidths(source: Worksheet, projection: Worksheet): void {
  const map = cssWidthsByWorksheet.get(source);
  if (map) cssWidthsByWorksheet.set(projection, new Map(map));
  else cssWidthsByWorksheet.delete(projection);
}
