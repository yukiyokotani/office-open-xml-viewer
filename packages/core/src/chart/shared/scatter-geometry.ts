// Classic chart scatter geometry helpers.
import type { ChartModel } from '../../types/chart';


// ═══════════════════════════════════════════════════════════════════════════
// Scatter chart — X values from series.categories, Y from series.values.
// ═══════════════════════════════════════════════════════════════════════════

// NB: scatter deliberately has NO secondary value axis. Unlike bar/line/area,
// an XY scatter's X axis is already a numeric VALUE axis (not a category axis),
// and Excel/PowerPoint do not define a second Y value axis for a scatter combo
// (`useSecondaryAxis` / a right-hand `<c:valAx>` pairs with a category-based
// family). So `computeSecondaryAxis` is never called here — the CH7 helper is
// wired only into the category-axis families (bar already; line + area now).
export function scatterXValue(cats: string[], index: number, useIndexX: boolean): number | null {
  // A string-backed `<c:xVal>` is plotted by Office as the one-based ordinal
  // sequence 1..N. Zero-based array indices remain an implementation detail.
  if (useIndexX) return index + 1;
  const raw = cats[index];
  if (raw == null) return null;
  const value = parseFloat(raw);
  return Number.isNaN(value) ? null : value;
}


/** Return the linear bubble magnitude prescribed by ST_SizeRepresents.
 * `area` is the schema default, hence sqrt(value); `w` makes radius linear. */
export type BubbleGroupSettings = Pick<
  ChartModel, 'bubbleScale' | 'bubbleSizeRepresents' | 'showNegativeBubbles'
>;


export function bubbleSizeMagnitude(chart: BubbleGroupSettings, value: number): number {
  return chart.bubbleSizeRepresents === 'w' ? value : Math.sqrt(value);
}
