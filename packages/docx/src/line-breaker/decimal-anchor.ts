import { computeLineVisualOrder, type LineOrderModel, type LineOrderResolver } from '../bidi-line.js';
import { resolvedBidiLevels } from '../character-spacing.js';
import type { LayoutSeg } from './model.js';

/** §17.3.1.37 / §17.18.84 select the decimal source address logically, while
 * WORD_RTL_DECIMAL_TAB_PHYSICAL_ALIGNMENT aligns its physical glyph origin.
 * Resolve the full actual/hypothetical line under placement's host model, then
 * filter by logical cell membership. A tab's UAX#9 L1 reset does not isolate
 * implicit/bracket context (X10, N0–N2); never resolve the cell in isolation.
 * Stop assignment still walks logically. Each preceding visual cell member
 * contributes its acquired advance; the selected slice keeps its measured
 * prefix and committed seam.
 *
 * Host HL1 overrides determine slice order/direction, but Canvas receives one
 * text operation and no per-character override. Its existing no-HL1 local
 * levels therefore own the internal glyph position, as in split paint. Only
 * an all-even selected paint slice admits source-order prefix geometry here,
 * INCLUDING spaces/controls: an odd edge space can move the physical start.
 * Mixed/odd local paint keeps prior alignment (#2076); no UTF-16 reversal,
 * source-space budget, new shaping authority or post-justification re-tab.
 * This projection handles a real selected period only; the numeric-end
 * no-separator fallback retains its previous allocation.
 */
export function physicalDecimalAlignmentPoint(input: Readonly<{
  segments: readonly LayoutSeg[];
  point: Readonly<{ segmentIndex: number; charOffset: number }>;
  cellStart: number;
  cellEnd: number;
  advances: readonly number[];
  selectedPrefixAdvance: number;
  model: LineOrderModel;
  resolveOrder?: LineOrderResolver;
}>): number | undefined {
  const { segments, point, advances } = input;
  const selected = segments[point.segmentIndex];
  if (!selected || !('text' in selected) || selected.text[point.charOffset] !== '.') return undefined;
  const visual = (input.resolveOrder ?? computeLineVisualOrder)(segments, true, input.model);
  const direction = visual.rtl[point.segmentIndex] ? 'rtl' : 'ltr';
  const levels = resolvedBidiLevels(selected.text, direction);
  if (levels.some(level => (level & 1) !== 0)) return undefined;
  let physicalOffset = 0;
  for (const index of visual.order) {
    if (index < input.cellStart || index >= input.cellEnd) continue;
    if (index === point.segmentIndex) return physicalOffset + input.selectedPrefixAdvance;
    physicalOffset += advances[index]!;
  }
  return undefined;
}
