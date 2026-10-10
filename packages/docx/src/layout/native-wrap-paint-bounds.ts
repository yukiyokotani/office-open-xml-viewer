import { drawingMLGeometryBounds, requireResolvedDrawingMLGeometry, type DrawingMLShapePaintPlan } from '@silurus/ooxml-core';
import type { DrawingLayout, LayoutRect } from './types.js';
import { isDeepFrozenPlainDataRoot } from './plain-data.js';
export class NativeWrapPaintBudgetError extends RangeError {}
export interface NativeWrapPaintExtent {
  readonly coordinateSpace: 'section-logical-pt';
  readonly bounds: Readonly<LayoutRect> | null;
  readonly observedOperations: number;
}
function assertSolidPlan(plan: DrawingMLShapePaintPlan): void {
  if (plan.fill && plan.fill.fillType !== 'none' && plan.fill.fillType !== 'solid') {
    throw new Error('Native wrap paint bounds require solid vector paint');
  }
  if (plan.stroke?.fill && plan.stroke.fill.fillType !== 'solid') {
    throw new Error('Native wrap paint bounds require solid vector strokes');
  }
}

/** Conservative library enclosure of the retained shared numeric geometry.
 * Library reading policy uses this enclosure instead of evaluating authored
 * tight/through contours. The source wrap facts stay intact. No painter runs. */
export function deriveNativeWrapPaintExtent(drawing: DrawingLayout, maximumOperations: number): NativeWrapPaintExtent {
  if (!Number.isSafeInteger(maximumOperations) || maximumOperations < 1) throw new NativeWrapPaintBudgetError('Invalid native wrap paint budget');
  if (!isDeepFrozenPlainDataRoot(drawing)) throw new Error('Native wrap drawing must be internally sealed');
  if (drawing.clip || drawing.textBoxIds?.length) throw new Error('Native wrap cannot prove clip or owned textboxes');
  let work = 0;
  for (const command of drawing.commands) {
    if (++work > maximumOperations) throw new NativeWrapPaintBudgetError('Native wrap paint budget exceeded');
    if (command.kind === 'noop') continue;
    if (command.kind !== 'drawingml-shape') throw new Error(`Native wrap cannot prove command ${command.kind}`);
    const units = command.plan.resolvedGeometry?.workUnits;
    if (typeof units !== 'number' || !Number.isSafeInteger(units) || units < 1) throw new Error('Invalid retained geometry work');
    work += units;
    if (work > maximumOperations) throw new NativeWrapPaintBudgetError('Native wrap paint budget exceeded');
    assertSolidPlan(command.plan);
    requireResolvedDrawingMLGeometry(command.plan, 1);
  }
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity;
  for (const command of drawing.commands) {
    if (command.kind === 'noop') continue;
    if (command.kind !== 'drawingml-shape') throw new Error('Unproved native wrap command');
    const box = drawingMLGeometryBounds(command.plan, drawing.transform);
    if (box) { left = Math.min(left, box.x); top = Math.min(top, box.y); right = Math.max(right, box.x + box.w); bottom = Math.max(bottom, box.y + box.h); }
  }
  if (work > maximumOperations) throw new NativeWrapPaintBudgetError('Native wrap paint budget exceeded');
  const bounds = left === Infinity ? null : { xPt: left, yPt: top, widthPt: right - left, heightPt: bottom - top };
  if (bounds && !Object.values(bounds).every(Number.isFinite)) throw new RangeError('Non-finite native wrap enclosure');
  return Object.freeze({ coordinateSpace: 'section-logical-pt', bounds: bounds && Object.freeze(bounds), observedOperations: work });
}
