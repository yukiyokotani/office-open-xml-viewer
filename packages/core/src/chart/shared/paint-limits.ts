// Classic chart paint limits helpers.
import { MAX_CHART_PAINT_COMPONENTS, MAX_CHART_PAINT_RECIPE_COMPONENTS } from '../resource-limits.js';


// The parser may preserve up to the OOXML cache ceiling, but expanding every
// point into several synchronous Canvas calls can monopolize the UI thread.
// Refuse an oversized paint atomically instead of drawing a misleading prefix.
// This is an availability boundary, not an automatic chart-layout heuristic.

// Marker gradients are resolved for each painted marker. Bound both one
// recipe and the chart-wide stop registrations so a valid public model cannot
// turn a bounded point count into unbounded synchronous Canvas work.
export const MAX_CANVAS_MARKER_GRADIENT_STOPS = MAX_CHART_PAINT_RECIPE_COMPONENTS;

export const MAX_CANVAS_MARKER_PAINT_COMPONENTS = MAX_CHART_PAINT_COMPONENTS;

export const MAX_CANVAS_LABEL_GRADIENT_STOPS = MAX_CANVAS_MARKER_GRADIENT_STOPS;

export const MAX_CANVAS_LABEL_PAINT_COMPONENTS = MAX_CANVAS_MARKER_PAINT_COMPONENTS;


export const CLASSIC_THREE_D_FAMILIES = new Set([
  'pie',
  'line', 'stackedLine', 'stackedLinePct',
  'area', 'stackedArea', 'stackedAreaPct',
  'clusteredBar', 'clusteredBarH',
  'stackedBar', 'stackedBarH', 'stackedBarPct', 'stackedBarHPct',
]);
