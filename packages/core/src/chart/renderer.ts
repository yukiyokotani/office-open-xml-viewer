// Unified chart renderer. Dispatches on canonical `ChartModel.chartType` and
// delegates to per-family implementations (bar, line, area, pie, radar,
// scatter, waterfall). Ported from the xlsx implementation with pptx
// extensions (valMin-aware axis, plotAreaBg, dataPointColors, waterfall).

import type { ChartModel, ChartRect } from '../types/chart';

import { classicDataMarkPaintWorkCount } from './classic-paint-work.js';
import { paintChartImageFill, type ChartImageLookup, withChartImageLookup } from './image-fill.js';

import { paintChartStyleEffects, withChartEffectBudget } from './style-effects.js';

import { strokeChartFrameRect } from './compound-frame.js';

export { chartVariesColorsByPoint } from './legend-entry-plan.js';

import {
  classicCanvasPointCount,
  chartEffectConsumerUpperBound,
  MAX_CANVAS_CHART_POINTS,
  sourceChartStructureCount,
} from './resource-limits.js';
import { classicPlotDispatch } from './plot-groups.js';
import { type ChartThreeDRenderer } from './three-d-contract.js';
import type { ChartRegionMapRenderer } from './region-map-contract.js';
import type { ChartExRenderer } from './chart-ex-contract.js';

export { resolveChartExLabel } from './chart-ex-label.js';
import { withChartStyleIndexCache, withEffectiveChartStyleRoles } from './effective-style.js';
import { withSparseStyleIndexCache } from './sparse-style-index.js';

import { applyPlotVisibleOnly } from './source-visibility.js';

import { resolveFill } from '../shape/paint.js';

import { EMU_PER_PT, PT_TO_PX } from '../units.js';

import { drawChartDisplayUnitLabels } from './shared/axis.js';
import { applyLinkedChartStyleRoles } from './shared/style-roles.js';
import { dashPatternForLine } from './shared/geometry.js';
import { MAX_CANVAS_MARKER_PAINT_COMPONENTS, MAX_CANVAS_LABEL_PAINT_COMPONENTS } from './shared/paint-limits.js';
import { classicMarkerPaintWorkCount, classicThreeDWorkCount, rejectOversizedCanvasChart } from './shared/resource.js';
import { chartLabelPaintWorkCount } from './shared/data-labels.js';
import { drawChartTextBoxes, CHART_SPACE_CORNER_RADIUS_PT, chartSpaceRoundedPath } from './shared/frame.js';
import { renderBarChart } from './families/bar.js';
import { renderLineChart } from './families/line.js';
import { renderStockChart } from './families/stock.js';
import { renderSurfaceChart } from './families/surface.js';
import { renderAreaChart } from './families/area.js';
import { renderPieChart, renderOfPieChart } from './families/pie.js';
import { renderRadarChart } from './families/radar.js';
import { renderScatterChart } from './families/scatter.js';

export { CHART_PALETTE, chartColor, indexPointOverrides, chartExSeriesFormatIndex } from './shared/palette.js';
export { chartFontFamily, chartFontCss } from './shared/fonts.js';
export { legendEntryColor } from './shared/legend.js';
export { drawAxisTitles } from './shared/axis.js';
export { measuredLegendReserve, drawLegendForLayout } from './shared/legend.js';
export { drawAxisTick, strokeAxisSegment, strokeValueGridlineH, valGridStroke, valMinorGridStroke, drawValMajorGridlines, formatPrimaryValueAxisTick, planValueAxis, axisLabelPx, wrapMeasuredText } from './shared/axis.js';
export { measuredCartesianTitleBand, drawChartTitleForLayout } from './shared/title.js';
export { drawMarker } from './shared/markers.js';
export { richDataLabelOptions, drawBoundedDataLabelText, chartLabelPaintWorkCount } from './shared/data-labels.js';
export { chartExStyleColor, chartExDataPointFill, chartExMarkerPaint, chartExDataPointPaint, chartExFillStyle, paintClassicDataPointPath, paintClassicDataPointRect, resolveChartExSeriesLineStyle, applyResolvedChartExLineStyle, applyChartExSeriesLineStyle, chartExLegendSeries } from './shared/chartex-style.js';
export { classicMarkerPaintWorkCount, rejectOversizedCanvasChart } from './shared/resource.js';
export type { ChartExStyle, ResolvedChartExLineStyle } from './shared/chartex-style.js';
export { chartExValueTickLabelOffsetPx, renderBarChart } from './families/bar.js';
export { renderLineChart } from './families/line.js';

type ClassicFamilyRenderer = (
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
) => void;

const barFamily: ClassicFamilyRenderer = (ctx, chart, rect, ptToPx, rotation) =>
  renderBarChart(ctx, chart, rect, ptToPx, {}, rotation);
const pieFamily: ClassicFamilyRenderer = (ctx, chart, rect, ptToPx, rotation) =>
  renderPieChart(ctx, chart, rect, false, ptToPx, rotation);
const doughnutFamily: ClassicFamilyRenderer = (ctx, chart, rect, ptToPx, rotation) =>
  renderPieChart(ctx, chart, rect, true, ptToPx, rotation);

const groupedFamilyRenderers: Record<string, ClassicFamilyRenderer> = {
  'bar-combo': barFamily,
  'line-groups': renderLineChart,
  'area-groups': renderAreaChart,
  'scatter-bubble': renderScatterChart,
  'stock-line': renderStockChart,
};

const classicFamilyRenderers: Record<string, ClassicFamilyRenderer> = {
  clusteredBar: barFamily,
  clusteredBarH: barFamily,
  stackedBar: barFamily,
  stackedBarH: barFamily,
  stackedBarPct: barFamily,
  stackedBarHPct: barFamily,
  line: renderLineChart,
  stackedLine: renderLineChart,
  stackedLinePct: renderLineChart,
  area: renderAreaChart,
  stackedArea: renderAreaChart,
  stackedAreaPct: renderAreaChart,
  pie: pieFamily,
  ofPie: renderOfPieChart,
  doughnut: doughnutFamily,
  radar: renderRadarChart,
  scatter: renderScatterChart,
  bubble: renderScatterChart,
  stock: renderStockChart,
  // Surface also paints without an injected 3-D renderer. Keep that classic path.
  surface: renderSurfaceChart,
  surface3D: renderSurfaceChart,
};

/**
 * Render a chart (background frame + dispatch on `chartType`).
 * `rect` is in pixel coordinates on the target canvas.
 */
function renderChartImpl(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  /**
   * Pixels per point at the caller's current display scale. For PPTX at
   * 960px/12192000EMU the value is ~1.05; xlsx's sheet view renders at
   * device-px where 1pt≈1.333. Used to size title/axis labels whose
   * XML-specified sizes are in OOXML hundredths of a point.
   */
  ptToPx: number = PT_TO_PX,
  /**
   * Rotation already applied by the host frame transform. DrawingML gradient
   * fills with `rotWithShape="0"` counter-rotate by this amount.
   */
  shapeRotationDeg = 0,
  /** Optional 3-D renderer. Without it, the canonical 2-D family remains
   * visible and no mesh/camera implementation enters the static render path. */
  threeD?: ChartThreeDRenderer,
  /** Optional offline Region Map renderer. */
  regionMap?: ChartRegionMapRenderer,
  /** Host-warmed image cache used by picture-fill availability preflight. */
  imageLookup?: ChartImageLookup,
  /** Optional Microsoft ChartEx family renderer. */
  chartEx?: ChartExRenderer,
  /** Raw, pre-projection source size checked before any visibility/style clone. */
  sourceStructureCount = sourceChartStructureCount(chart),
): void {
  // The per-family renderers (and the early-return/default text paths below)
  // mutate shared canvas state — textAlign, textBaseline, font, fillStyle,
  // etc. — without restoring it. Callers (docx/pptx draw chart shapes inline
  // with surrounding text; xlsx happens to wrap the call in its own
  // save/clip/restore) must not observe those mutations afterward. Wrapping
  // the whole body in a single save/restore here fixes it once for every
  // caller instead of requiring each call site to remember to do so.
  ctx.save();
  try {
    // Refuse oversized caller-supplied classic models before visibility/style
    // projections allocate replacement series, override, or trendline arrays.
    // Parsed packages are already bounded, but the public ChartModel contract
    // can also be constructed directly by an application.
    if (rejectOversizedCanvasChart(ctx, rect, sourceStructureCount)) return;
    chart = applyLinkedChartStyleRoles(chart, ptToPx);
    const { x, y, w, h } = rect;
    const rounded = chart.roundedCorners === true;
    const cornerRadius = rounded ? CHART_SPACE_CORNER_RADIUS_PT * ptToPx : 0;
    const fillChartSpace = (target: CanvasRenderingContext2D) => {
      if (rounded) {
        chartSpaceRoundedPath(target, x, y, w, h, cornerRadius);
        target.fill();
      } else {
        target.fillRect(x, y, w, h);
      }
    };
    // Only fill the outer chartSpace when chartBg is set; a null means noFill
    // (transparent) per OOXML, so the underlying slide/sheet shows through.
    // `roundedCorners` rounds that frame; it does not clip labels or other chart
    // content. Office-produced PPTX output keeps axis labels outside the rounded
    // frame silhouette visible. Picture fill alone needs a temporary local clip.
    paintChartStyleEffects(
      ctx,
      chart.chartAreaStyle,
      chart.chartStyleRoles?.chartArea,
      0,
      rect,
      ptToPx,
      target => {
        if (chart.chartFillHidden === true) {
          // Direct or linked `noFill`: retain the host surface beneath the chart.
        } else if (chart.chartFill?.fillType === 'image') {
          if (rounded) {
            target.save();
            chartSpaceRoundedPath(target, x, y, w, h, cornerRadius);
            target.clip();
            paintChartImageFill(
              target, chart.chartFill, x, y, w, h, ptToPx, shapeRotationDeg,
            );
            target.restore();
          } else {
            paintChartImageFill(
              target, chart.chartFill, x, y, w, h, ptToPx, shapeRotationDeg,
            );
          }
        } else if (chart.chartFill) {
          const fill = resolveFill(chart.chartFill, target, x, y, w, h, shapeRotationDeg);
          if (fill) target.fillStyle = fill;
          if (fill) fillChartSpace(target);
        } else if (chart.chartBg) {
          target.fillStyle = `#${chart.chartBg}`;
          fillChartSpace(target);
        }

        // Explicit chart border — drawn only when DrawingML declares a paintable
        // line. Width comes from
        // `<a:ln@w>` (EMU → pt → px); absent width falls back to a 1px hairline.
        if (chart.chartBorderHidden !== true
          && (chart.chartBorderLineFill || chart.chartBorderColor)) {
          target.save();
          const stroke = chart.chartBorderLineFill
            ? resolveFill(chart.chartBorderLineFill, target, x, y, w, h, shapeRotationDeg)
            : chart.chartBorderColor ? `#${chart.chartBorderColor}` : null;
          if (!stroke) {
            target.restore();
          } else {
            target.strokeStyle = stroke;
            // `<a:ln>` with no `@w` means width 0 per ECMA-376 §20.1.2.2.24, i.e. invisible;
            // but Excel renders a fill-without-width line as a ~hairline, so we draw 1px to
            // match the app rather than dropping a declared border.
            const totalLineWidth = chart.chartBorderWidthEmu
              ? Math.max(0.5, chart.chartBorderWidthEmu / EMU_PER_PT) * ptToPx
              : 1;
            target.setLineDash(dashPatternForLine(
              chart.chartBorderCustomDash, chart.chartBorderDash, totalLineWidth,
            ));
            target.lineCap = chart.chartBorderCap === 'rnd'
              ? 'round' : chart.chartBorderCap === 'sq' ? 'square' : 'butt';
            target.lineJoin = chart.chartBorderJoin === 'round' || chart.chartBorderJoin === 'bevel'
              ? chart.chartBorderJoin : 'miter';
            // Inset by half the line width so the full stroke stays inside the rect.
            strokeChartFrameRect(
              target, x, y, w, h, totalLineWidth, chart.chartBorderCompound,
              rounded ? cornerRadius : 0,
            );
            target.restore();
          }
        }
      },
    );

    // chartEx box-and-whisker / sunburst / treemap carry their data in the structured
    // `chartexBox` / `chartexSunburst` / `chartexTreemap` fields, not the flat `series` array, so the
    // empty-series "(no data)" guard must not fire for them.
    const hasChartexData = chart.chartexBox != null || chart.chartexSunburst != null
      || chart.chartexTreemap != null || chart.chartexRegionMap != null;
    if (chart.series.length === 0 && !hasChartexData) {
      // An authored series-less chart is Office's empty chart area.
      if (chart.authoredWithoutSeries !== true) {
        ctx.fillStyle = '#888';
        ctx.font = '12px sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('(no data)', x + w / 2, y + h / 2);
      }
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    const classicPointCount = classicCanvasPointCount(chart);
    const classicMarkerPaintWork = (classicPointCount != null || chart.chartexBox != null)
      && (classicPointCount ?? 0) <= MAX_CANVAS_CHART_POINTS
      ? classicMarkerPaintWorkCount(chart, imageLookup, ptToPx, rect) : null;
    const classicThreeDWork = classicThreeDWorkCount(chart, threeD);
    const classicDataMarkPaintWork = classicPointCount != null
      && classicPointCount <= MAX_CANVAS_CHART_POINTS
      ? classicDataMarkPaintWorkCount(
          chart, imageLookup, ptToPx, rect, classicThreeDWork != null,
        ) : null;
    const classicLabelPaintWork = chartLabelPaintWorkCount(
      chart, threeD, imageLookup, ptToPx, rect,
    );
    if (
      (classicPointCount != null || classicMarkerPaintWork != null
        || classicDataMarkPaintWork != null
        || classicThreeDWork != null || classicLabelPaintWork != null)
      && rejectOversizedCanvasChart(
        ctx,
        rect,
        Math.max(
          classicPointCount ?? 0,
          classicMarkerPaintWork != null
            && classicMarkerPaintWork + (classicDataMarkPaintWork ?? 0)
              > MAX_CANVAS_MARKER_PAINT_COMPONENTS
            ? MAX_CANVAS_CHART_POINTS + 1 : 0,
          classicDataMarkPaintWork != null
            && classicDataMarkPaintWork > MAX_CANVAS_MARKER_PAINT_COMPONENTS
            ? MAX_CANVAS_CHART_POINTS + 1 : 0,
          classicThreeDWork ?? 0,
          classicLabelPaintWork != null
            && classicLabelPaintWork > MAX_CANVAS_LABEL_PAINT_COMPONENTS
            ? MAX_CANVAS_CHART_POINTS + 1 : 0,
        ),
      )
    ) {
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    const plotDispatch = classicPlotDispatch(chart);
    if (plotDispatch === 'unsupported') {
      ctx.fillStyle = '#888';
      ctx.font = '11px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText('Unsupported chart', x + w / 2, y + h / 2);
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    if (plotDispatch !== 'legacy') {
      groupedFamilyRenderers[plotDispatch]?.(ctx, chart, rect, ptToPx, shapeRotationDeg);
      drawChartDisplayUnitLabels(ctx, chart, rect, ptToPx);
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    // Classic 3-D groups keep their canonical 2-D family name in the shared
    // model, while `threeD` carries the authored view/depth contract.  Consume
    // that contract before ordinary dispatch; 2-D charts return false and keep
    // their existing byte-stable family paths.
    if (threeD?.render(ctx, chart, rect, ptToPx, shapeRotationDeg)) {
      drawChartDisplayUnitLabels(ctx, chart, rect, ptToPx);
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }
    if (regionMap?.render(ctx, chart, rect, ptToPx, shapeRotationDeg)) {
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }
    if (chartEx?.render(ctx, chart, rect, ptToPx, shapeRotationDeg)) {
      drawChartDisplayUnitLabels(ctx, chart, rect, ptToPx);
      drawChartTextBoxes(ctx, chart, rect, ptToPx);
      return;
    }

    const familyRenderer = Object.hasOwn(classicFamilyRenderers, chart.chartType)
      ? classicFamilyRenderers[chart.chartType] : undefined;
    if (familyRenderer) {
      familyRenderer(ctx, chart, rect, ptToPx, shapeRotationDeg);
    } else {
      ctx.fillStyle = '#888';
      ctx.font = '11px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      // The public model can carry a future layout identifier of arbitrary
      // length. Preserve that identifier in the model, but keep the
      // fail-closed paint path constant-work instead of shaping attacker-
      // controlled text that is not part of the rendered document.
      ctx.fillText('Unsupported chart', x + w / 2, y + h / 2);
    }
    drawChartDisplayUnitLabels(ctx, chart, rect, ptToPx);
    drawChartTextBoxes(ctx, chart, rect, ptToPx);
  } finally {
    ctx.restore();
  }
}

export function renderChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number = PT_TO_PX,
  shapeRotationDeg = 0,
  threeD?: ChartThreeDRenderer,
  regionMap?: ChartRegionMapRenderer,
  imageLookup?: ChartImageLookup,
  chartEx?: ChartExRenderer,
): void {
  withChartStyleIndexCache(() => {
    withSparseStyleIndexCache(() => {
      withChartImageLookup(imageLookup, () => {
        const sourceStructureCount = sourceChartStructureCount(chart);
        // Project once before both the effect budget and paint dispatch. Hidden
        // source points must neither dilute a visible effect's allowance nor be
        // cloned a second time inside the renderer.
        const preparedChart = sourceStructureCount <= MAX_CANVAS_CHART_POINTS
          ? withEffectiveChartStyleRoles(applyPlotVisibleOnly(chart))
          : chart;
        const effectConsumers = sourceStructureCount <= MAX_CANVAS_CHART_POINTS
          ? chartEffectConsumerUpperBound(preparedChart)
          : 1;
        withChartEffectBudget(ctx, () => {
          renderChartImpl(
            ctx, preparedChart, rect, ptToPx, shapeRotationDeg,
            threeD, regionMap, imageLookup, chartEx, sourceStructureCount,
          );
        }, undefined, effectConsumers);
      });
    });
  });
}
