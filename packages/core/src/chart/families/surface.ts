// Classic surface chart family.
import type { ChartModel, ChartRect, ChartSeries } from '../../types/chart';
import type { Fill } from '../../types/common';

import { chartImageFillSource } from '../image-fill.js';
import { paintChartThreeDSurfacePicture } from '../three-d-surface-picture.js';
import { paintChartStyleOuterEffectBehind } from '../style-effects.js';

import { markerPaintComponents } from '../marker-style.js';

import {
  computeChartFrame,
  catAxisLabelBandH,
  chartLegendBands,
  chartTextFontSizePx,
} from '../layout.js';
import { automaticSurfaceMajorUnit, MAX_AXIS_TICKS } from '../axis-scale.js';
import { axisLineWidthPx, resolveGridline, isCrossBetween } from '../axis-style.js';
import { formatChartVal, formatChartValWithCode } from '../chart-number-format.js';

import {
  categoryMinorGridlineFractions,
  categoryLabelAnchorFraction,
  categoryLabelOffsetPx,
  categoryPositionFraction,
} from '../category-spacing.js';

import { effectiveChartStyleRole } from '../effective-style.js';

import { paintPlotAreaFrame } from '../plot-area-frame.js';
import {
  chartThreeDSurfacePaint,
  chartStyleDirectFillDecision,
  chartStyleDirectLineDecision,
  chartStyleDirectNoFillDecision,
  chartStyleDirectNoLineDecision,
  chartStyleFillDecision,
  chartStyleLineCascade,
  chartStyleLineDecision,
} from '../style-paint.js';

import { resolveFill } from '../../shape/paint.js';
import { drawingmlLineDashArray, pptxPresetDashArray } from '../../draw/dash.js';
import {
  isObservedAutomaticSurfaceCamera,
  scaleHexColor,
  surfaceMaterialFactor,
  surfacePerspectiveTangentGain,
} from '../material-color.js';
import {
  fitChartThreeDProjectionToWallThickness,
  planChartThreeDSurfaceGridSegments,
  planChartThreeDSurfaceGeometry,
  planChartThreeDProjection,
  type ThreeDScenePoint,
} from '../three-d.js';

import { chartColor } from '../shared/palette.js';
import { chartFontFamily, chartFontCss } from '../shared/fonts.js';
import { measuredLegendReserve, drawLegendForLayout } from '../shared/legend.js';
import { strokeAxisSegment, axisTickLengthPx, valMinorGridStroke, catMinorGridStroke, catGridlineFractions, valAxisReversed, catAxisReversed, drawValMajorGridlines, planValueAxis, axisLabelPx } from '../shared/axis.js';
import { measuredCartesianTitleBand, drawChartTitleForLayout } from '../shared/title.js';
import { chartCategories } from '../category-spacing.js';
import { MAX_CANVAS_MARKER_GRADIENT_STOPS, MAX_CANVAS_MARKER_PAINT_COMPONENTS } from '../shared/paint-limits.js';

// ═══════════════════════════════════════════════════════════════════════════
// Surface / contour chart (ECMA-376 §21.2.2.204)
// ═══════════════════════════════════════════════════════════════════════════
interface SurfaceVertex extends ThreeDScenePoint { value: number }

const surfaceCellTriangleIndices = (
  values: readonly [number, number, number, number],
): readonly [readonly [number, number, number], readonly [number, number, number]] =>
  values[0] + values[2] > values[1] + values[3]
    ? [[0, 1, 2], [0, 2, 3]]
    : [[0, 1, 3], [1, 2, 3]];

const MAX_SURFACE_PAINT_POLYGONS = 200_000;
const SURFACE_SCENE_DEPTH_SCALE = 1.25;

function clipSurfacePolygon(
  polygon: SurfaceVertex[],
  threshold: number,
  keepAbove: boolean,
): SurfaceVertex[] {
  if (polygon.length === 0) return [];
  const output: SurfaceVertex[] = [];
  const inside = (vertex: SurfaceVertex): boolean =>
    keepAbove ? vertex.value >= threshold : vertex.value <= threshold;
  let previous = polygon[polygon.length - 1];
  let previousInside = inside(previous);
  for (const current of polygon) {
    const currentInside = inside(current);
    if (currentInside !== previousInside) {
      const denominator = current.value - previous.value;
      const fraction = denominator === 0 ? 0 : (threshold - previous.value) / denominator;
      output.push({
        x: previous.x + (current.x - previous.x) * fraction,
        y: previous.y + (current.y - previous.y) * fraction,
        depth: previous.depth + (current.depth - previous.depth) * fraction,
        value: threshold,
      });
    }
    if (currentInside) output.push(current);
    previous = current;
    previousInside = currentInside;
  }
  return output;
}

export function renderSurfaceChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  r: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): void {
  const { x, y, w, h } = r;
  const categories = chartCategories(chart);
  const rows = chart.series;
  const columnCount = categories.length;
  const rowCount = rows.length;
  if (columnCount < 2 || rowCount < 2) return;

  let dataMin = Infinity;
  let dataMax = -Infinity;
  for (const row of rows) {
    for (let column = 0; column < columnCount; column++) {
      const value = row.values[column];
      if (value == null || !Number.isFinite(value)) continue;
      dataMin = Math.min(dataMin, value);
      dataMax = Math.max(dataMax, value);
    }
  }
  if (!Number.isFinite(dataMin) || !Number.isFinite(dataMax)) return;
  if (chart.valMin != null) dataMin = chart.valMin;
  if (chart.valMax != null) dataMax = chart.valMax;

  // Excel supplies the standard oblique perspective values for omitted
  // classic-Surface view fields. S1-S5 carry no c:view3D and isolate that
  // effective camera: parallel source-grid edges converge in the Office
  // output, so substituting a right-angle/orthographic camera loses the visible
  // depth. Keep these field defaults local to the Surface family; every
  // authored view3D field remains authoritative.
  const surfaceView = {
    ...(chart.threeD ?? {}),
    rotationX: chart.threeD?.rotationX ?? 15,
    rotationY: chart.threeD?.rotationY ?? 20,
    rightAngleAxes: chart.threeD?.rightAngleAxes ?? false,
    perspective: chart.threeD?.perspective ?? 30,
  };
  const observedAutomaticSurfaceCamera = isObservedAutomaticSurfaceCamera(surfaceView);
  const perspectiveTangentGain = surfacePerspectiveTangentGain(surfaceView);

  // Office bases an omitted Surface major unit on the projected value-axis
  // length. Probe the shared camera before reserving the legend: this is
  // independent of the eventual band count and correctly degenerates to the
  // compact five-interval class for a 90-degree contour view.
  const axisProbe = planChartThreeDProjection(surfaceView, r, {
    sceneDepthScale: SURFACE_SCENE_DEPTH_SCALE,
    perspectiveTangentGain,
  });
  let projectedValueAxisLenPt: number | undefined;
  if (axisProbe) {
    const probeX = axisProbe.topology.axisX === 'min'
      ? axisProbe.front.x : axisProbe.front.x + axisProbe.front.w;
    const probeBottom = axisProbe.project(
      probeX, axisProbe.front.y + axisProbe.front.h, axisProbe.topology.nearDepth,
    );
    const probeTop = axisProbe.project(
      probeX, axisProbe.front.y, axisProbe.topology.nearDepth,
    );
    projectedValueAxisLenPt = Math.hypot(
      probeTop.x - probeBottom.x,
      probeTop.y - probeBottom.y,
    ) / ptToPx;
  }
  const automaticSurfaceUnit = chart.valAxisMajorUnit == null
    ? automaticSurfaceMajorUnit(dataMin, dataMax, projectedValueAxisLenPt)
    : null;
  const provisionalAxis = planValueAxis(
    automaticSurfaceUnit == null
      ? chart
      : { ...chart, valAxisMajorUnit: automaticSurfaceUnit },
    dataMin,
    dataMax,
    projectedValueAxisLenPt,
  );
  // Surface value bands follow the shared value-axis plan. OOXML does not
  // define an Office-compatible automatic band count or implicit lighting, so
  // the renderer deliberately does not invent either here.
  const step = provisionalAxis.step;
  if (!(step > 0) || !Number.isFinite(step)) return;
  // Surface bands terminate on the first major boundary containing the data;
  // unlike ordinary value axes, Office does not append one headroom interval
  // when the maximum already lands on a boundary (S1/S4/S5 and their scaled
  // axis mirrors). Authored bounds remain authoritative.
  const surfaceMin = chart.valMin ?? provisionalAxis.min;
  const surfaceMax = chart.valMax ?? Math.max(
    surfaceMin + step,
    surfaceMin + Math.ceil((dataMax - surfaceMin) / step) * step,
  );
  const surfaceSpan = surfaceMax - surfaceMin;
  if (!(surfaceSpan > 0) || !Number.isFinite(surfaceSpan)) return;
  const rawBandCount = Math.ceil(surfaceSpan / step);
  const rawMajorLineCount = Math.floor(surfaceSpan / step + 1e-9) + 1;
  const triangleCount = (columnCount - 1) * (rowCount - 1) * 2;
  if (
    !Number.isSafeInteger(rawBandCount)
    || !Number.isSafeInteger(rawMajorLineCount)
    || rawBandCount < 1
    || rawMajorLineCount < 2
    || triangleCount < 1
    || rawBandCount > MAX_AXIS_TICKS
    || rawMajorLineCount > MAX_AXIS_TICKS
    || rawBandCount > Math.floor(MAX_SURFACE_PAINT_POLYGONS / triangleCount)
    || rawMajorLineCount > MAX_SURFACE_PAINT_POLYGONS
  ) return;
  const bandCount = rawBandCount;
  const surfaceFrac = valAxisReversed(chart)
    ? (value: number): number => 1 - (value - surfaceMin) / surfaceSpan
    : (value: number): number => (value - surfaceMin) / surfaceSpan;
  const surfaceMajorLines = Array.from(
    { length: rawMajorLineCount },
    (_, index) => surfaceMin + index * step,
  );
  const wireframeSplitFractions = (start: number, end: number): number[] => {
    const low = Math.min(start, end);
    const high = Math.max(start, end);
    const firstBoundary = Math.max(
      1,
      Math.floor((low - surfaceMin) / step) + 1,
    );
    const lastBoundary = Math.min(
      bandCount - 1,
      Math.ceil((high - surfaceMin) / step) - 1,
    );
    const fractions = [0];
    if (start !== end) {
      for (let boundary = firstBoundary; boundary <= lastBoundary; boundary++) {
        fractions.push((surfaceMin + boundary * step - start) / (end - start));
      }
    }
    fractions.push(1);
    fractions.sort((left, right) => left - right);
    return fractions;
  };
  const bandColors = Array.from({ length: bandCount }, (_, index) =>
    chart.themeAccentColors?.[index % 6]
      ? `#${chart.themeAccentColors[index % 6]}`
      : rows[index]?.color ? `#${rows[index].color}` : chartColor(index, rows[index]),
  );
  const bandFormats = new Map((chart.surfaceBandFormats ?? []).map(format => [format.idx, format]));
  // Surface value bands have their own semantic formatting-index domain. A
  // source-series-domain numeric role would repeat a two-series palette across
  // six value bands, so retain the raw linked role and select the separately
  // materialized numeric band role only after the final band count is known.
  const linkedBandStyle = chart.linkedChartStyleRoles?.dataPoint3D
    ?? (chart.classicChartStyleRoles == null ? chart.chartStyleRoles?.dataPoint3D : undefined);
  const numericBandStyle = bandCount <= 48
    ? chart.classicSurfaceBandStyles?.byBandCount?.[bandCount - 1]
      ?? chart.classicSurfaceBandStyles?.fixed
    : undefined;
  const effectiveBandStyle = effectiveChartStyleRole(numericBandStyle, linkedBandStyle);
  const linkedWireframeStyle = chart.linkedChartStyleRoles?.dataPointWireframe
    ?? (chart.classicChartStyleRoles == null
      ? chart.chartStyleRoles?.dataPointWireframe
      : undefined);
  const effectiveWireframeStyle = effectiveChartStyleRole(
    numericBandStyle,
    linkedWireframeStyle,
  );
  const styleHasLinePaint = (style: ChartSeries['chartexStyle']): boolean =>
    style?.lineNoStyle !== true && (
      style?.linePaintAuthored === true
      || style?.lineHidden === true
      || (style?.lineColors?.length ?? 0) > 0
      || (style?.linePaints?.length ?? 0) > 0
    );
  const bandFillPlans = chart.surfaceWireframe === true ? [] : Array.from(
    { length: bandCount }, (_, index) => {
      const format = bandFormats.get(index);
      let decision: Fill | null | undefined;
      let fromRole = true;
      if (format?.fillHidden === true) {
        decision = chartStyleDirectNoFillDecision(linkedBandStyle);
        if (decision !== undefined) fromRole = false;
      }
      else if (format?.fill) {
        decision = format.fill;
        fromRole = false;
      }
      else {
        const direct = chartStyleDirectFillDecision(
          format?.style, linkedBandStyle, index,
        );
        if (direct !== undefined) {
          decision = direct;
          fromRole = false;
        }
        else decision = chartStyleFillDecision(linkedBandStyle, index);
      }
      if (decision === undefined && format?.fillHidden === true) {
        decision = chartStyleFillDecision(linkedBandStyle, index);
      }
      if (decision === undefined) decision = chartStyleFillDecision(numericBandStyle, index);
      return {
        recipe: decision,
        // Keep provenance from the same modifier-aware cascade that selected
        // the recipe. A rejected local noFill leaves the linked/numeric role
        // responsible for Office's observed Surface material lighting.
        fromRole,
      };
    },
  );
  const bandFillRecipes = bandFillPlans.map(plan => plan.recipe);
  const bandLineRecipes = chart.surfaceWireframe === true ? [] : Array.from(
    { length: bandCount }, (_, index) => {
      const format = bandFormats.get(index);
      let decision: ChartModel['plotAreaLineFill'] | null | undefined;
      if (format?.lineHidden === true) {
        decision = chartStyleDirectNoLineDecision(linkedBandStyle);
      }
      else if (format?.lineColor) decision = { fillType: 'solid', color: format.lineColor };
      else {
        decision = chartStyleLineCascade(
          linkedBandStyle, linkedBandStyle, index, format?.style,
        );
      }
      if (decision === undefined && format?.lineHidden === true) {
        decision = chartStyleLineDecision(linkedBandStyle, index);
      }
      if (decision === undefined) decision = chartStyleLineDecision(numericBandStyle, index);
      return decision;
    },
  );
  interface SurfaceWireframeLineStyle {
    paint: ChartModel['plotAreaLineFill'] | null | undefined;
    lineWidthEmu: number | null | undefined;
    lineDash: string | null | undefined;
    lineCustomDash: ChartModel['plotAreaLineCustomDash'];
    lineCap: string | null | undefined;
    lineJoin: string | null | undefined;
    lineCompound: string | null | undefined;
  }
  const styleGeometry = (
    direct: ChartSeries['chartexStyle'],
    fallback: SurfaceWireframeLineStyle,
  ): Omit<SurfaceWireframeLineStyle, 'paint'> => {
    const effectiveDirect = direct;
    const dashAuthored = effectiveDirect?.lineDashAuthored === true
      || effectiveDirect?.lineDash != null
      || effectiveDirect?.lineCustomDash != null;
    return {
      lineWidthEmu: effectiveDirect?.lineWidthEmu ?? fallback.lineWidthEmu,
      lineDash: dashAuthored ? effectiveDirect?.lineDash : fallback.lineDash,
      lineCustomDash: dashAuthored
        ? effectiveDirect?.lineCustomDash : fallback.lineCustomDash,
      lineCap: effectiveDirect?.lineCap ?? fallback.lineCap,
      lineJoin: effectiveDirect?.lineJoin ?? fallback.lineJoin,
      lineCompound: effectiveDirect?.lineCompound ?? fallback.lineCompound,
    };
  };
  // A fixed dataPointWireframe reference supplies one mesh default. A relative
  // palette is indexed by value band below; selecting entry zero globally
  // would collapse the multicolour Surface wireframe defined by Tables 5/6.
  const fixedWireframeLineIndex = effectiveWireframeStyle?.lineColorIndex;
  let linkedWireframePaint: ChartModel['plotAreaLineFill'] | null | undefined;
  if (!styleHasLinePaint(effectiveWireframeStyle)) linkedWireframePaint = undefined;
  else if (effectiveWireframeStyle?.lineHidden === true) {
    linkedWireframePaint = chartStyleLineDecision(effectiveWireframeStyle, 0);
  } else if (fixedWireframeLineIndex != null) {
    linkedWireframePaint = fixedWireframeLineIndex != null
      ? chartStyleLineDecision(effectiveWireframeStyle, fixedWireframeLineIndex)
      : null;
  } else {
    linkedWireframePaint = undefined;
  }
  const linkedRelativeWireframePaints = fixedWireframeLineIndex == null
    && styleHasLinePaint(effectiveWireframeStyle)
    && effectiveWireframeStyle?.lineHidden !== true
    ? Array.from({ length: bandCount }, (_, index) =>
        chartStyleLineDecision(effectiveWireframeStyle, index))
    : [];
  const emptyWireframeStyle: SurfaceWireframeLineStyle = {
    paint: undefined,
    lineWidthEmu: undefined,
    lineDash: undefined,
    lineCustomDash: undefined,
    lineCap: undefined,
    lineJoin: undefined,
    lineCompound: undefined,
  };
  const linkedGeometry = effectiveWireframeStyle == null
    ? emptyWireframeStyle
    : {
      paint: linkedWireframePaint,
      lineWidthEmu: effectiveWireframeStyle?.lineWidthEmu,
      lineDash: effectiveWireframeStyle?.lineDash,
      lineCustomDash: effectiveWireframeStyle?.lineCustomDash,
      lineCap: effectiveWireframeStyle?.lineCap,
      lineJoin: effectiveWireframeStyle?.lineJoin,
      lineCompound: effectiveWireframeStyle?.lineCompound,
    };
  const firstSurfaceSeries = rows[0];
  let directSeriesPaint: ChartModel['plotAreaLineFill'] | null | undefined;
  if (firstSurfaceSeries?.lineHidden === true) {
    directSeriesPaint = chartStyleDirectNoLineDecision(linkedWireframeStyle);
  }
  else if (firstSurfaceSeries?.lineColor != null) {
    directSeriesPaint = { fillType: 'solid', color: firstSurfaceSeries.lineColor };
  } else {
    directSeriesPaint = chartStyleDirectLineDecision(
      firstSurfaceSeries?.chartexStyle, linkedWireframeStyle, 0,
    );
  }
  const baseGeometry = styleGeometry(firstSurfaceSeries?.chartexStyle, linkedGeometry);
  const baseWireframeStyle: SurfaceWireframeLineStyle = {
    paint: directSeriesPaint === undefined ? linkedWireframePaint : directSeriesPaint,
    ...baseGeometry,
    lineWidthEmu: firstSurfaceSeries?.lineWidthEmu ?? baseGeometry.lineWidthEmu,
  };
  // Canvas represents DrawingML `sng` directly. Only multi-rail compound
  // lines remain fail-closed here; treating every authored compound value as
  // unsupported suppresses the normal single line inherited from theme ln1.
  if (baseWireframeStyle.lineCompound != null
    && baseWireframeStyle.lineCompound !== 'sng') baseWireframeStyle.paint = null;
  const directBandLineDecisions = Array.from({ length: bandCount }, (_, index) => {
    const format = bandFormats.get(index);
    if (!format) return undefined;
    if (format.lineHidden === true) {
      return chartStyleDirectNoLineDecision(linkedWireframeStyle);
    }
    if (format.lineColor != null) return { fillType: 'solid' as const, color: format.lineColor };
    const direct = chartStyleDirectLineDecision(
      format.style, linkedWireframeStyle, index,
    );
    return direct;
  });
  const wireframeLineStyles = directBandLineDecisions.map((directPaint, index) => {
    const format = bandFormats.get(index);
    const geometry = styleGeometry(format?.style, baseWireframeStyle);
    const inheritedPaint = directSeriesPaint === undefined
      ? linkedRelativeWireframePaints[index] ?? baseWireframeStyle.paint
      : directSeriesPaint;
    const style: SurfaceWireframeLineStyle = {
      paint: directPaint === undefined ? inheritedPaint : directPaint,
      ...geometry,
      lineWidthEmu: format?.lineWidthEmu ?? geometry.lineWidthEmu,
    };
    if (style.lineCompound != null && style.lineCompound !== 'sng') style.paint = null;
    return style;
  });
  const sameWireframeStyle = (
    left: SurfaceWireframeLineStyle,
    right: SurfaceWireframeLineStyle,
  ): boolean => left.paint !== undefined
    && left.paint === right.paint
    && left.lineWidthEmu === right.lineWidthEmu
    && left.lineDash === right.lineDash
    && left.lineCustomDash === right.lineCustomDash
    && left.lineCap === right.lineCap
    && left.lineJoin === right.lineJoin
    && left.lineCompound === right.lineCompound;
  interface SurfaceWireframeBandRun {
    from: number;
    to: number;
    band: number;
  }
  const wireframeBandRuns = (start: number, end: number): SurfaceWireframeBandRun[] => {
    const fractions = wireframeSplitFractions(start, end);
    const runs: SurfaceWireframeBandRun[] = [];
    for (let index = 0; index < fractions.length - 1; index++) {
      const from = fractions[index];
      const to = fractions[index + 1];
      const midpoint = start + (end - start) * ((from + to) / 2);
      const band = Math.max(0, Math.min(
        bandCount - 1,
        Math.floor((midpoint - surfaceMin) / step),
      ));
      const previous = runs[runs.length - 1];
      if (previous
        && directBandLineDecisions[previous.band] === undefined
        && directBandLineDecisions[band] === undefined
        && sameWireframeStyle(
          wireframeLineStyles[previous.band], wireframeLineStyles[band],
        )) previous.to = to;
      else runs.push({ from, to, band });
    }
    return runs;
  };
  if (chart.surfaceWireframe === true) {
    let wireframeSegmentCount = 0;
    const chargeEdge = (start: number | null, end: number | null): boolean => {
      if (start == null || end == null || !Number.isFinite(start) || !Number.isFinite(end)) {
        return true;
      }
      wireframeSegmentCount += wireframeBandRuns(start, end).length;
      return wireframeSegmentCount <= MAX_SURFACE_PAINT_POLYGONS;
    };
    for (let row = 0; row < rowCount; row++) {
      for (let column = 0; column < columnCount - 1; column++) {
        if (!chargeEdge(rows[row].values[column], rows[row].values[column + 1])) return;
      }
    }
    for (let column = 0; column < columnCount; column++) {
      for (let row = 0; row < rowCount - 1; row++) {
        if (!chargeEdge(rows[row].values[column], rows[row + 1].values[column])) return;
      }
    }
    // A wireframe Surface contains both the source row/column mesh and the
    // contour at each value-band boundary. Charge the latter before any
    // projection or paint allocation, using the same cell triangulation as
    // the renderer below.
    for (let row = 0; row < rowCount - 1; row++) {
      for (let column = 0; column < columnCount - 1; column++) {
        const values = [
          rows[row].values[column],
          rows[row].values[column + 1],
          rows[row + 1].values[column + 1],
          rows[row + 1].values[column],
        ];
        if (values.some(value => value == null || !Number.isFinite(value))) continue;
        const finiteValues = values as [number, number, number, number];
        for (const indices of surfaceCellTriangleIndices(finiteValues)) {
          const triangleMin = Math.min(...indices.map(index => finiteValues[index]));
          const triangleMax = Math.max(...indices.map(index => finiteValues[index]));
          const firstBoundary = Math.max(
            1,
            Math.floor((triangleMin - surfaceMin) / step) + 1,
          );
          const lastBoundary = Math.min(
            bandCount - 1,
            Math.ceil((triangleMax - surfaceMin) / step) - 1,
          );
          if (lastBoundary < firstBoundary) continue;
          wireframeSegmentCount += lastBoundary - firstBoundary + 1;
          if (wireframeSegmentCount > MAX_SURFACE_PAINT_POLYGONS) return;
        }
      }
    }
  }
  const usesBaseWireframeLine = directBandLineDecisions.some(
    decision => decision === undefined,
  ) && (directSeriesPaint !== undefined || linkedRelativeWireframePaints.length === 0);
  const surfaceFacePaints = [
    { surface: chart.threeD?.floor, role: 'floor' as const },
    { surface: chart.threeD?.sideWall, role: 'wall' as const },
    { surface: chart.threeD?.backWall, role: 'wall' as const },
  ].map(({ surface, role }) => chartThreeDSurfacePaint(chart, surface, role));
  let surfacePaintComponents = 0;
  for (const recipe of [
    ...bandFillRecipes,
    ...bandLineRecipes,
    ...(chart.surfaceWireframe === true
      && usesBaseWireframeLine
      ? [baseWireframeStyle.paint]
      : []),
    ...(chart.surfaceWireframe === true ? linkedRelativeWireframePaints : []),
    ...(chart.surfaceWireframe === true
      ? directBandLineDecisions.filter(decision => decision !== undefined)
      : []),
    ...surfaceFacePaints.flatMap(paint => [paint.fill, paint.line]),
  ]) {
    if (recipe == null) continue;
    const components = markerPaintComponents(recipe);
    if ((recipe.fillType === 'gradient'
        && components > MAX_CANVAS_MARKER_GRADIENT_STOPS)
      || components > MAX_CANVAS_MARKER_PAINT_COMPONENTS - surfacePaintComponents) return;
    surfacePaintComponents += components;
  }
  const bandLabels = Array.from({ length: bandCount }, (_, index) => {
    const lower = surfaceMin + index * step;
    const upper = Math.min(surfaceMax, lower + step);
    return `${formatChartVal(lower)}-${formatChartVal(upper)}`;
  });
  const legendChart: ChartModel = {
    ...chart,
    series: bandLabels.map((name, index) => {
      const fill = bandFillRecipes[index];
      const line = bandLineRecipes[index];
      return {
        name,
        color: fill === null
          ? '00000000'
          : fill?.fillType === 'solid'
            ? fill.color.replace(/^#/, '')
            : bandColors[index].replace(/^#/, ''),
        ...(line?.fillType === 'solid'
          ? { lineColor: line.color.replace(/^#/, '') }
          : line === null ? { lineHidden: true } : {}),
        values: [],
      };
    }),
  };
  const legendFillPaints = [...bandFillRecipes];
  // A top-down contour lists the highest band first. In the ordinary oblique
  // Surface view Office lists low-to-high. S1-S5 and the 90° contour boundary
  // isolate this to the authored/effective camera elevation, not legend side.
  if (Math.abs(surfaceView.rotationX) === 90) {
    legendChart.series.reverse();
    legendFillPaints.reverse();
  }
  const legend = measuredLegendReserve(ctx, legendChart, w, h, 0.22, ptToPx);
  const { legRightW, legLeftW, legTopH, legBottomH } = chartLegendBands(
    legend, chart.legendOverlay === true,
  );
  const titleBand = measuredCartesianTitleBand(ctx, chart, w, h, ptToPx);
  const catFontPx = axisLabelPx(chart.catAxisFontSizeHpt, h, ptToPx);
  const seriesAxis = chart.threeD?.seriesAxis;
  const seriesFontPx = chartTextFontSizePx(seriesAxis?.fontSizeHpt, ptToPx) ?? catFontPx;
  const pad = {
    t: titleBand.bandH + legTopH + seriesFontPx / 2,
    r: legRightW + seriesFontPx * 3.2 + 12,
    b: catAxisLabelBandH(catFontPx, chart.catAxisLabelOffsetPercent) + legBottomH,
    l: legLeftW + catFontPx * 1.5,
  };
  const frame = computeChartFrame(chart, x, y, w, h, ptToPx, {
    titleBand,
    legendSideReserveFrac: 0.22,
    legendReserve: legend,
    pad,
    honorPlotAreaManualLayout: true,
  });
  const { px0, py0, pw, ph } = frame.plotRect;
  if (!(pw > 0) || !(ph > 0)) return;
  drawChartTitleForLayout(
    ctx, chart,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? x : px0, y,
    chart.titleManualLayout || !chart.titleRichRuns?.length ? w : pw, h,
    y + titleBand.topPad, titleBand.fontPx,
  );
  paintPlotAreaFrame(ctx, chart, px0, py0, pw, ph, ptToPx, shapeRotationDeg);

  let projection = planChartThreeDProjection(surfaceView, { x: px0, y: py0, w: pw, h: ph }, {
    // Surface rows occupy a real series axis rather than the compact prism
    // slab used by 3-D columns. The boundary corpus fits that grid with the
    // same depth occupancy as standard (depth-arranged) cartesian series.
    sceneDepthScale: SURFACE_SCENE_DEPTH_SCALE,
    perspectiveTangentGain,
  });
  if (!projection) return;
  projection = fitChartThreeDProjectionToWallThickness(
    projection,
    chart.threeD ?? {},
    { x: px0, y: py0, w: pw, h: ph },
  );
  const useObservedAutomaticMaterial = observedAutomaticSurfaceCamera || (
    Math.abs(surfaceView.rotationX) === 90
    && surfaceView.rotationY === 0
    && surfaceView.rightAngleAxes === false
    && surfaceView.perspective === 0
  );
  const { front } = projection;
  const categoryReversed = chart.catAxisOrientation === 'maxMin';
  const categoryBetween = isCrossBetween(chart);
  const toX = (column: number): number => front.x + categoryPositionFraction(
    column, columnCount, categoryBetween, categoryReversed,
  ) * front.w;
  const seriesReversed = seriesAxis?.orientation === 'maxMin';
  // Surface rows are values on `<c:serAx>`, whose schema has no
  // `<c:crossBetween>`. Office places the first and last rows on the series
  // axis endpoints, unlike category points centred by value-axis
  // crossBetween="between". Reuse the shared endpoint placement so reversal
  // and the single-row midpoint remain consistent with other ordinal axes.
  const toDepth = (row: number): number => categoryPositionFraction(
    row, rowCount, false, seriesReversed,
  );
  const toValueY = (value: number): number =>
    front.y + front.h - surfaceFrac(value) * front.h;
  interface SurfacePaint {
    points: Array<{ x: number; y: number }>;
    scenePoints: SurfaceVertex[];
    band: number;
    depth: number;
  }
  const paints: SurfacePaint[] = [];
  interface SurfaceWireframeSegment {
    points: [{ x: number; y: number }, { x: number; y: number }];
    band: number;
  }
  const wireframeSegments: SurfaceWireframeSegment[] = [];
  const appendWireframeEdge = (start: SurfaceVertex, end: SurfaceVertex): void => {
    const pointAt = (fraction: number): SurfaceVertex => ({
      x: start.x + (end.x - start.x) * fraction,
      y: start.y + (end.y - start.y) * fraction,
      depth: start.depth + (end.depth - start.depth) * fraction,
      value: start.value + (end.value - start.value) * fraction,
    });
    for (const run of wireframeBandRuns(start.value, end.value)) {
      const from = pointAt(run.from);
      const to = pointAt(run.to);
      wireframeSegments.push({
        points: [
          projection.project(from.x, from.y, from.depth),
          projection.project(to.x, to.y, to.depth),
        ],
        band: run.band,
      });
    }
  };
  const appendWireframeContours = (triangle: readonly SurfaceVertex[]): void => {
    const triangleMin = Math.min(...triangle.map(vertex => vertex.value));
    const triangleMax = Math.max(...triangle.map(vertex => vertex.value));
    const firstBoundary = Math.max(
      1,
      Math.floor((triangleMin - surfaceMin) / step) + 1,
    );
    const lastBoundary = Math.min(
      bandCount - 1,
      Math.ceil((triangleMax - surfaceMin) / step) - 1,
    );
    for (let boundary = firstBoundary; boundary <= lastBoundary; boundary++) {
      const threshold = surfaceMin + boundary * step;
      const intersections: SurfaceVertex[] = [];
      const addIntersection = (vertex: SurfaceVertex): void => {
        if (intersections.some(existing =>
          Math.abs(existing.x - vertex.x) < 1e-9
          && Math.abs(existing.y - vertex.y) < 1e-9
          && Math.abs(existing.depth - vertex.depth) < 1e-9
        )) return;
        intersections.push(vertex);
      };
      for (let index = 0; index < triangle.length; index++) {
        const start = triangle[index];
        const end = triangle[(index + 1) % triangle.length];
        if (start.value === threshold) addIntersection(start);
        if ((start.value < threshold && end.value > threshold)
          || (start.value > threshold && end.value < threshold)) {
          const fraction = (threshold - start.value) / (end.value - start.value);
          addIntersection({
            x: start.x + (end.x - start.x) * fraction,
            y: start.y + (end.y - start.y) * fraction,
            depth: start.depth + (end.depth - start.depth) * fraction,
            value: threshold,
          });
        }
      }
      if (intersections.length !== 2) continue;
      wireframeSegments.push({
        points: [
          projection.project(
            intersections[0].x, intersections[0].y, intersections[0].depth,
          ),
          projection.project(
            intersections[1].x, intersections[1].y, intersections[1].depth,
          ),
        ],
        // A boundary closes the band below it. This keeps c:bandFmt indexing
        // consistent with the lower-inclusive clipping used by filled Surface.
        band: boundary - 1,
      });
    }
  };
  const paintTriangle = (triangle: SurfaceVertex[]): void => {
    if (chart.surfaceWireframe === true) {
      appendWireframeContours(triangle);
      return;
    }
    const triangleMin = Math.min(...triangle.map(vertex => vertex.value));
    const triangleMax = Math.max(...triangle.map(vertex => vertex.value));
    const firstBand = Math.max(0, Math.floor((triangleMin - surfaceMin) / step));
    const lastBand = Math.min(
      bandCount - 1,
      Math.floor((triangleMax - surfaceMin) / step),
    );
    for (let band = firstBand; band <= lastBand; band++) {
      const lower = surfaceMin + band * step;
      const upper = band === bandCount - 1 ? surfaceMax : lower + step;
      let polygon = clipSurfacePolygon(triangle, lower, true);
      polygon = clipSurfacePolygon(polygon, upper, false);
      if (polygon.length < 3) continue;
      const points = polygon.map(vertex => projection.project(vertex.x, vertex.y, vertex.depth));
      paints.push({
        points,
        scenePoints: polygon,
        band,
        depth: polygon.reduce(
          (sum, vertex) => sum + projection.cameraDepth(vertex.x, vertex.y, vertex.depth),
          0,
        ) / polygon.length,
      });
    }
  };

  ctx.save();
  ctx.beginPath();
  ctx.rect(px0, py0, pw, ph);
  ctx.clip();

  const strokeScenePath = (
    scenePoints: readonly ThreeDScenePoint[],
    color: string,
    width: number,
    dash: string | null | undefined,
    unbounded = false,
  ): void => {
    if (scenePoints.length < 2) return;
    const points = unbounded
      ? scenePoints.map(point => projection.projectUnbounded(point.x, point.y, point.depth))
      : scenePoints.map(point => projection.project(point.x, point.y, point.depth));
    ctx.beginPath();
    ctx.moveTo(points[0].x, points[0].y);
    for (let index = 1; index < points.length; index++) ctx.lineTo(points[index].x, points[index].y);
    ctx.strokeStyle = color;
    ctx.lineWidth = width;
    ctx.setLineDash(pptxPresetDashArray(dash ?? 'solid', width));
    ctx.stroke();
  };
  const farDepth = projection.topology.farDepth;
  const nearDepth = projection.topology.nearDepth;
  const floorY = front.y + front.h;
  const wallTopY = front.y;
  const farX = projection.topology.farX === 'min' ? front.x : front.x + front.w;
  const surfaceSlabs = (
    ['floor', 'sideWall', 'backWall'] as const
  ).map(kind => {
    const surface = chart.threeD?.[kind];
    return planChartThreeDSurfaceGeometry(projection, kind, surface?.thicknessPercent);
  });
  const surfaceKinds = ['floor', 'sideWall', 'backWall'] as const;
  const strokeSurfaceGridRule = (
    slabIndex: number,
    coordinate: 'x' | 'y',
    fraction: number,
    color: string,
    width: number,
    dash: string | null | undefined,
  ): void => {
    const slab = surfaceSlabs[slabIndex];
    const kind = surfaceKinds[slabIndex];
    for (const segment of planChartThreeDSurfaceGridSegments(
      slab,
      kind,
      coordinate,
      fraction,
    )) {
      if (slab.thickness > 0
        && !projection.cameraFacing(slab.faces[segment.faceIndex])) continue;
      strokeScenePath(segment.scenePoints, color, width, dash, true);
    }
  };
  const strokeAuthoredValueSurfaceRules = (
    values: readonly number[],
    color: string,
    width: number,
    dash: string | null | undefined,
  ): void => {
    for (const value of values) {
      const fraction = surfaceFrac(value);
      strokeSurfaceGridRule(1, 'y', fraction, color, width, dash);
      strokeSurfaceGridRule(2, 'y', fraction, color, width, dash);
    }
  };
  const strokeAuthoredCategorySurfaceRules = (
    fractions: readonly number[],
    color: string,
    width: number,
    dash: string | null | undefined,
  ): void => {
    for (const fraction of fractions) {
      strokeSurfaceGridRule(0, 'x', fraction, color, width, dash);
      strokeSurfaceGridRule(2, 'x', fraction, color, width, dash);
    }
  };
  // Keep the pre-thickness Surface3D path byte-stable when all three values
  // are omitted/zero. Its existing floor plane is family-owned; positive
  // thickness opts into the shared camera-aware CT_Surface slabs.
  const surfaceFaceGroups = surfaceSlabs.some(slab => slab.thickness > 0)
    ? surfaceSlabs.map(slab => slab.faces
      .filter(face => slab.thickness === 0 || projection.cameraFacing(face))
      .map(face => face.map(point =>
        projection.projectUnbounded(point.x, point.y, point.depth)
      )))
    : [
      [
        projection.project(front.x, floorY, nearDepth),
        projection.project(front.x + front.w, floorY, nearDepth),
        projection.project(front.x + front.w, floorY, farDepth),
        projection.project(front.x, floorY, farDepth),
      ],
      [
        projection.project(farX, floorY, nearDepth),
        projection.project(farX, floorY, farDepth),
        projection.project(farX, wallTopY, farDepth),
        projection.project(farX, wallTopY, nearDepth),
      ],
      [
        projection.project(front.x, floorY, farDepth),
        projection.project(front.x + front.w, floorY, farDepth),
        projection.project(front.x + front.w, wallTopY, farDepth),
        projection.project(front.x, wallTopY, farDepth),
      ],
    ].map(face => [face]);
  for (let index = 0; index < surfaceFaceGroups.length; index++) {
    const faces = surfaceFaceGroups[index];
    if (!faces.length) continue;
    const points = faces.flat();
    const effective = surfaceFacePaints[index];
    const imageFill = effective.fill?.fillType === 'image' ? effective.fill : null;
    if (imageFill) {
      const image = chartImageFillSource(imageFill);
      const surface = chart.threeD?.[surfaceKinds[index]];
      if (image) {
        const project = (point: ThreeDScenePoint) =>
          projection.projectUnbounded(point.x, point.y, point.depth);
        paintChartThreeDSurfacePicture(
          ctx, imageFill, image, surface, surfaceKinds[index],
          surfaceSlabs[index], surfaceSlabs[index].faces
            .map((face, faceIndex) => ({ face, faceIndex }))
            .filter(({ face }) => surfaceSlabs[index].thickness === 0
              || projection.cameraFacing(face))
            .map(({ faceIndex }) => faceIndex),
          project, surfaceSpan,
        );
      }
    }
    const minX = Math.min(...points.map(point => point.x));
    const maxX = Math.max(...points.map(point => point.x));
    const minY = Math.min(...points.map(point => point.y));
    const maxY = Math.max(...points.map(point => point.y));
    const fill = imageFill
      ? null
      : effective.fill?.fillType === 'solid'
      ? `#${effective.fill.color}`
      : effective.fill
        ? resolveFill(effective.fill, ctx, minX, minY, maxX - minX, maxY - minY)
        : null;
    const line = effective.line?.fillType === 'solid'
      ? `#${effective.line.color}`
      : effective.line
        ? resolveFill(effective.line, ctx, minX, minY, maxX - minX, maxY - minY)
        : null;
    const width = effective.lineWidthEmu != null
      ? axisLineWidthPx(effective.lineWidthEmu, ptToPx) : 1;
    for (const face of faces) {
      ctx.beginPath();
      ctx.moveTo(face[0].x, face[0].y);
      for (let pointIndex = 1; pointIndex < face.length; pointIndex++) {
        ctx.lineTo(face[pointIndex].x, face[pointIndex].y);
      }
      ctx.closePath();
      if (fill) {
        ctx.fillStyle = fill;
        ctx.fill();
      }
      if (line) {
        ctx.strokeStyle = line;
        ctx.lineWidth = width;
        ctx.setLineDash(drawingmlLineDashArray(
          effective.lineCustomDash,
          effective.lineDash,
          width,
        ));
        ctx.lineCap = effective.lineCap === 'rnd'
          ? 'round' : effective.lineCap === 'sq' ? 'square' : 'butt';
        ctx.lineJoin = effective.lineJoin === 'round' || effective.lineJoin === 'bevel'
          ? effective.lineJoin : 'miter';
        ctx.stroke();
      }
    }
  }
  if (chart.valAxisMinorGridlines === true) {
    const line = valMinorGridStroke(chart, ptToPx);
    strokeAuthoredValueSurfaceRules(
      provisionalAxis.minorLines.filter(value => value >= surfaceMin && value <= surfaceMax),
      line.color,
      line.width,
      chart.valAxisMinorGridlineDash,
    );
  }
  if (drawValMajorGridlines(chart)) {
    const line = resolveGridline(
      chart.valAxisGridlineColor,
      chart.valAxisGridlineWidthEmu,
      ptToPx,
    );
    if (chart.valAxisMajorGridlines === true) {
      strokeAuthoredValueSurfaceRules(
        surfaceMajorLines,
        line.color,
        line.width,
        chart.valAxisGridlineDash,
      );
    } else {
      for (const value of surfaceMajorLines) {
        const fraction = surfaceFrac(value);
        const gridY = toValueY(value);
        if (surfaceSlabs[2].thickness > 0) {
          strokeSurfaceGridRule(
            2, 'y', fraction, line.color, line.width, chart.valAxisGridlineDash,
          );
        } else {
          strokeScenePath([
            { x: front.x, y: gridY, depth: farDepth },
            { x: front.x + front.w, y: gridY, depth: farDepth },
          ], line.color, line.width, chart.valAxisGridlineDash);
        }
        if (surfaceSlabs[1].thickness > 0) {
          strokeSurfaceGridRule(
            1, 'y', fraction, line.color, line.width, chart.valAxisGridlineDash,
          );
        } else {
          strokeScenePath([
            { x: farX, y: gridY, depth: nearDepth },
            { x: farX, y: gridY, depth: farDepth },
          ], line.color, line.width, chart.valAxisGridlineDash);
        }
      }
    }
  }
  if (chart.catAxisMinorGridlines === true) {
    const line = catMinorGridStroke(chart, ptToPx);
    strokeAuthoredCategorySurfaceRules(
      categoryMinorGridlineFractions(columnCount, categoryBetween),
      line.color,
      line.width,
      chart.catAxisMinorGridlineDash,
    );
  }
  if (chart.catAxisMajorGridlines) {
    const line = resolveGridline(
      chart.catAxisGridlineColor,
      chart.catAxisGridlineWidthEmu,
      ptToPx,
    );
    // Gridlines mark category boundaries under crossBetween="between"; data
    // points remain at the interval centres. Using `toX(column)` here would
    // incorrectly draw lines through the 25%/75% data points of a two-column
    // Surface instead of the 0%/50%/100% boundaries authored by the axis.
    strokeAuthoredCategorySurfaceRules(
      catGridlineFractions(chart, columnCount),
      line.color,
      line.width,
      chart.catAxisGridlineDash,
    );
  }

  for (let row = 0; row < rowCount - 1; row++) {
    for (let column = 0; column < columnCount - 1; column++) {
      const values = [
        rows[row].values[column],
        rows[row].values[column + 1],
        rows[row + 1].values[column + 1],
        rows[row + 1].values[column],
      ];
      if (values.some(value => value == null || !Number.isFinite(value))) continue;
      const rowColumn = {
        x: toX(column), y: toValueY(values[0] as number), depth: toDepth(row),
        value: values[0] as number,
      };
      const rowColumnNext = {
        x: toX(column + 1), y: toValueY(values[1] as number), depth: toDepth(row),
        value: values[1] as number,
      };
      const rowNextColumnNext = {
        x: toX(column + 1), y: toValueY(values[2] as number), depth: toDepth(row + 1),
        value: values[2] as number,
      };
      const rowNextColumn = {
        x: toX(column), y: toValueY(values[3] as number), depth: toDepth(row + 1),
        value: values[3] as number,
      };
      // Office's filled Surface keeps the upper of the two possible cell
      // diagonals. The 2x2 plane/saddle/reversed-saddle boundaries isolate all
      // three cases: equal opposing sums retain the source-grid B-D diagonal;
      // otherwise the opposing pair with the larger value sum forms the ridge.
      // This is an application rendering rule — OOXML stores only the matrix.
      const vertices = [
        rowColumn, rowColumnNext, rowNextColumnNext, rowNextColumn,
      ] as const;
      for (const indices of surfaceCellTriangleIndices(values as [number, number, number, number])) {
        paintTriangle(indices.map(index => vertices[index]));
      }
    }
  }
  if (chart.surfaceWireframe === true) {
    for (let row = 0; row < rowCount; row++) {
      for (let column = 0; column < columnCount - 1; column++) {
        const startValue = rows[row].values[column];
        const endValue = rows[row].values[column + 1];
        if (startValue == null || endValue == null
          || !Number.isFinite(startValue) || !Number.isFinite(endValue)) continue;
        appendWireframeEdge(
          {
            x: toX(column), y: toValueY(startValue), depth: toDepth(row), value: startValue,
          },
          {
            x: toX(column + 1), y: toValueY(endValue), depth: toDepth(row), value: endValue,
          },
        );
      }
    }
    for (let column = 0; column < columnCount; column++) {
      for (let row = 0; row < rowCount - 1; row++) {
        const startValue = rows[row].values[column];
        const endValue = rows[row + 1].values[column];
        if (startValue == null || endValue == null
          || !Number.isFinite(startValue) || !Number.isFinite(endValue)) continue;
        appendWireframeEdge(
          {
            x: toX(column), y: toValueY(startValue), depth: toDepth(row), value: startValue,
          },
          {
            x: toX(column), y: toValueY(endValue), depth: toDepth(row + 1), value: endValue,
          },
        );
      }
    }
  }
  paints.sort((left, right) => left.depth - right.depth);
  const bandBounds = Array.from({ length: bandCount }, () => ({
    minX: Number.POSITIVE_INFINITY,
    minY: Number.POSITIVE_INFINITY,
    maxX: Number.NEGATIVE_INFINITY,
    maxY: Number.NEGATIVE_INFINITY,
  }));
  const bandFaceGroups = Array.from(
    { length: bandCount },
    () => [] as Array<(typeof paints)[number]>,
  );
  for (const paint of paints) {
    bandFaceGroups[paint.band].push(paint);
    const bounds = bandBounds[paint.band];
    for (const point of paint.points) {
      bounds.minX = Math.min(bounds.minX, point.x);
      bounds.minY = Math.min(bounds.minY, point.y);
      bounds.maxX = Math.max(bounds.maxX, point.x);
      bounds.maxY = Math.max(bounds.maxY, point.y);
    }
  }
  for (const segment of wireframeSegments) {
    const bounds = bandBounds[segment.band];
    for (const point of segment.points) {
      bounds.minX = Math.min(bounds.minX, point.x);
      bounds.minY = Math.min(bounds.minY, point.y);
      bounds.maxX = Math.max(bounds.maxX, point.x);
      bounds.maxY = Math.max(bounds.maxY, point.y);
    }
  }
  type SurfaceCanvasPaint = string | CanvasGradient | CanvasPattern | null | undefined;
  const resolveBandPaint = (
    recipe: Fill | null | undefined,
    band: number,
  ): SurfaceCanvasPaint => {
    if (recipe == null) return recipe;
    if (recipe.fillType === 'solid') return `#${recipe.color}`;
    const bounds = bandBounds[band];
    if (!Number.isFinite(bounds.minX) || !Number.isFinite(bounds.minY)
      || !Number.isFinite(bounds.maxX) || !Number.isFinite(bounds.maxY)) return null;
    return resolveFill(
      recipe,
      ctx,
      bounds.minX,
      bounds.minY,
      bounds.maxX - bounds.minX,
      bounds.maxY - bounds.minY,
    );
  };
  // A c:bandFmt styles one logical band. Resolve its DrawingML recipes once
  // against the complete projected band bounds, then reuse the Canvas paint
  // for every clipped polygon instead of replaying gradient stops per face.
  const resolvedBandFills = bandFillRecipes.map(resolveBandPaint);
  const resolvedBandLines = bandLineRecipes.map(resolveBandPaint);
  for (let band = 0; band < bandCount; band++) {
    if (resolvedBandFills[band] === null) continue;
    const bandFaces = bandFaceGroups[band].filter(face => face.points.length >= 3);
    if (!bandFaces.length) continue;
    paintChartStyleOuterEffectBehind(
      ctx,
      bandFormats.get(band)?.style,
      effectiveBandStyle,
      0,
      ptToPx,
      target => {
        target.beginPath();
        for (const face of bandFaces) {
          target.moveTo(face.points[0].x, face.points[0].y);
          for (let index = 1; index < face.points.length; index++) {
            target.lineTo(face.points[index].x, face.points[index].y);
          }
          target.closePath();
        }
        target.fillStyle = '#000000';
        target.fill();
      },
      band,
    );
  }
  for (const paint of paints) {
    const format = bandFormats.get(paint.band);
    ctx.beginPath();
    ctx.moveTo(paint.points[0].x, paint.points[0].y);
    for (let index = 1; index < paint.points.length; index++) {
      ctx.lineTo(paint.points[index].x, paint.points[index].y);
    }
    ctx.closePath();
    const bandFill = resolvedBandFills[paint.band];
    if (bandFill !== null) {
      const materialFactor = useObservedAutomaticMaterial
        ? surfaceMaterialFactor(projection.cameraNormal(paint.scenePoints))
        : 1;
      // A linked/numeric Data Point (3-D) fill supplies the base material
      // colour; it does not turn a Surface face into directly-authored flat
      // paint. Preserve Office's observed camera lighting for solid role fills.
      ctx.fillStyle = typeof bandFill === 'string' && bandFillPlans[paint.band]?.fromRole
        ? scaleHexColor(bandFill, materialFactor)
        : bandFill ?? scaleHexColor(bandColors[paint.band], materialFactor);
      ctx.fill();
    }
    const bandLine = resolvedBandLines[paint.band];
    if (bandLine != null) {
      const directGeometry = format?.style;
      const linkedGeometry = linkedBandStyle;
      const numericGeometry = numericBandStyle;
      ctx.strokeStyle = bandLine;
      const widthEmu = format?.lineWidthEmu
        ?? directGeometry?.lineWidthEmu ?? linkedGeometry?.lineWidthEmu
        ?? numericGeometry?.lineWidthEmu;
      ctx.lineWidth = widthEmu != null
        ? axisLineWidthPx(widthEmu, ptToPx)
        : 1;
      ctx.setLineDash(drawingmlLineDashArray(
        directGeometry?.lineCustomDash ?? linkedGeometry?.lineCustomDash
          ?? numericGeometry?.lineCustomDash,
        directGeometry?.lineDash ?? linkedGeometry?.lineDash ?? numericGeometry?.lineDash,
        ctx.lineWidth,
      ));
      const cap = directGeometry?.lineCap ?? linkedGeometry?.lineCap ?? numericGeometry?.lineCap;
      const join = directGeometry?.lineJoin ?? linkedGeometry?.lineJoin ?? numericGeometry?.lineJoin;
      ctx.lineCap = cap === 'rnd' ? 'round' : cap === 'sq' ? 'square' : 'butt';
      ctx.lineJoin = join === 'round' || join === 'bevel' ? join : 'miter';
      ctx.stroke();
    }
  }
  if (chart.surfaceWireframe === true) {
    const baseWireframeLine = !usesBaseWireframeLine
      ? undefined
      : baseWireframeStyle.paint?.fillType === 'solid'
      ? `#${baseWireframeStyle.paint.color}`
      : baseWireframeStyle.paint
        ? resolveFill(baseWireframeStyle.paint, ctx, px0, py0, pw, ph)
        : baseWireframeStyle.paint;
    const resolvedWireframeLines = directBandLineDecisions.map((decision, band) => {
      if (decision !== undefined) return resolveBandPaint(decision, band);
      if (directSeriesPaint !== undefined || linkedRelativeWireframePaints.length === 0) {
        return baseWireframeLine;
      }
      return resolveBandPaint(linkedRelativeWireframePaints[band], band);
    });
    for (const segment of wireframeSegments) {
      const style = wireframeLineStyles[segment.band];
      const line = resolvedWireframeLines[segment.band];
      if (line === null) continue;
      ctx.beginPath();
      ctx.moveTo(segment.points[0].x, segment.points[0].y);
      ctx.lineTo(segment.points[1].x, segment.points[1].y);
      ctx.strokeStyle = line ?? bandColors[segment.band];
      ctx.lineWidth = style.lineWidthEmu != null
        ? axisLineWidthPx(style.lineWidthEmu, ptToPx)
        : Math.max(1, 0.75 * ptToPx);
      ctx.setLineDash(drawingmlLineDashArray(
        style.lineCustomDash,
        style.lineDash,
        ctx.lineWidth,
      ));
      const cap = style.lineCap;
      const join = style.lineJoin;
      ctx.lineCap = cap === 'rnd' ? 'round' : cap === 'sq' ? 'square' : 'butt';
      ctx.lineJoin = join === 'round' || join === 'bevel' ? join : 'miter';
      ctx.stroke();
    }
  }
  ctx.restore();

  const sceneCenter = projection.project(
    front.x + front.w / 2,
    front.y + front.h / 2,
    0.5,
  );
  const drawSurfaceAxisTick = (
    mode: string | null | undefined,
    point: { x: number; y: number },
    axisStart: { x: number; y: number },
    axisEnd: { x: number; y: number },
    color: string,
    lineWidth: number,
    lineHidden: boolean,
    level: 'major' | 'minor',
    dash?: string | null,
  ): void => {
    if (lineHidden || mode == null || mode === 'none') return;
    const dx = axisEnd.x - axisStart.x;
    const dy = axisEnd.y - axisStart.y;
    const axisLength = Math.hypot(dx, dy);
    if (!(axisLength > 1e-6)) return;
    let normalX = -dy / axisLength;
    let normalY = dx / axisLength;
    const midpointX = (axisStart.x + axisEnd.x) / 2;
    const midpointY = (axisStart.y + axisEnd.y) / 2;
    if ((midpointX - sceneCenter.x) * normalX
      + (midpointY - sceneCenter.y) * normalY < 0) {
      normalX = -normalX;
      normalY = -normalY;
    }
    const length = axisTickLengthPx(level, lineWidth, ptToPx);
    const sideLength = mode === 'cross' ? length / 2 : length;
    const outer = mode === 'out' || mode === 'cross' ? sideLength : 0;
    const inner = mode === 'in' || mode === 'cross' ? sideLength : 0;
    strokeAxisSegment(
      ctx,
      point.x + normalX * outer,
      point.y + normalY * outer,
      point.x - normalX * inner,
      point.y - normalY * inner,
      color,
      lineWidth,
      dash,
    );
  };

  const categoryAxisStart = projection.project(front.x, floorY, nearDepth);
  const categoryAxisEnd = projection.project(front.x + front.w, floorY, nearDepth);
  const surfaceCatLineWidth = chart.catAxisLineWidthEmu != null
    ? axisLineWidthPx(chart.catAxisLineWidthEmu, ptToPx)
    : 1;
  strokeAxisSegment(
    ctx,
    categoryAxisStart.x,
    categoryAxisStart.y,
    categoryAxisEnd.x,
    categoryAxisEnd.y,
    chart.catAxisLineColor ? `#${chart.catAxisLineColor}` : '#000000',
    surfaceCatLineWidth,
    chart.catAxisLineDash,
  );
  const categoryAxisColor = chart.catAxisLineColor
    ? `#${chart.catAxisLineColor}` : '#000000';
  const categoryAxisSuppressed = chart.catAxisHidden || chart.catAxisLineHidden === true;
  const categoryTickSkip = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
  for (let column = 0; column < columnCount; column += categoryTickSkip) {
    drawSurfaceAxisTick(
      chart.catAxisMajorTickMark,
      projection.project(toX(column), floorY, nearDepth),
      categoryAxisStart,
      categoryAxisEnd,
      categoryAxisColor,
      surfaceCatLineWidth,
      categoryAxisSuppressed,
      'major',
      chart.catAxisLineDash,
    );
  }
  if (chart.catAxisMinorTickMark != null && chart.catAxisMinorTickMark !== 'none') {
    for (let column = 0; column < columnCount - 1; column++) {
      const fraction = (
        categoryPositionFraction(column, columnCount, categoryBetween, categoryReversed)
        + categoryPositionFraction(column + 1, columnCount, categoryBetween, categoryReversed)
      ) / 2;
      drawSurfaceAxisTick(
        chart.catAxisMinorTickMark,
        projection.project(front.x + fraction * front.w, floorY, nearDepth),
        categoryAxisStart,
        categoryAxisEnd,
        categoryAxisColor,
        surfaceCatLineWidth,
        categoryAxisSuppressed,
        'minor',
        chart.catAxisLineDash,
      );
    }
  }
  ctx.font = chartFontCss(
    catFontPx,
    chartFontFamily(chart, chart.catAxisFontFace, 'minor'),
    chart.catAxisFontBold ?? false,
    chart.catAxisFontItalic ?? false,
  );
  ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#000000';
  ctx.textBaseline = 'top';
  for (let column = 0; column < columnCount; column++) {
    const anchor = categoryLabelAnchorFraction(
      column,
      columnCount,
      isCrossBetween(chart),
      catAxisReversed(chart),
      chart.catAxisLabelAlignment,
    );
    const point = projection.project(front.x + anchor.fraction * front.w, floorY, nearDepth);
    ctx.textAlign = anchor.textAlign;
    ctx.fillText(
      categories[column] ?? '',
      point.x,
      point.y + categoryLabelOffsetPx(8, chart.catAxisLabelOffsetPercent),
    );
  }

  if (!seriesAxis?.hidden) {
    const seriesAxisWidth = seriesAxis?.lineWidthEmu != null
      ? axisLineWidthPx(seriesAxis.lineWidthEmu, ptToPx)
      : 1;
    const seriesAxisMinPoint = projection.project(front.x, floorY, 0.5);
    const seriesAxisMaxPoint = projection.project(front.x + front.w, floorY, 0.5);
    const seriesAxisX = seriesAxisMinPoint.x >= seriesAxisMaxPoint.x
      ? front.x : front.x + front.w;
    const seriesStart = projection.project(seriesAxisX, floorY, nearDepth);
    const seriesEnd = projection.project(seriesAxisX, floorY, farDepth);
    strokeAxisSegment(
      ctx, seriesStart.x, seriesStart.y, seriesEnd.x, seriesEnd.y,
      seriesAxis?.lineColor ? `#${seriesAxis.lineColor}` : '#000000',
      seriesAxisWidth, seriesAxis?.lineDash,
    );
    const seriesAxisColor = seriesAxis?.lineColor
      ? `#${seriesAxis.lineColor}` : '#000000';
    const seriesTickSkip = Math.max(1, Math.floor(seriesAxis?.tickMarkSkip ?? 1));
    for (let row = 0; row < rowCount; row += seriesTickSkip) {
      drawSurfaceAxisTick(
        seriesAxis?.majorTickMark,
        projection.project(seriesAxisX, floorY, toDepth(row)),
        seriesStart,
        seriesEnd,
        seriesAxisColor,
        seriesAxisWidth,
        seriesAxis?.lineHidden === true,
        'major',
        seriesAxis?.lineDash,
      );
    }
    if (seriesAxis?.minorTickMark != null && seriesAxis.minorTickMark !== 'none') {
      for (let row = 0; row < rowCount - 1; row++) {
        drawSurfaceAxisTick(
          seriesAxis.minorTickMark,
          projection.project(seriesAxisX, floorY, (toDepth(row) + toDepth(row + 1)) / 2),
          seriesStart,
          seriesEnd,
          seriesAxisColor,
          seriesAxisWidth,
          seriesAxis.lineHidden === true,
          'minor',
          seriesAxis.lineDash,
        );
      }
    }
    ctx.font = chartFontCss(
      seriesFontPx,
      chartFontFamily(chart, seriesAxis?.fontFace, 'minor'),
      seriesAxis?.fontBold ?? false,
      seriesAxis?.fontItalic ?? false,
    );
    ctx.fillStyle = seriesAxis?.fontColor ? `#${seriesAxis.fontColor}` : '#000000';
    ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
    for (let row = 0; row < rowCount; row++) {
      const point = projection.project(seriesAxisX, floorY, toDepth(row));
      ctx.fillText(rows[row].name, point.x + 8, point.y);
    }
  }

  if (!chart.valAxisHidden) {
    const valueAxisX = projection.topology.axisX === 'min' ? front.x : front.x + front.w;
    const valueAxisBottom = projection.project(valueAxisX, front.y + front.h, nearDepth);
    const valueAxisTop = projection.project(valueAxisX, front.y, nearDepth);
    if (Math.hypot(valueAxisTop.x - valueAxisBottom.x, valueAxisTop.y - valueAxisBottom.y) > 4) {
      const valueAxisWidth = chart.valAxisLineWidthEmu != null
        ? axisLineWidthPx(chart.valAxisLineWidthEmu, ptToPx)
        : 1;
      strokeAxisSegment(
        ctx, valueAxisBottom.x, valueAxisBottom.y, valueAxisTop.x, valueAxisTop.y,
        chart.valAxisLineColor ? `#${chart.valAxisLineColor}` : '#000000',
        valueAxisWidth, chart.valAxisLineDash,
      );
      const valFontPx = axisLabelPx(chart.valAxisFontSizeHpt, h, ptToPx);
      ctx.font = chartFontCss(
        valFontPx,
        chartFontFamily(chart, chart.valAxisFontFace, 'minor'),
        chart.valAxisFontBold ?? false,
        chart.valAxisFontItalic ?? false,
      );
      ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#000000';
      const left = (valueAxisBottom.x + valueAxisTop.x) / 2 < px0 + pw / 2;
      ctx.textAlign = left ? 'right' : 'left';
      ctx.textBaseline = 'middle';
      for (const value of surfaceMajorLines) {
        const point = projection.project(valueAxisX, toValueY(value), nearDepth);
        ctx.fillText(
          formatChartValWithCode(value, chart.valAxisFormatCode, chart.date1904),
          point.x + (left ? -6 : 6),
          point.y,
        );
      }
    }
  }
  drawLegendForLayout(
    ctx,
    legendChart,
    legend,
    x,
    y,
    w,
    h,
    px0,
    py0,
    pw,
    ph,
    titleBand.bandH + 2,
    ptToPx,
    legendFillPaints,
  );
}
