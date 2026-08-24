import type {
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartDisplayUnits,
  ChartLegendEntryOverride,
  ChartModel,
  ChartRect,
  ChartSeries,
} from '../types/chart.js';
import type { Fill } from '../types/common.js';
import { chartImageFillSource, paintChartImageFill } from './image-fill.js';
import { paintChartThreeDSurfacePicture } from './three-d-surface-picture.js';
import {
  effectiveMarkerSymbol,
  hasVisiblePointMarkerOverride,
  markerFillColorFor,
  markerFillPaintFor,
  markerPaintComponents,
  seriesMarkerFillColor,
  seriesMarkerFillPaint,
  seriesHasMarkerDetail,
} from './marker-style.js';
import { dataLabelIsDeleted } from './data-label-style.js';

/** Stable bundle-audit marker. Re-exported only by the opt-in package entry so
 * build verification can prove the mesh/camera implementation is absent from
 * every base document-format dependency closure. */
export const THREE_D_IMPLEMENTATION_MARKER = 'ooxml-three-d-mesh-implementation-v1';
import { EMU_PER_PT, PT_TO_PX } from '../units.js';
import { pptxPresetDashArray } from '../draw/dash.js';
import { finiteDataExtent, planNumericValueAxis, type NumericValueAxisPlan } from './axis-scale.js';
import { formatCategoryLabel, formatChartValWithCode } from './chart-number-format.js';
import {
  axisTitleFontPx,
  axisTitleMargin,
  axisTitleRotationRad,
  axisTitleVerticalInsetPx,
  cartesianTitleBand,
  chartLegendBands,
  chartLegendReserve,
  chartAxisTitleBands,
  chartTextFontSizePx,
  computeChartFrame,
  packLegendRows,
  resolveManualLayoutRect,
  type ChartAxisTitleSide,
} from './layout.js';
import { elideToWidth } from './text-elide.js';
import { fitDataLabelLines, resolveDataLabelPlacement } from './data-label-layout.js';
import { paintRichDataLabelBlock, resolveRichDataLabelBlock } from './rich-data-label.js';
import { effectiveDataLabelText } from './data-label-content.js';
import { mergeChartLabelBoxes, paintChartLabelBox } from './label-box.js';
import {
  anchoredDataLabelPoint,
  dataLabelCanvasTextAlign,
  dataLabelInsets,
  effectiveDataLabelTextStyle,
  fitStyledDataLabelLines,
  rotatedDataLabelSize,
  transformDataLabelText,
} from './data-label-style.js';
import {
  fitChartThreeDProjectionToPoints,
  fitChartThreeDProjectionToWallThickness,
  pieThreeDThicknessMultiplier,
  planChartThreeDSurfaceGridSegments,
  planChartThreeDSurfaceGeometry,
  planChartThreeDProjection,
  planThreeDBarClusterSlot,
  threeDToMaxScale,
  type ChartThreeDProjection,
  type ThreeDScenePoint,
} from './three-d.js';
import { scaleHexColor } from './material-color.js';
import {
  categoryGridlineFractions,
  categoryMinorGridlineFractions,
  categoryLabelOffsetPx,
  categoryPositionFraction,
} from './category-spacing.js';
import {
  buildThreeDAreaStripMeshes,
  buildThreeDLineRibbonMesh,
  buildThreeDPieSectorMesh,
  buildThreeDShapeMesh,
  DEFAULT_THREE_D_ROUND_SEGMENTS,
  type ThreeDMesh,
} from './three-d-mesh.js';
import {
  clipAxisFractionBand,
  clipAxisFractionSegment,
  sortProjectedSceneFaces,
  type ProjectedSceneFace,
} from './three-d-scene.js';
import {
  buildProjectedStrokeJunction,
  buildProjectedStrokePrimitives,
  MAX_PROJECTED_STROKE_PRIMITIVES,
  type ProjectedStrokePoint,
} from './three-d-stroke.js';
import { buildThreeDOutlineTopology } from './three-d-outline.js';
import { paintLegendFrame } from './legend-frame.js';
import { paintPlotAreaFrame } from './plot-area-frame.js';
import { resolveFill } from '../shape/paint.js';
import {
  chartThreeDSurfacePaint,
  chartStyleFillDecision,
  chartStyleLineDecision,
} from './style-paint.js';
import { drawingmlLineDashArray } from '../draw/dash.js';
import {
  MAX_CHART_PAINT_COMPONENTS,
  MAX_CHART_PAINT_RECIPE_COMPONENTS,
} from './resource-limits.js';

interface ThreeDLegendTextStyle {
  fontPx: number;
  font: string;
  color: string;
}

interface ThreeDLegendMeasure {
  labels: string[];
  styles: ThreeDLegendTextStyle[];
  itemWidths: number[];
}

function threeDLegendOverrideMap(chart: ChartModel): Map<number, ChartLegendEntryOverride> {
  const overrides = new Map<number, ChartLegendEntryOverride>();
  for (const override of chart.legendEntries ?? []) overrides.set(override.idx, override);
  return overrides;
}

function threeDLegendTextStyle(
  chart: ChartModel,
  override: ChartLegendEntryOverride | undefined,
  ptToPx: number,
): ThreeDLegendTextStyle {
  const fontPx = chartTextFontSizePx(
    override?.fontSizeHpt ?? chart.legendFontSizeHpt,
    ptToPx,
  ) ?? 9 * ptToPx;
  const face = override?.fontFace ?? chart.legendFontFace;
  const bold = override?.fontBold ?? chart.legendFontBold ?? false;
  return {
    fontPx,
    font: `${bold ? 'bold ' : ''}${fontPx}px ${fontFamily(face)}`,
    color: override?.fontColor
      ? `#${override.fontColor}`
      : chart.legendFontColor ? `#${chart.legendFontColor}` : '#595959',
  };
}

const PALETTE = ['4472C4', 'ED7D31', '70AD47', 'A5A5A5', 'FFC000', '5B9BD5'] as const;
const MAX_THREE_D_PAINT_RECIPE_COMPONENTS = MAX_CHART_PAINT_RECIPE_COMPONENTS;
const MAX_THREE_D_PAINT_COMPONENTS = MAX_CHART_PAINT_COMPONENTS;
const SUPPORTED_CARTESIAN_THREE_D_TYPES = new Set([
  'line', 'stackedLine', 'stackedLinePct',
  'area', 'stackedArea', 'stackedAreaPct',
  'clusteredBar', 'clusteredBarH',
  'stackedBar', 'stackedBarH', 'stackedBarPct', 'stackedBarHPct',
]);
/** Worst-case cap + side-face count for one revolved 3-D datum. Shared with
 * the renderer preflight so higher fidelity never expands unbounded work. */
export const THREE_D_ROUND_MESH_FACE_WEIGHT = DEFAULT_THREE_D_ROUND_SEGMENTS + 4;

/** Office renders authored 3-D mesh edges in chart/sheet display pixels, not
 * with the 96-DPI typography expansion used for text. The supplied Excel
 * reference has `a:ln@w=12700` (1pt) and a one-pixel solid at 100% zoom across
 * all six equivalent 3-D bar charts. Keep that observed material-edge rule
 * separate from 2-D DrawingML strokes and scale it only with viewer zoom. */
export function threeDMeshOutlineWidthPx(lineWidthEmu: number, ptToPx: number): number {
  const zoom = Number.isFinite(ptToPx) && ptToPx > 0 ? ptToPx / PT_TO_PX : 1;
  const points = Number.isFinite(lineWidthEmu) && lineWidthEmu >= 0
    ? lineWidthEmu / EMU_PER_PT : 0;
  return Math.max(0.25, points) * zoom;
}

/** Group stack geometry in one pass before resolving shared internal caps.
 * The parser/render preflight bounds the item count, but rescanning the whole
 * primitive array once per category still turns a valid wide chart into
 * quadratic CPU work. Invalid public-model category indexes are ignored. */
export function bucketThreeDStackItems<T extends { readonly categoryIndex: number }>(
  items: readonly T[],
  categoryCount: number,
): T[][] {
  const boundedCount = Number.isSafeInteger(categoryCount) && categoryCount > 0
    ? categoryCount : 0;
  const buckets = Array.from({ length: boundedCount }, () => [] as T[]);
  for (const item of items) {
    const categoryIndex = item.categoryIndex;
    if (Number.isSafeInteger(categoryIndex)
      && categoryIndex >= 0
      && categoryIndex < boundedCount) {
      buckets[categoryIndex].push(item);
    }
  }
  return buckets;
}

interface Point { x: number; y: number }

interface SceneFace {
  points: readonly Point[];
  color: string;
  /** One already-resolved paint shared by every face of the same datum.
   * `null` is authored noFill/unresolved and suppresses the automatic solid. */
  paint?: string | CanvasGradient | CanvasPattern | null;
  paintRole: 'fill' | 'outline';
  shade: number;
  cameraDepth: number;
  cameraDepths?: readonly number[];
  cameraWeights?: readonly number[];
  outline: boolean;
  outlineColor?: string;
  outlineWidth?: number;
  outlineDash?: string;
  outlineCap?: CanvasLineCap;
  outlineJoin?: CanvasLineJoin;
}

interface MeshOutlineStyle {
  readonly color: string;
  readonly width: number;
  readonly dash: readonly number[];
  readonly cap: CanvasLineCap;
  readonly join: CanvasLineJoin;
}

interface ScenePrimitiveBudget {
  remaining: number;
  exceeded: boolean;
}

const colorFor = (index: number, series?: ChartSeries): string =>
  `#${series?.color ?? PALETTE[index % PALETTE.length]}`;

interface ThreeDDatumPaint {
  color: string;
  fill: Fill | null | undefined;
  lineColor: string | null;
  lineFill: ChartModel['plotAreaLineFill'] | null | undefined;
  lineWidthEmu: number | null | undefined;
  lineDash: string | null | undefined;
  lineCustomDash: ChartModel['plotAreaLineCustomDash'];
  lineCap: CanvasLineCap;
  lineJoin: CanvasLineJoin;
}

function threeDPaintsWithinBudget<T>(
  items: readonly T[],
  fillFor: (item: T) => Fill | null | undefined,
  lineFor: (item: T) => ChartModel['plotAreaLineFill'] | null | undefined,
): boolean {
  let total = 0;
  for (const item of items) {
    for (const recipe of [fillFor(item), lineFor(item)]) {
      if (recipe == null) continue;
      const components = markerPaintComponents(recipe);
      if ((recipe.fillType === 'gradient'
          && components > MAX_THREE_D_PAINT_RECIPE_COMPONENTS)
        || components > MAX_THREE_D_PAINT_COMPONENTS - total) return false;
      total += components;
    }
  }
  return true;
}

interface ThreeDPercentCategoryScale {
  maxMagnitude: number;
  scaledTotal: number;
}

/** Plan percent denominators once per category in O(series × categories).
 * Dividing by the category maximum first keeps finite Number.MAX_VALUE inputs
 * proportional without overflowing their absolute-value sum. */
function planThreeDPercentCategoryScales(
  chart: ChartModel,
  categoryCount: number,
): ThreeDPercentCategoryScale[] {
  const scales = Array.from({ length: categoryCount }, () => ({
    maxMagnitude: 0,
    scaledTotal: 0,
  }));
  for (const series of chart.series) {
    for (let index = 0; index < categoryCount; index++) {
      const candidate = series.values[index];
      if (candidate != null && Number.isFinite(candidate)) {
        scales[index].maxMagnitude = Math.max(
          scales[index].maxMagnitude,
          Math.abs(candidate),
        );
      }
    }
  }
  for (const series of chart.series) {
    for (let index = 0; index < categoryCount; index++) {
      const candidate = series.values[index];
      const maxMagnitude = scales[index].maxMagnitude;
      if (candidate != null && Number.isFinite(candidate) && maxMagnitude > 0) {
        scales[index].scaledTotal += Math.abs(candidate) / maxMagnitude;
      }
    }
  }
  return scales;
}

/** Convert one value into the plotted percent-stack domain from the bounded
 * category plan shared by preflight, extent planning, and geometry paint. */
function threeDPlottedValue(
  chart: ChartModel,
  seriesIndex: number,
  categoryIndex: number,
  percentScales: readonly ThreeDPercentCategoryScale[] | undefined,
): number {
  const raw = chart.series[seriesIndex]?.values[categoryIndex] ?? 0;
  if (!Number.isFinite(raw)) return 0;
  if (!percentScales) return raw;
  const { maxMagnitude, scaledTotal } = percentScales[categoryIndex]
    ?? { maxMagnitude: 0, scaledTotal: 0 };
  if (!(maxMagnitude > 0)) return 0;
  return scaledTotal > 0 ? raw / maxMagnitude / scaledTotal * 100 : 0;
}

function threeDSafeAdd(left: number, right: number): number {
  const sum = left + right;
  if (Number.isFinite(sum)) return sum;
  return right < 0 ? -Number.MAX_VALUE : Number.MAX_VALUE;
}

/** Preflight every structured recipe that the selected optional 3-D family
 * can consume before title, plot, wall, or datum paint begins. This keeps a
 * public model or linked style with an oversized gradient atomic instead of
 * leaving a partially painted chart behind. */
function threeDChartPaintsWithinBudget(chart: ChartModel): boolean {
  const paints: Array<{
    fill: Fill | null | undefined;
    line: ChartModel['plotAreaLineFill'] | null | undefined;
  }> = [];
  for (const [surface, role] of [
    [chart.threeD?.floor, 'floor'],
    [chart.threeD?.sideWall, 'wall'],
    [chart.threeD?.backWall, 'wall'],
  ] as const) {
    const paint = chartThreeDSurfacePaint(chart, surface, role);
    paints.push({ fill: paint.fill, line: paint.line });
  }

  if (chart.chartType === 'pie') {
    const series = chart.series[0];
    if (series) {
      const overrides = new Map(
        series.dataPointOverrides?.map(point => [point.idx, point]) ?? [],
      );
      for (let index = 0; index < series.values.length; index++) {
        const value = series.values[index];
        if (value == null || !Number.isFinite(value) || Math.abs(value) <= 0) continue;
        const paint = threeDDatumPaint(chart, series, overrides.get(index), index, 0);
        paints.push({ fill: paint.fill, line: paint.lineFill });
      }
    }
  } else {
    const bars = chart.chartType === 'clusteredBar'
      || chart.chartType === 'clusteredBarH'
      || chart.chartType.startsWith('stackedBar');
    if (bars) {
      const stacked = chart.chartType.startsWith('stacked');
      const percent = chart.chartType.endsWith('Pct');
      const zeroBlanks = stacked || chart.dispBlanksAs === 'zero';
      const logarithmic = chart.valAxisLogBase != null
        && Number.isFinite(chart.valAxisLogBase) && chart.valAxisLogBase >= 2;
      const categoryCount = Math.max(
        1,
        chart.categories.length,
        ...chart.series.map(series => Math.max(
          series.values.length,
          series.categories?.length ?? 0,
        )),
      );
      const percentScales = percent
        ? planThreeDPercentCategoryScales(chart, categoryCount)
        : undefined;
      const visibleInterval = (base: number, end: number): boolean => {
        if (base === end) return false;
        const lower = Math.min(base, end);
        const upper = Math.max(base, end);
        const min = chart.valMin;
        const max = chart.valMax;
        if (min != null && Number.isFinite(min) && upper <= min) return false;
        if (max != null && Number.isFinite(max) && lower >= max) return false;
        return !logarithmic || upper > 0;
      };
      const positiveBase = new Array(categoryCount).fill(0) as number[];
      const negativeBase = new Array(categoryCount).fill(0) as number[];
      for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
        const series = chart.series[seriesIndex];
        const overrides = new Map(
          series.dataPointOverrides?.map(point => [point.idx, point]) ?? [],
        );
        for (let index = 0; index < categoryCount; index++) {
          const value = series.values[index];
          const visible = value != null && Number.isFinite(value)
            && (!logarithmic || value > 0)
            || (value == null && zeroBlanks);
          if (!visible) continue;
          const plottedValue = threeDPlottedValue(chart, seriesIndex, index, percentScales);
          const base = stacked
            ? (plottedValue >= 0 ? positiveBase[index] : negativeBase[index])
            : 0;
          const end = threeDSafeAdd(base, plottedValue);
          if (stacked) {
            if (plottedValue >= 0) positiveBase[index] = end;
            else negativeBase[index] = end;
          }
          if (!visibleInterval(base, end)) continue;
          const paint = threeDDatumPaint(
            chart, series, overrides.get(index), index, seriesIndex,
          );
          paints.push({ fill: paint.fill, line: paint.lineFill });
        }
      }
    } else {
      const area = chart.chartType.toLowerCase().includes('area');
      const stacked = chart.chartType.startsWith('stacked');
      const percent = chart.chartType.endsWith('Pct');
      const zeroBlanks = stacked || chart.dispBlanksAs === 'zero';
      const logarithmic = chart.valAxisLogBase != null
        && Number.isFinite(chart.valAxisLogBase) && chart.valAxisLogBase >= 2;
      const categoryCount = Math.max(
        chart.categories.length,
        ...chart.series.map(series => Math.max(
          series.categories?.length ?? 0,
          series.values.length,
        )),
      );
      const percentScales = percent
        ? planThreeDPercentCategoryScales(chart, categoryCount)
        : undefined;
      const stackedUpper = chart.series.map(() => new Array<number>(categoryCount).fill(0));
      if (stacked) {
        const positive = new Array<number>(categoryCount).fill(0);
        const negative = new Array<number>(categoryCount).fill(0);
        for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
          for (let index = 0; index < categoryCount; index++) {
            const value = threeDPlottedValue(chart, seriesIndex, index, percentScales);
            const base = value >= 0 ? positive[index] : negative[index];
            const end = threeDSafeAdd(base, value);
            stackedUpper[seriesIndex][index] = end;
            if (value >= 0) positive[index] = end;
            else negative[index] = end;
          }
        }
      }
      for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
        const series = chart.series[seriesIndex];
        const overrides = new Map(
          series.dataPointOverrides?.map(point => [point.idx, point]) ?? [],
        );
        const pointCount = categoryCount;
        let runLength = 0;
        let previousValue: number | null = null;
        let hasGeometry = false;
        let defaultLinePaintUsed = false;
        for (let index = 0; index < pointCount; index++) {
          const value = series.values[index];
          const visible = value != null && Number.isFinite(value)
            && (!logarithmic || value > 0)
            || (value == null && zeroBlanks);
          if (visible) {
            runLength++;
            const currentValue = stacked
              ? stackedUpper[seriesIndex][index]
              : threeDPlottedValue(chart, seriesIndex, index, percentScales);
            const lower = previousValue == null ? currentValue : Math.min(previousValue, currentValue);
            const upper = previousValue == null ? currentValue : Math.max(previousValue, currentValue);
            const outsideAuthoredAxis = chart.valMin != null && Number.isFinite(chart.valMin)
              && upper <= chart.valMin
              || chart.valMax != null && Number.isFinite(chart.valMax) && lower >= chart.valMax;
            if (runLength >= 2 && (area || series.smooth === true || !outsideAuthoredAxis)) {
              hasGeometry = true;
              if (area) break;
              const point = overrides.get(index);
              if (threeDPointOwnsLineSegment(point)) {
                const paint = threeDDatumPaint(chart, series, point, index, seriesIndex);
                paints.push({ fill: undefined, line: paint.lineFill });
              } else {
                defaultLinePaintUsed = true;
              }
            }
            previousValue = currentValue;
          } else if ((chart.dispBlanksAs ?? 'gap') === 'gap') {
            runLength = 0;
            previousValue = null;
          }
        }
        if (!hasGeometry) continue;
        const paint = threeDDatumPaint(chart, series, undefined, seriesIndex, seriesIndex);
        if (area) paints.push({ fill: paint.fill, line: paint.lineFill });
        else if (defaultLinePaintUsed) paints.push({ fill: undefined, line: paint.lineFill });
      }
    }
  }
  return threeDPaintsWithinBudget(paints, paint => paint.fill, paint => paint.line);
}

function solidOrStructured(
  decision: Fill | null | undefined,
  fallbackColor: string,
): { color: string; fill: Fill | null | undefined } {
  if (decision?.fillType === 'solid') {
    return { color: `#${decision.color}`, fill: undefined };
  }
  return { color: fallbackColor, fill: decision };
}

/** Resolve one classic 3-D datum without materializing per-face style state.
 * Direct point > direct series > linked dataPoint3D > automatic. Solid colors
 * continue through the bounded automatic material; structured recipes are
 * resolved once for the complete datum after projection. */
function threeDDatumPaint(
  chart: ChartModel,
  series: ChartSeries,
  point: ChartDataPointOverride | undefined,
  pointIndex: number,
  seriesIndex: number,
): ThreeDDatumPaint {
  // `ChartSeries.color` and `dataPointColors` are effective legacy colors:
  // the parser may materialize the automatic accent palette there. Keep them
  // as the final automatic fallback so an authored/linked 3-D paint recipe can
  // still own the mark. Direct spPr provenance is carried by chartexStyle.
  const automaticColor = series.dataPointColors?.[pointIndex]
    ? `#${series.dataPointColors[pointIndex]}`
    : colorFor(seriesIndex, series);
  // Chart Style palettes normally advance by series. Pie slices, and the
  // single-series bar case explicitly carrying varyColors, advance by point.
  // Keep this independent from `pointIndex`: that index still selects direct
  // dPt formatting, while linked/series roles describe the owning series.
  const styleIndex = chart.chartType === 'pie'
    || (chart.varyColors === true
      && chart.series.length === 1
      && chart.chartType.includes('Bar'))
    ? pointIndex
    : seriesIndex;
  let fill: Fill | null | undefined;
  let color = automaticColor;
  const pointDecision = chartStyleFillDecision(point?.chartexStyle, pointIndex);
  if (pointDecision !== undefined) {
    ({ color, fill } = solidOrStructured(pointDecision, color));
  } else if (point?.fillHidden === true || point?.color === '00000000') {
    fill = null;
  } else if (point?.color) {
    color = `#${point.color}`;
  } else {
    const seriesDecision = chartStyleFillDecision(series.chartexStyle, styleIndex);
    if (seriesDecision !== undefined) {
      ({ color, fill } = solidOrStructured(seriesDecision, color));
    } else {
      const linkedDecision = chartStyleFillDecision(
        chart.chartStyleRoles?.dataPoint3D, styleIndex,
      );
      if (linkedDecision !== undefined) {
        ({ color, fill } = solidOrStructured(linkedDecision, color));
      }
    }
  }

  let lineFill: ChartModel['plotAreaLineFill'] | null | undefined;
  let lineColor: string | null = null;
  const pointLine = chartStyleLineDecision(point?.chartexStyle, pointIndex);
  if (pointLine !== undefined) {
    if (pointLine?.fillType === 'solid') lineColor = `#${pointLine.color}`;
    else lineFill = pointLine;
  } else if (point?.lineHidden === true) lineFill = null;
  else if (point?.lineColor) lineColor = `#${point.lineColor}`;
  else {
    const seriesLine = chartStyleLineDecision(series.chartexStyle, styleIndex);
    if (seriesLine !== undefined) {
      if (seriesLine?.fillType === 'solid') lineColor = `#${seriesLine.color}`;
      else lineFill = seriesLine;
    } else if (series.lineHidden === true) lineFill = null;
    else if (series.lineColor) lineColor = `#${series.lineColor}`;
    else {
      const linkedLine = chartStyleLineDecision(
        chart.chartStyleRoles?.dataPoint3D, styleIndex,
      );
      if (linkedLine !== undefined) {
        if (linkedLine?.fillType === 'solid') lineColor = `#${linkedLine.color}`;
        else lineFill = linkedLine;
      }
    }
  }
  const linkedGeometry = chart.chartStyleRoles?.dataPoint3D?.lineNoStyle === true
    ? undefined : chart.chartStyleRoles?.dataPoint3D;
  const lineCapValue = point?.chartexStyle?.lineCap
    ?? series.chartexStyle?.lineCap ?? linkedGeometry?.lineCap;
  const lineJoinValue = point?.chartexStyle?.lineJoin
    ?? series.chartexStyle?.lineJoin ?? linkedGeometry?.lineJoin;
  return {
    color,
    fill,
    lineColor,
    lineFill,
    lineWidthEmu: point?.lineWidthEmu ?? point?.chartexStyle?.lineWidthEmu
      ?? series.lineWidthEmu ?? series.chartexStyle?.lineWidthEmu
      ?? linkedGeometry?.lineWidthEmu,
    lineDash: point?.lineDash ?? point?.chartexStyle?.lineDash
      ?? series.chartexStyle?.lineDash ?? linkedGeometry?.lineDash,
    lineCustomDash: point?.chartexStyle?.lineCustomDash
      ?? series.chartexStyle?.lineCustomDash ?? linkedGeometry?.lineCustomDash,
    lineCap: lineCapValue === 'rnd' ? 'round' : lineCapValue === 'sq' ? 'square' : 'butt',
    lineJoin: lineJoinValue === 'round' || lineJoinValue === 'bevel'
      ? lineJoinValue : 'miter',
  };
}

/** A 3-D line dPt owns only the incoming segment's line component. Shape fill
 * remains irrelevant to Excel's line ribbon, and area dPt paint is not applied
 * to the series body. Keep this predicate narrow so color-only fill overrides
 * do not split or recolor line geometry. */
function threeDPointOwnsLineSegment(
  point: ChartDataPointOverride | undefined,
): boolean {
  if (!point) return false;
  const style = point.chartexStyle;
  return point.lineHidden != null
    || point.lineColor != null
    || point.lineWidthEmu != null
    || point.lineDash != null
    || style?.linePaintAuthored === true
    || style?.lineHidden != null
    || style?.lineNoStyle === true
    || style?.lineColors?.some(color => color != null) === true
    || style?.linePaints?.some(paint => paint != null) === true
    || style?.lineWidthEmu != null
    || style?.lineDash != null
    || style?.lineCustomDash != null
    || style?.lineCap != null
    || style?.lineJoin != null;
}

/** Automatic 3-D chart material evaluated from a real camera-space face
 * normal. This stage changes color only; it never moves or invents geometry. */
function meshMaterialFactor(
  projection: ChartThreeDProjection,
  points: readonly ThreeDScenePoint[],
): number {
  const normal = projection.cameraNormal(points);
  if (!normal) return 1;
  // Office's automatic material behaves primarily like a camera-side soft
  // light: the centre of a revolved face remains close to the authored color,
  // while the silhouette shoulder darkens. A small upper-left component keeps
  // planar caps distinguishable without turning side faces muddy.
  const light = { x: -0.20, y: 0.25, z: 1 };
  const lightLength = Math.hypot(light.x, light.y, light.z);
  const lambert = Math.max(0,
    (normal.x * light.x + normal.y * light.y + normal.z * light.z) / lightLength,
  );
  // Office's automatic 3-D material keeps the illuminated band near the
  // authored accent and lowers visible shoulders to roughly 78%. The former
  // 62% ambient term over-darkened every real mesh even though its normals and
  // occlusion were correct. Keep one directional Lambert response, but raise
  // the ambient contribution instead of applying family-specific fill hacks.
  return Math.max(0.78, Math.min(1, 0.78 + 0.24 * lambert));
}

const fontFamily = (face: string | null | undefined): string => {
  const safe = face && !face.startsWith('+') ? face.replace(/["\\]/g, '') : 'Arial';
  return `"${safe}"`;
};

const chartFontFamily = (
  chart: ChartModel,
  face: string | null | undefined,
  role: 'major' | 'minor' = 'minor',
): string => {
  const reference = face?.startsWith('+mj-')
    ? chart.themeMajorFontLatin
    : face?.startsWith('+mn-') ? chart.themeMinorFontLatin : face;
  const resolved = reference ?? (role === 'major' ? chart.themeMajorFontLatin : chart.themeMinorFontLatin);
  return fontFamily(resolved);
};

function polygon(ctx: CanvasRenderingContext2D, points: readonly Point[]): void {
  if (!points.length) return;
  ctx.beginPath();
  ctx.moveTo(points[0].x, points[0].y);
  for (let index = 1; index < points.length; index++) ctx.lineTo(points[index].x, points[index].y);
  ctx.closePath();
}

const isTransparentPaint = (color: string): boolean =>
  color === 'transparent' || color === '#00000000' || color === 'rgba(0,0,0,0)';

function face(
  ctx: CanvasRenderingContext2D,
  points: readonly Point[],
  color: string,
  shade: number,
): void {
  if (isTransparentPaint(color)) return;
  polygon(ctx, points);
  ctx.fillStyle = color;
  ctx.fill();
  if (shade > 0) {
    polygon(ctx, points);
    ctx.fillStyle = `rgba(0,0,0,${shade})`;
    ctx.fill();
  }
}

/** Build an extruded area strip from one adjacent category interval.
 *
 * Excel's 3-D area is a solid ribbon, not a zero-thickness polygon.  Both
 * broad faces, the value ridge and the datum edge are projected from the same
 * model-space slab as bars, walls and axes.  Internal category seams are not
 * emitted here; callers add only the two run-end caps so adjacent strips form
 * one continuous solid.
 */
function areaStripFaces(
  projection: ChartThreeDProjection,
  x0: number,
  x1: number,
  lower0: number,
  lower1: number,
  upper0: number,
  upper1: number,
  nearDepth: number,
  farDepth: number,
  color: string,
  capStart: boolean,
  capEnd: boolean,
): SceneFace[] {
  return buildThreeDAreaStripMeshes({
    x0, x1, lower0, lower1, upper0, upper1,
    nearDepth, farDepth, capStart, capEnd,
  }).flatMap(mesh => projectThreeDMesh(projection, mesh, color)
    .map(projected => ({ ...projected, outline: false, outlineSegments: undefined })));
}

function shapeMeshFaces(
  projection: ChartThreeDProjection,
  shape: string,
  horizontal: boolean,
  x: number,
  y: number,
  width: number,
  height: number,
  baseCoord: number,
  endCoord: number,
  nearDepth: number,
  farDepth: number,
  color: string,
  baseScale: number,
  endScale: number,
  omitBaseCap = false,
  omitEndCap = false,
  outlineStyle?: MeshOutlineStyle,
  budget?: ScenePrimitiveBudget,
): SceneFace[] {
  if (![x, y, width, height, baseCoord, endCoord, nearDepth, farDepth].every(Number.isFinite)
    || width <= 0 || height <= 0 || baseCoord === endCoord) return [];
  const mesh = buildThreeDShapeMesh({
    shape,
    horizontal,
    crossStart: horizontal ? y : x,
    crossSize: horizontal ? height : width,
    baseCoord,
    endCoord,
    nearDepth,
    farDepth,
    baseScale,
    endScale,
    omitBaseCap,
    omitEndCap,
  });
  if (!mesh) return [];
  return projectThreeDMesh(projection, mesh, color, outlineStyle, budget);
}

/** Project one model-space solid without changing its topology.
 *
 * Mesh construction owns vertices, faces and outward winding. This stage owns
 * only camera visibility, depth and material response; Canvas paint receives
 * already projected faces. Keeping these stages separate prevents a shape
 * from being approximated by screen-space rectangles or ad-hoc fill bands.
 */
function projectThreeDMesh(
  projection: ChartThreeDProjection,
  mesh: ThreeDMesh,
  color: string,
  outlineStyle?: MeshOutlineStyle,
  budget?: ScenePrimitiveBudget,
  outlineOnly = false,
): SceneFace[] {
  if (budget?.exceeded) return [];
  const candidates = mesh.faces.map(meshFace => {
    const scenePoints = meshFace.indices.map(index => mesh.vertices[index]);
    const facing = projection.cameraFacing(scenePoints);
    if (!facing) return { meshFace, facing, face: null as SceneFace | null };
    const points = scenePoints.map(point => projection.project(point.x, point.y, point.depth));
    const twiceArea = points.reduce((sum, point, index) => {
      const next = points[(index + 1) % points.length];
      return sum + point.x * next.y - point.y * next.x;
    }, 0);
    if (!Number.isFinite(twiceArea) || Math.abs(twiceArea) < 1e-7) {
      return { meshFace, facing, face: null as SceneFace | null };
    }
    const face: SceneFace = {
      points,
      color: scaleHexColor(color, meshMaterialFactor(projection, scenePoints)),
      paintRole: 'fill',
      shade: 0,
      cameraDepth: scenePoints.reduce(
        (sum, point) => sum + projection.cameraDepth(point.x, point.y, point.depth),
        0,
      ) / scenePoints.length,
      cameraDepths: scenePoints.map(point =>
        projection.cameraDepth(point.x, point.y, point.depth)),
      cameraWeights: scenePoints.map(point =>
        projection.cameraProjectionWeight(point.x, point.y, point.depth)),
      // Authored outlines are derived once from mesh topology below. Painting
      // each face perimeter would double shared edges and restart dash phase.
      outline: false,
    };
    return { meshFace, facing, face };
  });
  const visibleFaces = outlineOnly ? [] : candidates
    .map(candidate => candidate.face)
    .filter((face): face is SceneFace => face != null);
  if (budget) {
    if (visibleFaces.length > budget.remaining) {
      budget.exceeded = true;
      return [];
    }
    budget.remaining -= visibleFaces.length;
  }
  const outlineEdges = new Map<string, readonly [number, number]>();
  const addOutlineEdge = (startIndex: number, endIndex: number) => {
    if (startIndex === endIndex) return;
    const key = startIndex < endIndex
      ? `${startIndex}:${endIndex}` : `${endIndex}:${startIndex}`;
    if (!outlineEdges.has(key)) outlineEdges.set(key, [startIndex, endIndex]);
  };
  // Polygonal faces and smooth caps have real boundary edges. Add them once
  // for the complete visible mesh, never once per face.
  for (const candidate of candidates) {
    if (!candidate.facing || !candidate.face) continue;
    const { indices, smoothSurface, role } = candidate.meshFace;
    if (smoothSurface && role === 'side') continue;
    for (let index = 0; index < indices.length; index++) {
      addOutlineEdge(indices[index], indices[(index + 1) % indices.length]);
    }
  }

  // A smooth mesh has no authored facet seams. If an outline is authored,
  // derive only its longitudinal silhouette from visibility transitions in
  // the actual side-face topology.
  const sideCandidates = candidates
    .filter(candidate => candidate.meshFace.role === 'side')
    .sort((a, b) => (a.meshFace.segmentIndex ?? 0) - (b.meshFace.segmentIndex ?? 0));
  if (mesh.silhouetteEdges.length === sideCandidates.length && sideCandidates.length > 0) {
    for (let index = 0; index < sideCandidates.length; index++) {
      const previous = sideCandidates[(index + sideCandidates.length - 1) % sideCandidates.length];
      const current = sideCandidates[index];
      if (previous.facing === current.facing) continue;
      const visible = current.facing ? current.face : previous.face;
      if (!visible) continue;
      const [startIndex, endIndex] = mesh.silhouetteEdges[index];
      addOutlineEdge(startIndex, endIndex);
    }
  }
  const baseCapFacing = candidates.some(candidate =>
    candidate.meshFace.role === 'baseCap' && candidate.facing && candidate.face != null);
  const endCapFacing = candidates.some(candidate =>
    candidate.meshFace.role === 'endCap' && candidate.facing && candidate.face != null);
  for (const candidate of sideCandidates) {
    if (!candidate.facing || !candidate.face) continue;
    const rimEdges = [
      !baseCapFacing ? candidate.meshFace.baseRimEdge : undefined,
      !endCapFacing ? candidate.meshFace.endRimEdge : undefined,
    ].filter((edge): edge is readonly [number, number] => edge != null);
    for (const [startIndex, endIndex] of rimEdges) {
      addOutlineEdge(startIndex, endIndex);
    }
  }
  const topology = buildThreeDOutlineTopology([...outlineEdges.values()].map(([startIndex, endIndex]) => [
    mesh.vertices[startIndex], mesh.vertices[endIndex],
  ]));
  if (!outlineStyle) return visibleFaces;
  const sameScenePoint = (left: ThreeDScenePoint, right: ThreeDScenePoint) =>
    Math.hypot(left.x - right.x, left.y - right.y, left.depth - right.depth) <= 1e-9;
  const isJunction = (point: ThreeDScenePoint) =>
    topology.junctions.some(junction => sameScenePoint(junction.point, point));
  for (const scenePath of topology.paths) {
    const projectedPath = scenePath.map(point => projection.project(point.x, point.y, point.depth));
    const cameraDepths = scenePath.map(point =>
      projection.cameraDepth(point.x, point.y, point.depth));
    const cameraWeights = scenePath.map(point =>
      projection.cameraProjectionWeight(point.x, point.y, point.depth));
    const stroke = buildProjectedStrokePrimitives(projectedPath.map((point, index) => ({
      ...point,
      cameraDepth: cameraDepths[index],
      cameraWeight: cameraWeights[index],
    })), {
      width: outlineStyle.width,
      dash: outlineStyle.dash,
      lineCap: outlineStyle.cap,
      startCap: isJunction(scenePath[0]) ? 'butt' : outlineStyle.cap,
      endCap: isJunction(scenePath.at(-1)!) ? 'butt' : outlineStyle.cap,
      overlapStart: isJunction(scenePath[0]),
      overlapEnd: isJunction(scenePath.at(-1)!),
      lineJoin: outlineStyle.join,
    });
    if (!stroke || (budget && stroke.length > budget.remaining)) {
      if (budget) budget.exceeded = true;
      return [];
    }
    if (budget) budget.remaining -= stroke.length;
    visibleFaces.push(...stroke.map(primitive => {
      // Mesh edges are coplanar with their adjacent fill. A minute camera-space
      // bias prevents a later coplanar slice from erasing one side of an
      // authored circumference; genuinely nearer solids still occlude it.
      const bias = 1e-6 * Math.max(1, Math.abs(primitive.cameraDepth));
      return {
        points: primitive.points,
        color: outlineStyle.color,
        paintRole: 'outline' as const,
        shade: 0,
        cameraDepth: primitive.cameraDepth + bias,
        cameraDepths: primitive.cameraDepths?.map(depth => depth + bias),
        cameraWeights: primitive.cameraWeights,
        outline: false,
      };
    }));
  }
  for (const junction of topology.junctions) {
    const centerPoint = projection.project(
      junction.point.x, junction.point.y, junction.point.depth,
    );
    const center: ProjectedStrokePoint = {
      ...centerPoint,
      cameraDepth: projection.cameraDepth(
        junction.point.x, junction.point.y, junction.point.depth,
      ),
      cameraWeight: projection.cameraProjectionWeight(
        junction.point.x, junction.point.y, junction.point.depth,
      ),
    };
    const neighbours: ProjectedStrokePoint[] = junction.neighbours.map(point => ({
      ...projection.project(point.x, point.y, point.depth),
      cameraDepth: projection.cameraDepth(point.x, point.y, point.depth),
      cameraWeight: projection.cameraProjectionWeight(point.x, point.y, point.depth),
    }));
    const primitive = buildProjectedStrokeJunction(center, neighbours, {
      width: outlineStyle.width,
      lineJoin: outlineStyle.join,
    });
    if (!primitive) continue;
    if (budget && budget.remaining < 1) {
      budget.exceeded = true;
      return [];
    }
    if (budget) budget.remaining -= 1;
    const bias = 1e-6 * Math.max(1, Math.abs(primitive.cameraDepth));
    visibleFaces.push({
      points: primitive.points,
      color: outlineStyle.color,
      paintRole: 'outline',
      shade: 0,
      cameraDepth: primitive.cameraDepth + bias,
      cameraDepths: primitive.cameraDepths?.map(depth => depth + bias),
      cameraWeights: primitive.cameraWeights,
      outline: false,
    });
  }
  return visibleFaces;
}

interface ThreeDPieOutlineSlice {
  readonly start: number;
  readonly end: number;
  readonly segments: number;
  readonly centerX: number;
  readonly centerDepth: number;
}

/** Build a pie outline from semantic curves instead of the generic mesh-edge
 * graph. A pie has four meaningful stroke families: the visible cap rim,
 * radial separators, the visible lower wall rim, and exposed outer vertical
 * separators. Keeping them as independent continuous paths avoids high-degree
 * graph junctions that create spikes and locally doubled widths. */
function projectThreeDPieOutline(
  projection: ChartThreeDProjection,
  slices: readonly ThreeDPieOutlineSlice[],
  centerY: number,
  radius: number,
  thickness: number,
  style: MeshOutlineStyle,
  budget: ScenePrimitiveBudget,
): SceneFace[] {
  if (!slices.length || budget.exceeded) return [];
  // `threeDPieSliceAngles` stores each clockwise slice as an ascending model
  // interval, so source order runs from high angles toward low angles. Curved
  // paths must be traversed in geometric angle order; concatenating source
  // order inserts long chords between slices and paints those chords again as
  // radial separators—the severe doubled lines seen in enlarged pies.
  const orderedSlices = [...slices].sort((left, right) => left.start - right.start);
  const topY = centerY - thickness / 2;
  const bottomY = centerY + thickness / 2;
  const reference = orderedSlices[0];
  const capY = projection.cameraDepth(reference.centerX, topY, reference.centerDepth)
    >= projection.cameraDepth(reference.centerX, bottomY, reference.centerDepth)
    ? topY : bottomY;
  const lowerY = capY === topY ? bottomY : topY;
  const pointAt = (slice: ThreeDPieOutlineSlice, angle: number, y: number): ThreeDScenePoint => ({
    x: slice.centerX + Math.cos(angle) * radius,
    y,
    depth: slice.centerDepth + Math.sin(angle) * radius / projection.modelDepth,
  });
  const centerAt = (slice: ThreeDPieOutlineSlice): ThreeDScenePoint => ({
    x: slice.centerX, y: capY, depth: slice.centerDepth,
  });
  const result: SceneFace[] = [];
  const addPath = (scenePath: readonly ThreeDScenePoint[], closed = false): void => {
    if (scenePath.length < 2 || budget.exceeded) return;
    const first = scenePath[0];
    const last = scenePath.at(-1)!;
    const alreadyClosed = Math.hypot(
      first.x - last.x, first.y - last.y, first.depth - last.depth,
    ) <= 1e-9;
    const points = closed && !alreadyClosed ? [...scenePath, first] : [...scenePath];
    const projected: ProjectedStrokePoint[] = points.map(point => ({
      ...projection.project(point.x, point.y, point.depth),
      cameraDepth: projection.cameraDepth(point.x, point.y, point.depth),
      cameraWeight: projection.cameraProjectionWeight(point.x, point.y, point.depth),
    }));
    const primitives = buildProjectedStrokePrimitives(projected, {
      width: style.width,
      dash: style.dash,
      lineCap: style.cap,
      startCap: 'butt',
      endCap: 'butt',
      lineJoin: style.join,
    });
    if (!primitives || primitives.length > budget.remaining) {
      budget.exceeded = true;
      return;
    }
    budget.remaining -= primitives.length;
    for (const primitive of primitives) {
      const bias = 1e-6 * Math.max(1, Math.abs(primitive.cameraDepth));
      result.push({
        points: primitive.points,
        color: style.color,
        paintRole: 'outline',
        shade: 0,
        cameraDepth: primitive.cameraDepth + bias,
        cameraDepths: primitive.cameraDepths?.map(depth => depth + bias),
        cameraWeights: primitive.cameraWeights,
        outline: false,
      });
    }
  };

  const capRim: ThreeDScenePoint[] = [];
  const lowerRuns: ThreeDScenePoint[][] = [];
  let lowerRun: ThreeDScenePoint[] = [];
  const visibleBoundaryAngles = new Set<string>();
  const angleKey = (angle: number) => {
    const turn = Math.PI * 2;
    const normalized = ((angle % turn) + turn) % turn;
    return (normalized < 1e-8 || turn - normalized < 1e-8 ? 0 : normalized).toFixed(8);
  };
  const boundarySlices = new Map<string, { slice: ThreeDPieOutlineSlice; angle: number }>();
  for (const slice of orderedSlices) {
    boundarySlices.set(angleKey(slice.start), { slice, angle: slice.start });
    boundarySlices.set(angleKey(slice.end), { slice, angle: slice.end });
    for (let segment = 0; segment <= slice.segments; segment++) {
      const angle = slice.start + (slice.end - slice.start) * segment / slice.segments;
      const capPoint = pointAt(slice, angle, capY);
      if (!capRim.length || Math.hypot(
        capRim.at(-1)!.x - capPoint.x,
        capRim.at(-1)!.y - capPoint.y,
        capRim.at(-1)!.depth - capPoint.depth,
      ) > 1e-9) capRim.push(capPoint);
      if (segment === slice.segments) continue;
      const nextAngle = slice.start + (slice.end - slice.start) * (segment + 1)
        / slice.segments;
      const topStart = pointAt(slice, angle, topY);
      const bottomStart = pointAt(slice, angle, bottomY);
      const bottomEnd = pointAt(slice, nextAngle, bottomY);
      const topEnd = pointAt(slice, nextAngle, topY);
      const sideVisible = projection.cameraFacing([
        topStart, bottomStart, bottomEnd, topEnd,
      ]);
      if (sideVisible) {
        const lowerStart = pointAt(slice, angle, lowerY);
        const lowerEnd = pointAt(slice, nextAngle, lowerY);
        if (!lowerRun.length) lowerRun.push(lowerStart);
        lowerRun.push(lowerEnd);
        visibleBoundaryAngles.add(angleKey(angle));
        visibleBoundaryAngles.add(angleKey(nextAngle));
      } else if (lowerRun.length) {
        lowerRuns.push(lowerRun);
        lowerRun = [];
      }
    }
  }
  if (lowerRun.length) lowerRuns.push(lowerRun);
  if (lowerRuns.length > 1) {
    const first = lowerRuns[0];
    const last = lowerRuns.at(-1)!;
    const firstPoint = first[0];
    const lastPoint = last.at(-1)!;
    if (Math.hypot(
      firstPoint.x - lastPoint.x,
      firstPoint.y - lastPoint.y,
      firstPoint.depth - lastPoint.depth,
    ) <= 1e-9) {
      lowerRuns[0] = [...last, ...first.slice(1)];
      lowerRuns.pop();
    }
  }

  addPath(capRim, true);
  const verticalAngles = new Map<string, { slice: ThreeDPieOutlineSlice; angle: number }>();
  for (const { slice, angle } of boundarySlices.values()) {
    addPath([centerAt(slice), pointAt(slice, angle, capY)]);
    if (visibleBoundaryAngles.has(angleKey(angle))) {
      verticalAngles.set(angleKey(angle), { slice, angle });
    }
  }
  for (const run of lowerRuns) {
    addPath(run);
    // The visible cylindrical wall ends at two silhouette generators. They
    // are not necessarily slice boundaries, but they must connect the upper
    // and lower rims; omitting them leaves the two short black stubs visible
    // at the left and right extremes of the pie.
    for (const endpoint of [run[0], run.at(-1)!]) {
      const angle = Math.atan2(
        (endpoint.depth - reference.centerDepth) * projection.modelDepth,
        endpoint.x - reference.centerX,
      );
      verticalAngles.set(angleKey(angle), { slice: reference, angle });
    }
  }
  for (const { slice, angle } of verticalAngles.values()) {
    addPath([pointAt(slice, angle, capY), pointAt(slice, angle, lowerY)]);
  }
  return result;
}

function paintSceneFace(ctx: CanvasRenderingContext2D, item: SceneFace): void {
  if (item.paint === null) return;
  if (item.paint !== undefined) {
    polygon(ctx, item.points);
    ctx.fillStyle = item.paint;
    ctx.fill();
  } else {
    face(ctx, item.points, item.color, item.shade);
  }
  if (!item.outline) return;
  ctx.strokeStyle = item.outlineColor ?? 'rgba(0,0,0,0.42)';
  ctx.lineWidth = item.outlineWidth ?? 0.75;
  ctx.setLineDash(pptxPresetDashArray(item.outlineDash ?? 'solid', ctx.lineWidth));
  ctx.lineCap = item.outlineCap ?? 'butt';
  ctx.lineJoin = item.outlineJoin ?? 'miter';
  if (item.outline) {
    polygon(ctx, item.points);
    ctx.stroke();
  }
  ctx.setLineDash([]);
}

/** Resolve one authored fill recipe once for a complete projected datum, then
 * share the Canvas paint across its faces. This keeps gradient-stop work
 * O(datums × stops), never O(faces × stops). */
function applyScenePaint(
  ctx: CanvasRenderingContext2D,
  faces: SceneFace[],
  role: SceneFace['paintRole'],
  fill: Fill | null | undefined,
  shapeRotationDeg: number,
): void {
  if (fill === undefined) return;
  const targets = faces.filter(face => face.paintRole === role);
  if (!targets.length) return;
  if (fill === null) {
    for (const face of targets) face.paint = null;
    return;
  }
  const points = targets.flatMap(face => face.points);
  const minX = Math.min(...points.map(point => point.x));
  const maxX = Math.max(...points.map(point => point.x));
  const minY = Math.min(...points.map(point => point.y));
  const maxY = Math.max(...points.map(point => point.y));
  const paint = resolveFill(
    fill, ctx, minX, minY, maxX - minX, maxY - minY, shapeRotationDeg,
  );
  for (const face of targets) face.paint = paint;
}

function paintThreeDTooManyDataPoints(
  ctx: CanvasRenderingContext2D,
  rect: ChartRect,
): void {
  ctx.fillStyle = '#888';
  ctx.font = '12px sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText('(too many data points)', rect.x + rect.w / 2, rect.y + rect.h / 2);
}

function authoredCategoryRotation(chart: ChartModel): number | null {
  const raw = chart.catAxisLabelRotation;
  if (raw == null) return null;
  if (!Number.isFinite(raw) || Math.abs(raw) > 5_400_000) return 0;
  return raw / 60_000 * Math.PI / 180;
}

function drawAngledCategoryLabel(
  ctx: CanvasRenderingContext2D,
  label: string,
  point: Point,
  rotation: number,
  horizontal: boolean,
  outwardY = 1,
  offset = 6,
): void {
  if (horizontal || rotation === 0) {
    ctx.textAlign = horizontal ? 'right' : 'center';
    ctx.textBaseline = horizontal ? 'middle' : outwardY < 0 ? 'bottom' : 'top';
    ctx.fillText(
      label,
      point.x + (horizontal ? -offset : 0),
      point.y + (horizontal ? 0 : outwardY * offset),
    );
    return;
  }
  ctx.save();
  ctx.translate(point.x, point.y + outwardY * offset);
  ctx.rotate(rotation);
  ctx.textAlign = outwardY < 0 ? 'left' : 'right';
  ctx.textBaseline = 'middle';
  ctx.fillText(label, 0, 0);
  ctx.restore();
}

function paintThreeDMarker(
  ctx: CanvasRenderingContext2D,
  point: Point,
  symbol: string,
  size: number,
  fill: string,
  line: string,
  lineWidth: number,
  fillPaint: Fill | null | undefined = undefined,
  shapeRotationDeg = 0,
  ptToPx = PT_TO_PX,
): void {
  if (!(size > 0) || !Number.isFinite(point.x) || !Number.isFinite(point.y)) return;
  const radius = size / 2;
  ctx.beginPath();
  switch (symbol) {
    case 'square':
      ctx.rect(point.x - radius, point.y - radius, size, size);
      break;
    case 'diamond':
      ctx.moveTo(point.x, point.y - radius);
      ctx.lineTo(point.x + radius, point.y);
      ctx.lineTo(point.x, point.y + radius);
      ctx.lineTo(point.x - radius, point.y);
      ctx.closePath();
      break;
    case 'triangle':
      ctx.moveTo(point.x, point.y - radius);
      ctx.lineTo(point.x + radius, point.y + radius);
      ctx.lineTo(point.x - radius, point.y + radius);
      ctx.closePath();
      break;
    case 'x':
    case 'plus': {
      const diagonal = symbol === 'x';
      ctx.moveTo(point.x - radius, point.y + (diagonal ? -radius : 0));
      ctx.lineTo(point.x + radius, point.y + (diagonal ? radius : 0));
      ctx.moveTo(point.x + (diagonal ? -radius : 0), point.y + radius);
      ctx.lineTo(point.x + (diagonal ? radius : 0), point.y - radius);
      ctx.strokeStyle = line;
      ctx.lineWidth = lineWidth;
      ctx.stroke();
      return;
    }
    case 'dash':
      // ECMA-376 ST_MarkerStyle `dash` is the same filled short rectangle in
      // the 2-D and optional 3-D paths; unlike x/plus it consumes fill paint.
      ctx.rect(point.x - radius, point.y - size * 0.1, size, size * 0.2);
      break;
    case 'star': {
      for (let index = 0; index < 10; index++) {
        const angle = -Math.PI / 2 + index * Math.PI / 5;
        const r = index % 2 === 0 ? radius : radius * 0.4;
        const x = point.x + Math.cos(angle) * r;
        const y = point.y + Math.sin(angle) * r;
        index ? ctx.lineTo(x, y) : ctx.moveTo(x, y);
      }
      ctx.closePath();
      break;
    }
    case 'dot':
      // ECMA-376 §21.2.3.27: width=1/2 and height=1/5 of marker size.
      ctx.ellipse(point.x, point.y, size * 0.25, size * 0.1, 0, 0, Math.PI * 2);
      break;
    case 'picture': {
      if (fillPaint?.fillType === 'image') {
        paintChartImageFill(
          ctx, fillPaint, point.x - radius, point.y - radius, size, size, ptToPx,
          shapeRotationDeg,
        );
      }
      ctx.strokeStyle = line;
      ctx.lineWidth = lineWidth;
      ctx.strokeRect(point.x - radius, point.y - radius, size, size);
      return;
    }
    default:
      ctx.arc(point.x, point.y, radius, 0, Math.PI * 2);
      break;
  }
  const imageFill = fillPaint?.fillType === 'image' ? fillPaint : undefined;
  const resolved = imageFill ? null : fillPaint === undefined
    ? (fill === 'transparent' ? null : fill)
    : fillPaint == null
      ? null
      : resolveFill(
          fillPaint, ctx, point.x - radius, point.y - radius, size, size, shapeRotationDeg,
        );
  if (resolved != null) {
    ctx.fillStyle = resolved;
    ctx.fill();
  } else if (imageFill) {
    ctx.save();
    ctx.clip();
    paintChartImageFill(
      ctx, imageFill, point.x - radius, point.y - radius, size, size, ptToPx,
      shapeRotationDeg,
    );
    ctx.restore();
  }
  ctx.strokeStyle = line;
  ctx.lineWidth = lineWidth;
  ctx.stroke();
}

function drawThreeDDataLabel(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  series: ChartSeries,
  seriesIndex: number,
  categoryIndex: number,
  value: number,
  anchor: Point,
  bounds: ChartRect,
  ptToPx: number,
  markerGap = 0,
  percentValue?: number,
  resolvedOverride?: ChartDataLabelOverride,
  defaultPosition = 't',
  leaderAnchor: Point = anchor,
  leaderLineEligible = true,
  valueDisplayUnits?: ChartDisplayUnits | null,
  axisMaximum?: number,
  plottedValueForAxis = value,
  shapeRotationDeg = 0,
): void {
  // Callers resolve indexed overrides through their per-series Map before the
  // paint loop. Falling back to Array.find here would quietly reintroduce
  // quadratic work for a fully-authored maximum-size public model.
  const override = resolvedOverride;
  const defaults = series.seriesDataLabels;
  if (dataLabelIsDeleted(defaults, override)) return;
  if (chart.showDataLabelsOverMax !== true
    && axisMaximum != null
    && Number.isFinite(axisMaximum)
    && plottedValueForAxis > axisMaximum) return;
  const showVal = override?.showVal ?? defaults?.showVal ?? chart.showDataLabels;
  const showCat = override?.showCatName ?? defaults?.showCatName ?? false;
  const showSeries = override?.showSerName ?? defaults?.showSerName ?? false;
  const showPercent = override?.showPercent ?? defaults?.showPercent ?? false;
  const explicitText = override?.text;
  const text = effectiveDataLabelText({
    customText: explicitText,
    showCategory: showCat,
    showSeries,
    showValue: showVal,
    showPercent,
    category: series.categories?.[categoryIndex]
      ?? chart.categories[categoryIndex] ?? `${categoryIndex + 1}`,
    seriesName: series.name || `Series ${seriesIndex + 1}`,
    sourceValue: value,
    valueDivisor: valueDisplayUnits?.divisor,
    percentRatio: percentValue != null && Number.isFinite(percentValue) ? percentValue : undefined,
    formatCode: override?.formatCode
      ?? defaults?.formatCode ?? chart.dataLabelFormatCode ?? series.valFormatCode,
    percentFormatCode: override?.formatCode
      ?? defaults?.formatCode ?? chart.dataLabelFormatCode ?? '0%',
    date1904: chart.date1904,
    separator: override?.separator ?? defaults?.separator,
  });
  if (!text) return;
  const fontPx = chartTextFontSizePx(
    override?.fontSizeHpt ?? defaults?.fontSizeHpt ?? chart.dataLabelFontSizeHpt,
    ptToPx,
  ) ?? 9 * ptToPx;
  const bold = override?.fontBold ?? defaults?.fontBold ?? chart.dataLabelFontBold ?? false;
  const textStyle = effectiveDataLabelTextStyle(override, defaults);
  const fallbackFamily = chartFontFamily(
    chart, override?.fontFace ?? defaults?.fontFace ?? chart.dataLabelFontFace,
  );
  ctx.font = `${textStyle.fontItalic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontPx}px ${fallbackFamily}`;
  const fallbackColor = `#${override?.fontColor ?? defaults?.fontColor
    ?? series.labelColor ?? chart.dataLabelFontColor ?? '111111'}`;
  const rich = explicitText && override?.richRuns?.length
    ? resolveRichDataLabelBlock(ctx, {
      runs: override.richRuns,
      ptToPx,
      fontFamily: fallbackFamily,
      fallbackBold: bold,
      fallbackItalic: textStyle.fontItalic,
      fallbackBaseline: textStyle.fontBaseline,
      fallbackColorHidden: textStyle.fontPaintAuthored === true
        && (textStyle.fontHidden === true || textStyle.fontColor == null),
      fontFamilyForFace: face => chartFontFamily(chart, face),
    }, fontPx, fallbackColor)
    : null;
  const lines = rich ? [] : fitStyledDataLabelLines(
    text,
    Math.max(fontPx, bounds.w * 0.45),
    Math.max(fontPx * 1.2, bounds.h * 0.35),
    fontPx * 1.2,
    value => ctx.measureText(value).width,
    textStyle,
  );
  if (!rich && !lines.length) return;
  const textWidth = rich?.width ?? Math.max(...lines.map(line => ctx.measureText(line).width));
  const textHeight = rich?.height ?? lines.length * fontPx * 1.2;
  const insets = dataLabelInsets(textStyle, ptToPx);
  const rotated = rotatedDataLabelSize(
    textWidth + insets.left + insets.right,
    textHeight + insets.top + insets.bottom,
    textStyle.textRotation,
    textStyle.textVerticalMode,
  );
  const placement = resolveDataLabelPlacement(
    {
      kind: 'point', x: anchor.x, y: anchor.y,
      position: override?.position ?? defaults?.position ?? chart.dataLabelPosition ?? defaultPosition,
      markerGap,
    },
    bounds,
    { w: rotated.w, h: rotated.h },
    fontPx,
    override?.manualLayout,
    bounds,
  );
  if (!placement) return;
  const labelBox = mergeChartLabelBoxes(override?.labelBox, defaults?.labelBox);
  if (defaults?.showLeaderLines && defaults.leaderLineHidden !== true && leaderLineEligible) {
    ctx.beginPath();
    ctx.moveTo(leaderAnchor.x, leaderAnchor.y);
    ctx.lineTo(
      Math.max(placement.rect.x, Math.min(leaderAnchor.x, placement.rect.x + placement.rect.w)),
      Math.max(placement.rect.y, Math.min(leaderAnchor.y, placement.rect.y + placement.rect.h)),
    );
    ctx.strokeStyle = `#${defaults.leaderLineColor ?? '808080'}`;
    ctx.lineWidth = defaults.leaderLineWidthEmu != null
      ? Math.max(0.25, defaults.leaderLineWidthEmu / EMU_PER_PT * ptToPx)
      : 0.75 * ptToPx;
    ctx.setLineDash(pptxPresetDashArray(defaults.leaderLineDash ?? 'solid', ctx.lineWidth));
    ctx.stroke();
  }
  paintChartLabelBox(ctx, labelBox, placement.rect, ptToPx, shapeRotationDeg);
  ctx.save();
  ctx.beginPath();
  ctx.rect(placement.clip.x, placement.clip.y, placement.clip.w, placement.clip.h);
  ctx.clip();
  const paintAlign = dataLabelCanvasTextAlign(textStyle, placement.textAlign);
  const anchored = anchoredDataLabelPoint(
    placement.x, placement.y, placement.rect,
    textHeight + insets.top + insets.bottom, textStyle, override?.manualLayout != null,
    paintAlign, placement.textAlign,
    textWidth + insets.left + insets.right, rotated.radians,
  );
  const transformed = transformDataLabelText(
    ctx, anchored.x, anchored.y, rotated.radians, paintAlign,
    placement.textBaseline, insets,
  );
  if (rich) {
    paintRichDataLabelBlock(
      ctx, rich, transformed.x, transformed.y, paintAlign, placement.textBaseline,
      override?.manualLayout
        ? Math.max(0, placement.rect.w - insets.left - insets.right) : rich.width,
    );
    ctx.restore();
    return;
  }
  if (!(textStyle.fontPaintAuthored === true
    && (textStyle.fontHidden === true || textStyle.fontColor == null))) {
    ctx.fillStyle = fallbackColor;
    ctx.textAlign = paintAlign;
    ctx.textBaseline = 'middle';
    const lineHeight = fontPx * 1.2;
    const baselineShift = (textStyle.fontBaseline ?? 0) * fontPx;
    const firstY = transformed.y - (lines.length - 1) * lineHeight / 2 - baselineShift;
    lines.forEach((line, index) => ctx.fillText(line, transformed.x, firstY + index * lineHeight));
  }
  ctx.restore();
}

function wantsThreeDDataLabel(
  chart: ChartModel,
  series: ChartSeries,
  override: ChartDataLabelOverride | undefined,
): boolean {
  const defaults = series.seriesDataLabels;
  if (dataLabelIsDeleted(defaults, override)) return false;
  if (override?.text) return true;
  return (override?.showVal ?? defaults?.showVal ?? chart.showDataLabels)
    || (override?.showCatName ?? defaults?.showCatName ?? false)
    || (override?.showSerName ?? defaults?.showSerName ?? false)
    || (override?.showPercent ?? defaults?.showPercent ?? false);
}

function titleAndPlot(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  orientation: 'vertical' | 'horizontal' | 'radial',
  shapeRotationDeg: number,
): { plot: ChartRect; legend: ChartRect | null; legendMeasure: ThreeDLegendMeasure } {
  const band = cartesianTitleBand(chart, rect.h, ptToPx);
  if (chart.title) {
    ctx.font = `${chart.titleFontBold === false ? '' : 'bold '}${band.fontPx}px ${fontFamily(chart.titleFontFace)}`;
    ctx.fillStyle = chart.titleFontColor ? `#${chart.titleFontColor}` : '#111111';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    const measuredTitleWidth = ctx.measureText(chart.title).width;
    const automatic = {
      x: rect.x + (rect.w - measuredTitleWidth) / 2,
      y: rect.y + band.topPad,
      w: measuredTitleWidth,
      h: Math.max(1, band.fontPx),
    };
    const authored = chart.titleManualLayout
      ? resolveManualLayoutRect(
        { ...chart.titleManualLayout, w: undefined, h: undefined },
        rect,
        automatic,
      )
      : null;
    ctx.fillText(
      chart.title,
      authored ? authored.x + authored.w / 2 : rect.x + rect.w / 2,
      authored?.y ?? automatic.y,
    );
  }
  const legendOverrides = threeDLegendOverrideMap(chart);
  ctx.save();
  const rawLegendLabels = chart.chartType === 'pie'
    ? (chart.series[0]?.categories?.length ? chart.series[0].categories : chart.categories)
    : chart.series.map((series, index) => series.name || `Series ${index + 1}`);
  const legendLabels = rawLegendLabels.flatMap((label, index) =>
    legendOverrides.get(index)?.deleted === true ? [] : [{ label, index }]
  );
  const legendStyles = legendLabels.map(entry =>
    threeDLegendTextStyle(chart, legendOverrides.get(entry.index), ptToPx)
  );
  const legendFontPx = Math.max(0, ...legendStyles.map(style => style.fontPx));
  const itemWidths = legendLabels.map((entry, index) => {
      ctx.font = legendStyles[index].font;
      return 7 * ptToPx + 4 + ctx.measureText(entry.label).width;
    });
  const legendReserve = chartLegendReserve(chart, rect.w, rect.h, 0.23, {
    itemWidths,
    rowHeight: Math.max(legendFontPx * 1.45, 12),
    itemGap: 12,
    horizontalPadding: 8,
    verticalPadding: 4,
  });
  ctx.restore();
  const legendBands = chartLegendBands(legendReserve, chart.legendOverlay === true);
  const axisBands = chartAxisTitleBands(chart, rect.w, rect.h, ptToPx);
  const catAxisTitleInset = axisTitleVerticalInsetPx(
    chart.catAxisTitleTextVerticalInsetEmu, ptToPx,
  );
  const valAxisTitleInset = axisTitleVerticalInsetPx(
    chart.valAxisTitleTextVerticalInsetEmu, ptToPx,
  );
  const leftTitleBand = orientation === 'horizontal'
    ? chart.catAxisTitle
      ? axisBands.catFontPx + catAxisTitleInset + axisTitleMargin(rect.w) + 4
      : 0
    : orientation === 'vertical' ? axisBands.valBandW : 0;
  const bottomTitleBand = orientation === 'horizontal'
    ? chart.valAxisTitle
      ? axisBands.valFontPx + valAxisTitleInset + axisTitleMargin(rect.h) + 4
      : 0
    : orientation === 'vertical' ? axisBands.catBandH : 0;
  const frame = computeChartFrame(chart, rect.x, rect.y, rect.w, rect.h, ptToPx, {
    titleBand: band,
    legendSideReserveFrac: 0.23,
    legendReserve,
    pad: {
      t: band.bandH + legendBands.legTopH + rect.h * 0.04,
      r: legendBands.legRightW + rect.w * 0.05,
      b: legendBands.legBottomH + rect.h * 0.19 + bottomTitleBand,
      l: legendBands.legLeftW + rect.w * 0.13 + leftTitleBand,
    },
    honorPlotAreaManualLayout: true,
    manualOuterInsets: {
      t: band.bandH,
      r: legendBands.legRightW,
      b: legendBands.legBottomH + bottomTitleBand,
      l: legendBands.legLeftW + leftTitleBand,
    },
  });
  const plot: ChartRect = {
    x: frame.plotRect.px0,
    y: frame.plotRect.py0,
    w: Math.max(1, frame.plotRect.pw),
    h: Math.max(1, frame.plotRect.ph),
  };
  paintPlotAreaFrame(
    ctx, chart, plot.x, plot.y, plot.w, plot.h, ptToPx, shapeRotationDeg,
  );
  const defaultLegend = !legendReserve ? null
    : legendReserve.side === 'r'
      ? { x: rect.x + rect.w - legendReserve.reserveW, y: plot.y, w: legendReserve.reserveW, h: plot.h }
      : legendReserve.side === 'l'
        ? { x: rect.x, y: plot.y, w: legendReserve.reserveW, h: plot.h }
        : legendReserve.side === 't'
          ? { x: rect.x + 4, y: rect.y + band.bandH, w: Math.max(1, rect.w - 8), h: legendReserve.reserveH }
          : { x: rect.x + 4, y: rect.y + rect.h - legendReserve.reserveH, w: Math.max(1, rect.w - 8), h: legendReserve.reserveH };
  const manualLegend = defaultLegend && chart.legendManualLayout
    ? resolveManualLayoutRect(chart.legendManualLayout, rect, defaultLegend)
    : null;
  return {
    plot,
    legend: manualLegend ?? defaultLegend,
    legendMeasure: {
      labels: legendLabels.map(entry => entry.label),
      styles: legendStyles,
      itemWidths,
    },
  };
}

function drawThreeDAxisTitle(
  ctx: CanvasRenderingContext2D,
  text: string,
  chartRect: ChartRect,
  anchor: Point,
  side: ChartAxisTitleSide,
  fontSizePx: number,
  fontCssFamily: string,
  color: string | null | undefined,
  bold: boolean,
  italic: boolean,
  authoredRotation: number | null | undefined,
  authoredVerticalMode: ChartModel['catAxisTitleVerticalMode'],
  manualLayout: ChartModel['catAxisTitleManualLayout'],
  maxPx: number,
): void {
  ctx.save();
  ctx.font = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${fontSizePx}px ${fontCssFamily}`;
  ctx.fillStyle = color ? `#${color}` : '#555';
  const label = manualLayout ? text : elideToWidth(ctx, text, maxPx);
  const rotation = axisTitleRotationRad(side, authoredRotation, authoredVerticalMode);
  let center = anchor;
  if (manualLayout) {
    const textWidth = ctx.measureText(label).width;
    const cos = Math.abs(Math.cos(rotation));
    const sin = Math.abs(Math.sin(rotation));
    const automatic = {
      x: anchor.x - (textWidth * cos + fontSizePx * sin) / 2,
      y: anchor.y - (textWidth * sin + fontSizePx * cos) / 2,
      w: textWidth * cos + fontSizePx * sin,
      h: textWidth * sin + fontSizePx * cos,
    };
    const resolved = resolveManualLayoutRect(
      { ...manualLayout, w: undefined, h: undefined }, chartRect, automatic,
    );
    if (resolved) center = { x: resolved.x + resolved.w / 2, y: resolved.y + resolved.h / 2 };
  }
  ctx.translate(center.x, center.y);
  if (rotation) ctx.rotate(rotation);
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(label, 0, 0);
  ctx.restore();
}

/** Paint 3-D cartesian axis titles through the same authored font/orientation
 * contract as the 2-D families. Projection selects the axis rules; titles are
 * screen annotations in the already-reserved outer gutter, as in Office. */
function drawThreeDAxisTitles(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  chartRect: ChartRect,
  plot: ChartRect,
  horizontal: boolean,
  ptToPx: number,
): void {
  if (chart.valAxisTitle) {
    const fontPx = axisTitleFontPx(chart.valAxisTitleFontSizeHpt, ptToPx);
    const side: ChartAxisTitleSide = horizontal ? 'horizontal' : 'left';
    drawThreeDAxisTitle(
      ctx, chart.valAxisTitle, chartRect,
      horizontal
        ? { x: plot.x + plot.w / 2, y: plot.y + plot.h + axisTitleMargin(chartRect.h) + fontPx / 2 }
        : { x: plot.x - axisTitleMargin(chartRect.w) - fontPx / 2, y: plot.y + plot.h / 2 },
      side, fontPx, chartFontFamily(chart, chart.valAxisTitleFontFace, 'major'), chart.valAxisTitleFontColor,
      chart.valAxisTitleFontBold ?? true, chart.valAxisTitleFontItalic ?? false,
      chart.valAxisTitleRotation, chart.valAxisTitleVerticalMode,
      chart.valAxisTitleManualLayout, horizontal ? plot.w : plot.h,
    );
  }
  if (chart.catAxisTitle) {
    const fontPx = axisTitleFontPx(chart.catAxisTitleFontSizeHpt, ptToPx);
    const side: ChartAxisTitleSide = horizontal ? 'left' : 'horizontal';
    drawThreeDAxisTitle(
      ctx, chart.catAxisTitle, chartRect,
      horizontal
        ? { x: plot.x - axisTitleMargin(chartRect.w) - fontPx / 2, y: plot.y + plot.h / 2 }
        : { x: plot.x + plot.w / 2, y: plot.y + plot.h + axisTitleMargin(chartRect.h) + fontPx / 2 },
      side, fontPx, chartFontFamily(chart, chart.catAxisTitleFontFace, 'major'), chart.catAxisTitleFontColor,
      chart.catAxisTitleFontBold ?? true, chart.catAxisTitleFontItalic ?? false,
      chart.catAxisTitleRotation, chart.catAxisTitleVerticalMode,
      chart.catAxisTitleManualLayout, horizontal ? plot.h : plot.w,
    );
  }
}

/** Paint `<c:serAx>` for the `standard` 3-D bar arrangement. Unlike a legend,
 * this is a projected coordinate axis: its rule and tick anchors use the same
 * model-space depth coordinate as the bar solids. */
function drawThreeDSeriesAxis(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  chartRect: ChartRect,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  orientation: 'vertical' | 'horizontal',
  barFamily: boolean,
  ptToPx: number,
): void {
  const spec = chart.threeD?.seriesAxis;
  if (!spec || spec.hidden || chart.threeD?.barGrouping !== 'standard' || chart.series.length === 0) return;
  const geometry = threeDAxisGeometry(
    chart, projection, axis, categoryCount, categoryBetween, categoryReversed, orientation,
  );
  const wallGeometry = threeDWallGeometry(projection);
  // A series axis is the depth edge opposite the visible side wall. In the
  // default Office view the side wall is on the left and `<c:serAx>` is the
  // right-hand floor edge; anchoring it to the value-axis corner incorrectly
  // puts depth ticks and labels on the left.
  const seriesAxisX = orientation === 'vertical'
    ? wallGeometry.seriesAxisX : geometry.axisX;
  const seriesAxisY = orientation === 'horizontal'
    ? (wallGeometry.floorY === projection.front.y
      ? projection.front.y + projection.front.h : projection.front.y)
    : wallGeometry.floorY;
  const start = projection.project(
    seriesAxisX, seriesAxisY, projection.topology.nearDepth,
  );
  const end = projection.project(
    seriesAxisX, seriesAxisY, projection.topology.farDepth,
  );
  const length = Math.hypot(end.x - start.x, end.y - start.y);
  if (!(length > 1e-6)) return;
  const tangent = { x: (end.x - start.x) / length, y: (end.y - start.y) / length };
  let normal = { x: -tangent.y, y: tangent.x };
  const sceneCenter = projection.project(
    projection.front.x + projection.front.w / 2,
    projection.front.y + projection.front.h / 2,
    0.5,
  );
  const midpoint = { x: (start.x + end.x) / 2, y: (start.y + end.y) / 2 };
  if ((midpoint.x - sceneCenter.x) * normal.x + (midpoint.y - sceneCenter.y) * normal.y < 0) {
    normal = { x: -normal.x, y: -normal.y };
  }
  if ((spec.tickLabelPos === 'low' && normal.y < 0)
    || (spec.tickLabelPos === 'high' && normal.y > 0)) {
    normal = { x: -normal.x, y: -normal.y };
  }

  if (!spec.lineHidden) {
    applyThreeDStroke(ctx, threeDAxisStroke(
      spec.lineColor, spec.lineWidthEmu, spec.lineDash, ptToPx,
    ));
    ctx.beginPath();
    ctx.moveTo(start.x, start.y);
    ctx.lineTo(end.x, end.y);
    ctx.stroke();
  }

  const markSkip = Math.max(1, Math.floor(spec.tickMarkSkip ?? 1));
  const labelSkip = Math.max(1, Math.floor(spec.tickLabelSkip ?? 1));
  const tickMode = spec.majorTickMark ?? 'out';
  const fontPx = chartTextFontSizePx(spec.fontSizeHpt, ptToPx) ?? 9 * ptToPx;
  ctx.font = `${spec.fontItalic ? 'italic ' : ''}${spec.fontBold ? 'bold ' : ''}${fontPx}px ${chartFontFamily(chart, spec.fontFace)}`;
  ctx.fillStyle = spec.fontColor ? `#${spec.fontColor}` : '#595959';
  ctx.textAlign = Math.abs(normal.x) < 0.2 ? 'center' : normal.x < 0 ? 'right' : 'left';
  ctx.textBaseline = Math.abs(normal.y) < 0.2 ? 'middle' : normal.y < 0 ? 'bottom' : 'top';
  if (barFamily && !spec.lineHidden && tickMode !== 'none') {
    applyThreeDStroke(ctx, threeDAxisTickStroke(
      spec.lineColor, spec.lineWidthEmu, spec.lineDash, ptToPx,
    ));
    const boundaryCount = chart.series.length;
    for (let boundaryIndex = 0; boundaryIndex <= boundaryCount; boundaryIndex += markSkip) {
      const authoredDepth = boundaryIndex / boundaryCount;
      const depth = spec.orientation === 'maxMin' ? 1 - authoredDepth : authoredDepth;
      const point = projection.project(seriesAxisX, seriesAxisY, depth);
      const total = 6 * ptToPx;
      const outer = tickMode === 'cross' ? total / 2 : tickMode === 'out' ? total : 0;
      const inner = tickMode === 'cross' ? total / 2 : tickMode === 'in' ? total : 0;
      ctx.beginPath();
      ctx.moveTo(point.x + normal.x * outer, point.y + normal.y * outer);
      ctx.lineTo(point.x - normal.x * inner, point.y - normal.y * inner);
      ctx.stroke();
    }
  }
  for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
    const authoredDepth = projection.seriesDepth(seriesIndex, chart.series.length, false);
    const depth = spec.orientation === 'maxMin' ? 1 - authoredDepth : authoredDepth;
    const point = projection.project(seriesAxisX, seriesAxisY, depth);
    if (!barFamily && !spec.lineHidden && seriesIndex % markSkip === 0 && tickMode !== 'none') {
      const total = 6 * ptToPx;
      const outer = tickMode === 'cross' ? total / 2 : tickMode === 'out' ? total : 0;
      const inner = tickMode === 'cross' ? total / 2 : tickMode === 'in' ? total : 0;
      ctx.beginPath();
      ctx.moveTo(point.x + normal.x * outer, point.y + normal.y * outer);
      ctx.lineTo(point.x - normal.x * inner, point.y - normal.y * inner);
      ctx.stroke();
    }
    if (spec.tickLabelPos !== 'none' && seriesIndex % labelSkip === 0) {
      ctx.fillText(
        chart.series[seriesIndex].name || `Series ${seriesIndex + 1}`,
        point.x + normal.x * (6 * ptToPx + 3),
        point.y + normal.y * (6 * ptToPx + 3),
      );
    }
  }
  ctx.setLineDash([]);

  if (spec.title) {
    const titleFontPx = axisTitleFontPx(spec.titleFontSizeHpt, ptToPx);
    drawThreeDAxisTitle(
      ctx,
      spec.title,
      chartRect,
      {
        x: midpoint.x + normal.x * (fontPx + titleFontPx + 12),
        y: midpoint.y + normal.y * (fontPx + titleFontPx + 12),
      },
      'horizontal',
      titleFontPx,
      chartFontFamily(chart, spec.titleFontFace, 'major'),
      spec.titleFontColor,
      spec.titleFontBold ?? true,
      spec.titleFontItalic ?? false,
      spec.titleRotation,
      spec.titleVerticalMode,
      spec.titleManualLayout,
      Math.max(projection.front.w, projection.front.h),
    );
  }
}

function paintThreeDLineLegendKey(
  ctx: CanvasRenderingContext2D,
  series: ChartSeries,
  color: string,
  x: number,
  y: number,
  key: number,
  ptToPx: number,
  shapeRotationDeg: number,
): void {
  if (series.lineHidden !== true) {
    const lineWidth = series.lineWidthEmu != null
      ? Math.max(0.5, series.lineWidthEmu / EMU_PER_PT * ptToPx)
      : Math.max(1, 2 * ptToPx);
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + key, y);
    ctx.strokeStyle = series.lineColor ? `#${series.lineColor}` : scaleHexColor(color, 0.70);
    ctx.lineWidth = lineWidth;
    ctx.setLineDash(pptxPresetDashArray(series.chartexStyle?.lineDash ?? 'solid', lineWidth));
    ctx.stroke();
    ctx.setLineDash([]);
  }
  if (series.showMarker !== true || series.markerSymbol === 'none') return;
  const symbol = series.markerSymbol ?? 'circle';
  const fill = seriesMarkerFillColor(series, color.replace(/^#/, ''));
  const fillCss = fill === '00000000'
    ? 'transparent' : fill.startsWith('#') ? fill : `#${fill}`;
  const markerLine = series.markerLine ?? series.lineColor ?? color.replace(/^#/, '');
  const lineCss = markerLine === '00000000'
    ? 'rgba(0,0,0,0)' : markerLine.startsWith('#') ? markerLine : `#${markerLine}`;
  const markerLineWidth = series.markerLineWidthEmu != null
    ? Math.max(0.25, series.markerLineWidthEmu / EMU_PER_PT * ptToPx)
    : Math.max(0.75, ptToPx);
  paintThreeDMarker(
    ctx,
    { x: x + key / 2, y },
    symbol,
    Math.min(key, Math.max(2, (series.markerSize ?? 5) * ptToPx)),
    fillCss,
    lineCss,
    markerLineWidth,
    seriesMarkerFillPaint(series),
    shapeRotationDeg,
    ptToPx,
  );
}

function simpleLegend(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  bounds: ChartRect | null,
  ptToPx: number,
  categoryDriven = false,
  measured?: ThreeDLegendMeasure,
  shapeRotationDeg = 0,
): void {
  if (!bounds) return;
  paintLegendFrame(ctx, chart, bounds, ptToPx, shapeRotationDeg);
  const indexedPoints = new Map<number, ChartDataPointOverride>(
    chart.series[0]?.dataPointOverrides?.map(override => [override.idx, override]) ?? [],
  );
  const rawEntries = categoryDriven
    ? (chart.series[0]?.categories?.length ? chart.series[0].categories : chart.categories)
      .map((label, index) => {
        const authored = chart.series[0]?.dataPointColors?.[index];
        return {
          label,
          color: authored === '00000000'
            ? 'transparent'
            : authored ? `#${authored}` : colorFor(index),
          series: chart.series[0],
          point: indexedPoints.get(index),
          sourceIndex: index,
        };
      })
    : chart.series.map((series, index) => ({
      label: series.name || `Series ${index + 1}`,
      color: colorFor(index, series),
      series,
      point: undefined,
      sourceIndex: index,
    }));
  const legendOverrides = threeDLegendOverrideMap(chart);
  const entries = rawEntries.filter(entry =>
    legendOverrides.get(entry.sourceIndex)?.deleted !== true
  );
  const canReuseMeasure = measured != null
    && measured.labels.length === entries.length
    && measured.labels.every((label, index) => label === entries[index].label);
  const entryStyles = canReuseMeasure
    ? measured.styles
    : entries.map(entry =>
        threeDLegendTextStyle(chart, legendOverrides.get(entry.sourceIndex), ptToPx)
      );
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  const row = Math.max(Math.max(0, ...entryStyles.map(style => style.fontPx)) * 1.45, 12);
  const key = Math.min(7 * ptToPx, row * 0.7);
  const horizontal = chart.legendPos === 't' || chart.legendPos === 'b'
    || (chart.legendManualLayout != null && bounds.w >= bounds.h);
  if (horizontal) {
    const itemWidths = canReuseMeasure
      ? measured.itemWidths
      : entries.map((entry, index) => {
          ctx.font = entryStyles[index].font;
          return key + 4 + ctx.measureText(entry.label).width;
        });
    const rows = packLegendRows(itemWidths, Math.max(1, bounds.w - 8), 12);
    const visibleRows = rows.slice(
      0,
      Math.max(0, Math.floor((bounds.h - 4 + 1e-6) / row)),
    );
    let rowY = bounds.y + 2 + row / 2;
    for (const indices of visibleRows) {
      const widths = indices.map(index => Math.min(bounds.w, itemWidths[index]));
      const total = widths.reduce((sum, width) => sum + width, 0)
        + Math.max(0, indices.length - 1) * 12;
      let itemX = bounds.x + Math.max(4, (bounds.w - total) / 2);
      for (let position = 0; position < indices.length; position++) {
        const index = indices[position];
        const entry = entries[index];
        const textStyle = entryStyles[index];
        ctx.font = textStyle.font;
        const available = Math.max(0, widths[position] - key - 4);
        const lineKey = !categoryDriven && chart.chartType.toLowerCase().includes('line');
        if (lineKey && entry.series) {
          paintThreeDLineLegendKey(
            ctx, entry.series, entry.color, itemX, rowY, key, ptToPx, shapeRotationDeg,
          );
        } else {
          if (entry.color !== 'transparent') {
            ctx.fillStyle = entry.color;
            ctx.fillRect(itemX, rowY - key / 2, key, key);
          }
          const lineHidden = entry.point?.lineHidden ?? entry.series?.lineHidden;
          const lineColor = entry.point?.lineColor ?? entry.series?.lineColor;
          if (lineHidden !== true && lineColor) {
            ctx.strokeStyle = `#${lineColor}`;
            ctx.lineWidth = (entry.point?.lineWidthEmu ?? entry.series?.lineWidthEmu) != null
              ? Math.max(
                0.25,
                (entry.point?.lineWidthEmu ?? entry.series?.lineWidthEmu ?? 0)
                  / EMU_PER_PT * ptToPx,
              )
              : 0.75 * ptToPx;
            ctx.setLineDash(pptxPresetDashArray(
              entry.point?.lineDash ?? entry.series?.chartexStyle?.lineDash ?? 'solid',
              ctx.lineWidth,
            ));
            ctx.strokeRect(itemX, rowY - key / 2, key, key);
            ctx.setLineDash([]);
          }
        }
        ctx.fillStyle = textStyle.color;
        ctx.fillText(elideToWidth(ctx, entry.label, available), itemX + key + 4, rowY);
        itemX += widths[position] + 12;
      }
      rowY += row;
    }
    return;
  }
  let rowTop = bounds.y;
  for (let index = 0; index < entries.length; index++) {
    const entry = entries[index];
    const textStyle = entryStyles[index];
    ctx.font = textStyle.font;
    const textX = bounds.x + 8 + key;
    const availableTextWidth = Math.max(0, bounds.x + bounds.w - 4 - textX);
    const lineHeight = Math.max(textStyle.fontPx * 1.2, 10);
    const lines = fitDataLabelLines(
      entry.label,
      availableTextWidth,
      lineHeight * (categoryDriven ? 1 : 2),
      lineHeight,
      value => ctx.measureText(value).width,
    );
    if (lines.length === 0) continue;
    const itemHeight = Math.max(row, lines.length * lineHeight + 2);
    if (rowTop + itemHeight > bounds.y + bounds.h + 1e-6) break;
    const cy = rowTop + itemHeight / 2;
    const lineKey = !categoryDriven && chart.chartType.toLowerCase().includes('line');
    if (lineKey && entry.series) {
      paintThreeDLineLegendKey(
        ctx, entry.series, entry.color, bounds.x + 4, cy, key, ptToPx, shapeRotationDeg,
      );
    } else {
      if (entry.color !== 'transparent') {
        ctx.fillStyle = entry.color;
        ctx.fillRect(bounds.x + 4, cy - key / 2, key, key);
      }
      const lineHidden = entry.point?.lineHidden ?? entry.series?.lineHidden;
      const lineColor = entry.point?.lineColor ?? entry.series?.lineColor;
      if (lineHidden !== true && lineColor) {
        ctx.strokeStyle = `#${lineColor}`;
        ctx.lineWidth = (entry.point?.lineWidthEmu ?? entry.series?.lineWidthEmu) != null
          ? Math.max(0.25, (entry.point?.lineWidthEmu ?? entry.series?.lineWidthEmu ?? 0) / EMU_PER_PT * ptToPx)
          : 0.75 * ptToPx;
        ctx.setLineDash(pptxPresetDashArray(
          entry.point?.lineDash ?? entry.series?.chartexStyle?.lineDash ?? 'solid',
          ctx.lineWidth,
        ));
        ctx.strokeRect(bounds.x + 4, cy - key / 2, key, key);
        ctx.setLineDash([]);
      }
    }
    ctx.fillStyle = textStyle.color;
    const firstY = cy - (lines.length - 1) * lineHeight / 2;
    lines.forEach((line, lineIndex) => ctx.fillText(line, textX, firstY + lineIndex * lineHeight));
    rowTop += itemHeight;
  }
}

function axisPlan(
  chart: ChartModel,
  dataMin: number,
  dataMax: number,
  axisLenPt: number,
  percent: boolean,
  orientation: 'vertical' | 'horizontal',
): NumericValueAxisPlan {
  const factor = percent ? 100 : 1;
  const minorTickMark = chart.valAxisMinorTickMark ?? 'none';
  const options = {
    dataMin,
    dataMax,
    explicitMin: chart.valMin == null ? (percent ? dataMin : null) : chart.valMin * factor,
    explicitMax: chart.valMax == null ? (percent ? dataMax : null) : chart.valMax * factor,
    majorUnit: chart.valAxisMajorUnit == null ? null : chart.valAxisMajorUnit * factor,
    minorUnit: chart.valAxisMinorUnit == null ? null : chart.valAxisMinorUnit * factor,
    axisLenPt,
    axisOrientation: orientation,
    logBase: chart.valAxisLogBase,
    reversed: chart.valAxisOrientation === 'maxMin',
    needMinor: chart.valAxisMinorGridlines === true || minorTickMark !== 'none',
  } as const;
  const automaticPlan = planNumericValueAxis(options);

  // In the measured compact classic 3-D column charts, Excel applies its
  // explicit-bounds vertical density rule when only one value-axis bound is
  // authored. Keep that application-defined behavior local to the observed
  // family instead of changing the shared 2-D automatic-axis policy.
  const compactThreeDColumnWithOneBound = !percent
    && orientation === 'vertical'
    && (chart.chartType === 'clusteredBar' || chart.chartType === 'stackedBar')
    && chart.valAxisMajorUnit == null
    && chart.valAxisLogBase == null
    && chart.valAxisDisplayUnits == null
    && ((chart.valMin == null) !== (chart.valMax == null));
  if (!compactThreeDColumnWithOneBound) return automaticPlan;

  return planNumericValueAxis({
    ...options,
    explicitMin: automaticPlan.min,
    explicitMax: automaticPlan.max,
  });
}

interface ThreeDStroke {
  color: string;
  width: number;
  dash: number[];
}

export interface ThreeDWallGeometry {
  floor: readonly Point[];
  sideWall: readonly Point[];
  backWall: readonly Point[];
  /** Model-space category coordinate used by the visible side wall. */
  sideX: number;
  /** Opposite category coordinate used by the series/depth axis. */
  seriesAxisX: number;
  floorY: number;
  oppositeFloorY: number;
  nearDepth: 0 | 1;
  farDepth: 0 | 1;
}

export interface ThreeDPieSliceAngles {
  /** Ascending model-angle interval consumed by the sector mesh. */
  start: number;
  end: number;
  middle: number;
  /** Authored leading ray, where the clockwise slice begins. */
  leading: number;
}

/** Convert OOXML's clockwise first-slice angle from screen twelve o'clock into
 * the X/Z model plane used by the shared 3-D camera. Positive model Z projects
 * upward for the ordinary positive elevation, so twelve o'clock is +π/2 (not
 * Canvas' -π/2). The mesh interval stays ascending while retaining the
 * authored clockwise sector. */
export function threeDPieSliceAngles(
  firstSliceAngle: number | null | undefined,
  cumulativeFraction: number,
  sliceFraction: number,
): ThreeDPieSliceAngles {
  const normalizedFirst = firstSliceAngle != null && Number.isFinite(firstSliceAngle)
    ? ((firstSliceAngle % 360) + 360) % 360 : 0;
  const cumulative = Number.isFinite(cumulativeFraction)
    ? Math.max(0, Math.min(1, cumulativeFraction)) : 0;
  const fraction = Number.isFinite(sliceFraction)
    ? Math.max(0, Math.min(1 - cumulative, sliceFraction)) : 0;
  const leading = Math.PI / 2 - (normalizedFirst * Math.PI / 180 + cumulative * Math.PI * 2);
  const trailing = leading - fraction * Math.PI * 2;
  return {
    start: Math.min(leading, trailing),
    end: Math.max(leading, trailing),
    middle: (leading + trailing) / 2,
    leading,
  };
}

/** Resolve the three authored CT_Surface faces from the same projected chart
 * cuboid used by bars, grids and axes. Keeping the shared edges identical is
 * what makes the side wall, floor and back-wall separator close as one drawing
 * instead of a collection of screen-space approximations. */
export function threeDWallGeometry(
  projection: ChartThreeDProjection,
): ThreeDWallGeometry {
  const { front } = projection;
  const xMin = front.x;
  const xMax = front.x + front.w;
  const sideX = projection.topology.farX === 'min' ? xMin : xMax;
  const seriesAxisX = sideX === xMin ? xMax : xMin;
  const floorY = projection.topology.axisY === 'min' ? front.y : front.y + front.h;
  const oppositeFloorY = floorY === front.y ? front.y + front.h : front.y;
  const { nearDepth, farDepth } = projection.topology;
  const floorNearMin = projection.project(xMin, floorY, nearDepth);
  const floorNearMax = projection.project(xMax, floorY, nearDepth);
  const floorFarMax = projection.project(xMax, floorY, farDepth);
  const floorFarMin = projection.project(xMin, floorY, farDepth);
  const backTopMin = projection.project(xMin, oppositeFloorY, farDepth);
  const backTopMax = projection.project(xMax, oppositeFloorY, farDepth);
  const sideNearFloor = projection.project(sideX, floorY, nearDepth);
  const sideFarFloor = projection.project(sideX, floorY, farDepth);
  const sideFarTop = projection.project(sideX, oppositeFloorY, farDepth);
  const sideNearTop = projection.project(sideX, oppositeFloorY, nearDepth);
  return {
    floor: [floorNearMin, floorNearMax, floorFarMax, floorFarMin],
    sideWall: [sideNearFloor, sideFarFloor, sideFarTop, sideNearTop],
    backWall: [floorFarMin, floorFarMax, backTopMax, backTopMin],
    sideX,
    seriesAxisX,
    floorY,
    oppositeFloorY,
    nearDepth,
    farDepth,
  };
}

function threeDStroke(
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  ptToPx: number,
  fallbackColor = 'A6A6A6',
  fallbackWidth = 0.75,
): ThreeDStroke {
  const width = widthEmu != null && Number.isFinite(widthEmu) && widthEmu >= 0
    ? Math.max(0.25, widthEmu / EMU_PER_PT * ptToPx)
    : fallbackWidth * ptToPx;
  return {
    color: `#${color ?? fallbackColor}`,
    width,
    dash: pptxPresetDashArray(dash ?? 'solid', width),
  };
}

function threeDAxisStroke(
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  ptToPx: number,
): ThreeDStroke {
  const stroke = threeDStroke(color, widthEmu, dash, ptToPx, '898989', 1);
  if (widthEmu == null || !Number.isFinite(widthEmu) || widthEmu < 0) return stroke;
  // `<c:*Ax><c:spPr><a:ln@w>` owns the coordinate-rule width. Keep it in the
  // same projected point scale as the chart instead of applying a device-pixel
  // floor; otherwise zoom changes and a thin authored axis can become thicker
  // than the surrounding CT_Surface rules.
  const authoredWidth = widthEmu / EMU_PER_PT * ptToPx;
  if (!(authoredWidth > 0)) return stroke;
  const width = authoredWidth;
  return { ...stroke, width, dash: pptxPresetDashArray(dash ?? 'solid', width) };
}

function threeDAxisTickStroke(
  color: string | null | undefined,
  widthEmu: number | null | undefined,
  dash: string | null | undefined,
  ptToPx: number,
): ThreeDStroke {
  // Tick rules use the same authored DrawingML point scale as the coordinate
  // axes; they are kept separate here because their projected geometry and
  // default visibility are resolved independently.
  return threeDStroke(color, widthEmu, dash, ptToPx, '898989', 1);
}

function applyThreeDStroke(ctx: CanvasRenderingContext2D, stroke: ThreeDStroke): void {
  ctx.strokeStyle = stroke.color;
  ctx.lineWidth = stroke.width;
  ctx.setLineDash(stroke.dash);
}

function walls(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  orientation: 'vertical' | 'horizontal',
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  ptToPx: number,
): void {
  const { front } = projection;
  const xMin = front.x;
  const xMax = front.x + front.w;
  const geometry = threeDWallGeometry(projection);
  const {
    sideX: farX, floorY, oppositeFloorY, nearDepth, farDepth,
  } = geometry;
  const axisGeometry = threeDAxisGeometry(
    chart, projection, axis, categoryCount, categoryBetween, categoryReversed, orientation,
  );
  const axisOwnedEdges: Array<readonly [Point, Point]> = [];
  // Linked chart styles may resolve the same color/width fields without a
  // direct axis line. Only a direct `<c:*Ax><c:spPr><a:ln>` owns a coincident
  // CT_Surface perimeter; style-derived axes retain the complete wall frame.
  const categoryAxisPaintAuthored = chart.catAxisLinePaintAuthored === true;
  const valueAxisPaintAuthored = chart.valAxisLinePaintAuthored === true;
  const horizontalAxisVisible = orientation === 'vertical'
    ? categoryAxisPaintAuthored && !chart.catAxisHidden && !chart.catAxisLineHidden
    : valueAxisPaintAuthored && !chart.valAxisHidden && !chart.valAxisLineHidden;
  const verticalAxisVisible = orientation === 'vertical'
    ? valueAxisPaintAuthored && !chart.valAxisHidden && !chart.valAxisLineHidden
    : categoryAxisPaintAuthored && !chart.catAxisHidden && !chart.catAxisLineHidden;
  if (horizontalAxisVisible) {
    axisOwnedEdges.push([axisGeometry.horizontalStart, axisGeometry.horizontalEnd]);
  }
  if (verticalAxisVisible) {
    axisOwnedEdges.push([axisGeometry.verticalStart, axisGeometry.verticalEnd]);
  }
  const seriesAxis = chart.threeD?.seriesAxis;
  const seriesAxisPaintAuthored = seriesAxis?.linePaintAuthored === true;
  if (seriesAxis && !seriesAxis.hidden && !seriesAxis.lineHidden
    && seriesAxisPaintAuthored
    && chart.threeD?.barGrouping === 'standard' && chart.series.length > 0) {
    const seriesAxisX = orientation === 'vertical' ? geometry.seriesAxisX : axisGeometry.axisX;
    const seriesAxisY = orientation === 'horizontal'
      ? (floorY === projection.front.y
        ? projection.front.y + projection.front.h : projection.front.y)
      : floorY;
    axisOwnedEdges.push([
      projection.project(seriesAxisX, seriesAxisY, nearDepth),
      projection.project(seriesAxisX, seriesAxisY, farDepth),
    ]);
  }
  const samePoint = (left: Point, right: Point): boolean =>
    Math.hypot(left.x - right.x, left.y - right.y) <= 1e-6;
  const axisOwnsEdge = (start: Point, end: Point): boolean =>
    axisOwnedEdges.some(([axisStart, axisEnd]) =>
      (samePoint(start, axisStart) && samePoint(end, axisEnd))
      || (samePoint(start, axisEnd) && samePoint(end, axisStart))
    );
  const projectedSurfaceFaces = (
    kind: 'floor' | 'sideWall' | 'backWall',
    surface: NonNullable<ChartModel['threeD']>['floor'],
  ) => {
    const slab = planChartThreeDSurfaceGeometry(projection, kind, surface?.thicknessPercent);
    const faces = slab.faces
      .map((scenePoints, faceIndex) => ({ scenePoints, faceIndex }))
      .filter(({ scenePoints }) => slab.thickness === 0 || projection.cameraFacing(scenePoints))
      .map(({ scenePoints, faceIndex }) => ({
        faceIndex,
        scenePoints,
        points: scenePoints.map(point =>
          projection.projectUnbounded(point.x, point.y, point.depth)
        ),
        depth: scenePoints.reduce(
          (sum, point) => sum + projection.cameraDepth(point.x, point.y, point.depth),
          0,
        ) / scenePoints.length,
      }))
      .sort((left, right) => left.depth - right.depth);
    return {
      slab,
      faces,
      visibleFaceIndices: new Set(faces.map(face => face.faceIndex)),
    };
  };
  const floorSurface = projectedSurfaceFaces('floor', chart.threeD?.floor);
  const sideWallSurface = projectedSurfaceFaces('sideWall', chart.threeD?.sideWall);
  const backWallSurface = projectedSurfaceFaces('backWall', chart.threeD?.backWall);
  const floorFaces = floorSurface.faces;
  const sideWallFaces = sideWallSurface.faces;
  const backWallFaces = backWallSurface.faces;
  const drawSurfaceFill = (
    group: ReturnType<typeof projectedSurfaceFaces>,
    surface: NonNullable<ChartModel['threeD']>['floor'],
    role: 'floor' | 'wall',
    kind: 'floor' | 'sideWall' | 'backWall',
  ) => {
    // CT_Surface does not imply a paint when c:spPr/fill is omitted. Excel's
    // classic 3-D default leaves these faces unfilled; only an authored fill
    // creates an opaque wall. The wall/grid rules are painted independently.
    const effective = chartThreeDSurfacePaint(chart, surface, role);
    const faces = group.faces;
    if (!effective.fill || !faces.length) return;
    if (effective.fill.fillType === 'image') {
      const image = chartImageFillSource(effective.fill);
      if (!image) return;
      const project = (point: ThreeDScenePoint): Point =>
        projection.projectUnbounded(point.x, point.y, point.depth);
      paintChartThreeDSurfacePicture(
        ctx, effective.fill, image, surface, kind,
        group.slab, group.faces.map(face => face.faceIndex),
        project, axis.max - axis.min,
      );
      return;
    }
    const points = faces.flatMap(face => face.points);
    const minX = Math.min(...points.map(point => point.x));
    const maxX = Math.max(...points.map(point => point.x));
    const minY = Math.min(...points.map(point => point.y));
    const maxY = Math.max(...points.map(point => point.y));
    const fill = effective.fill.fillType === 'solid'
      ? `#${effective.fill.color}`
      : resolveFill(effective.fill, ctx, minX, minY, maxX - minX, maxY - minY);
    if (!fill) return;
    for (const face of faces) {
      if (face.points.length < 3) continue;
      polygon(ctx, face.points);
      ctx.fillStyle = fill;
      ctx.fill();
    }
  };
  drawSurfaceFill(floorSurface, chart.threeD?.floor, 'floor', 'floor');
  drawSurfaceFill(sideWallSurface, chart.threeD?.sideWall, 'wall', 'sideWall');
  drawSurfaceFill(backWallSurface, chart.threeD?.backWall, 'wall', 'backWall');
  const strokeSurfaceGridRule = (
    group: ReturnType<typeof projectedSurfaceFaces>,
    kind: 'floor' | 'sideWall' | 'backWall',
    coordinate: 'x' | 'y',
    fraction: number,
    includeExterior = true,
  ): void => {
    for (const segment of planChartThreeDSurfaceGridSegments(
      group.slab,
      kind,
      coordinate,
      fraction,
    )) {
      if (!includeExterior && segment.faceIndex !== 0) continue;
      if (!group.visibleFaceIndices.has(segment.faceIndex)) continue;
      const [start, end] = segment.scenePoints.map(point =>
        projection.projectUnbounded(point.x, point.y, point.depth)
      );
      ctx.beginPath();
      ctx.moveTo(start.x, start.y);
      ctx.lineTo(end.x, end.y);
      ctx.stroke();
    }
  };
  const drawValueGrid = (
    values: readonly number[],
    stroke: ThreeDStroke,
    authoredSurfaceRule: boolean,
  ) => {
    applyThreeDStroke(ctx, stroke);
    for (const value of values) {
      const fraction = axis.fraction(value);
      if (authoredSurfaceRule) {
        if (orientation === 'horizontal') {
          strokeSurfaceGridRule(floorSurface, 'floor', 'x', fraction);
          strokeSurfaceGridRule(backWallSurface, 'backWall', 'x', fraction);
        } else {
          strokeSurfaceGridRule(sideWallSurface, 'sideWall', 'y', fraction);
          strokeSurfaceGridRule(backWallSurface, 'backWall', 'y', fraction);
        }
        continue;
      }
      if (orientation === 'horizontal') {
        const x = front.x + fraction * front.w;
        const near = projection.project(x, floorY, farDepth);
        const far = projection.project(x, oppositeFloorY, farDepth);
        ctx.beginPath(); ctx.moveTo(near.x, near.y); ctx.lineTo(far.x, far.y); ctx.stroke();
      } else {
        // The same value plane crosses both the visible side wall and the back
        // wall. Leaving out the side segment makes the grid appear to change
        // direction at the wall boundary even though the scene is projective.
        const y = front.y + front.h - fraction * front.h;
        const sideNear = projection.project(farX, y, nearDepth);
        const sideFar = projection.project(farX, y, farDepth);
        const backEnd = projection.project(farX === xMin ? xMax : xMin, y, farDepth);
        ctx.beginPath(); ctx.moveTo(sideNear.x, sideNear.y); ctx.lineTo(sideFar.x, sideFar.y); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(sideFar.x, sideFar.y); ctx.lineTo(backEnd.x, backEnd.y); ctx.stroke();
      }
    }
  };
  if (chart.valAxisMinorGridlines === true) {
    drawValueGrid(axis.minorTicks, threeDStroke(
      chart.valAxisMinorGridlineColor,
      chart.valAxisMinorGridlineWidthEmu,
      chart.valAxisMinorGridlineDash,
      ptToPx,
      'D9D9D9',
      0.5,
    ), true);
  }
  // CT_ValAx/majorGridlines is an optional element, not an on-by-default
  // property. Its absence therefore contributes no value-grid geometry. This
  // mirrors the 2-D axis path and prevents synthetic back-wall rules from
  // appearing in 3-D charts that author only the axis line.
  if (chart.valAxisMajorGridlines === true) {
    drawValueGrid(axis.majorTicks, threeDStroke(
      chart.valAxisGridlineColor,
      chart.valAxisGridlineWidthEmu,
      chart.valAxisGridlineDash,
      ptToPx,
      '898989',
      1,
    ), true);
  }
  // CT_CatAx/majorGridlines is optional. Only an authored element contributes
  // category-grid geometry; omission must not synthesize depth rays. Authored
  // rules continue across visible thick surface faces.
  const strokeCategorySurfaceGrid = (
    fractions: readonly number[],
    stroke: ThreeDStroke,
  ): void => {
    applyThreeDStroke(ctx, stroke);
    for (const fraction of fractions) {
      if (orientation === 'vertical') {
        strokeSurfaceGridRule(floorSurface, 'floor', 'x', fraction);
        strokeSurfaceGridRule(backWallSurface, 'backWall', 'x', fraction);
      } else {
        strokeSurfaceGridRule(sideWallSurface, 'sideWall', 'y', fraction);
        strokeSurfaceGridRule(backWallSurface, 'backWall', 'y', fraction);
      }
    }
  };
  if (chart.catAxisMinorGridlines === true) {
    strokeCategorySurfaceGrid(
      categoryMinorGridlineFractions(categoryCount, categoryBetween),
      threeDStroke(
        chart.catAxisMinorGridlineColor,
        chart.catAxisMinorGridlineWidthEmu,
        chart.catAxisMinorGridlineDash,
        ptToPx,
        'E0E0E0',
        0.5,
      ),
    );
  }
  if (chart.catAxisMajorGridlines === true) {
    strokeCategorySurfaceGrid(
      categoryGridlineFractions(categoryCount, categoryBetween),
      threeDStroke(
        chart.catAxisGridlineColor,
        chart.catAxisGridlineWidthEmu,
        chart.catAxisGridlineDash,
        ptToPx,
        'E0E0E0',
        0.5,
      ),
    );
  }
  const strokeSurface = (
    faces: readonly { points: readonly Point[] }[],
    surface: NonNullable<ChartModel['threeD']>['floor'],
    role: 'floor' | 'wall',
  ) => {
    if (!faces.length) return;
    const effective = chartThreeDSurfacePaint(chart, surface, role);
    if (effective.line === null) return;
    const linePaint = effective.line ?? { fillType: 'solid' as const, color: '898989' };
    const points = faces.flatMap(face => face.points);
    const minX = Math.min(...points.map(point => point.x));
    const maxX = Math.max(...points.map(point => point.x));
    const minY = Math.min(...points.map(point => point.y));
    const maxY = Math.max(...points.map(point => point.y));
    const line = linePaint.fillType === 'solid'
      ? `#${linePaint.color}`
      : resolveFill(linePaint, ctx, minX, minY, maxX - minX, maxY - minY);
    if (!line) return;
    const width = effective.lineWidthEmu != null
      ? Math.max(0.25, effective.lineWidthEmu / EMU_PER_PT * ptToPx)
      : ptToPx;
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
    for (const face of faces) {
      if (face.points.length < 2) continue;
      const edges = face.points.map((point, index) => ({
        start: point,
        end: face.points[(index + 1) % face.points.length],
      }));
      if (edges.some(edge => axisOwnsEdge(edge.start, edge.end))) {
        // A visible axis and a CT_Surface perimeter can describe the same
        // projected cuboid edge with different authored paints. Office gives
        // the coordinate axis visual ownership; painting both produces the
        // gray halo seen around a thinner black axis. Retain every independent
        // surface edge and omit only the exactly coincident segment.
        for (const edge of edges) {
          if (axisOwnsEdge(edge.start, edge.end)) continue;
          ctx.beginPath();
          ctx.moveTo(edge.start.x, edge.start.y);
          ctx.lineTo(edge.end.x, edge.end.y);
          ctx.stroke();
        }
        continue;
      }
      ctx.beginPath();
      ctx.moveTo(face.points[0].x, face.points[0].y);
      for (let index = 1; index < face.points.length; index++) {
        ctx.lineTo(face.points[index].x, face.points[index].y);
      }
      ctx.lineTo(face.points[0].x, face.points[0].y);
      ctx.closePath();
      ctx.stroke();
    }
  };
  // Paint each CT_Surface perimeter in chart order. The back wall is last, so
  // its authored rule owns the shared floor/back separator without drawing an
  // extra synthetic line. This also restores the complete side-wall outline.
  strokeSurface(floorFaces, chart.threeD?.floor, 'floor');
  strokeSurface(sideWallFaces, chart.threeD?.sideWall, 'wall');
  strokeSurface(backWallFaces, chart.threeD?.backWall, 'wall');
  ctx.setLineDash([]);
}

/** Front scene axes are an overlay, not part of the back-wall paint. Repaint
 * them after data faces so a near cuboid cannot erase the coordinate frame. */
function frontAxes(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  ptToPx: number,
  orientation: 'vertical' | 'horizontal',
): void {
  const geometry = threeDAxisGeometry(
    chart, projection, axis, categoryCount, categoryBetween, categoryReversed, orientation,
  );
  const drawHorizontal = orientation === 'vertical'
    ? !chart.catAxisHidden && !chart.catAxisLineHidden
    : !chart.valAxisHidden && !chart.valAxisLineHidden;
  const drawVertical = orientation === 'vertical'
    ? !chart.valAxisHidden && !chart.valAxisLineHidden
    : !chart.catAxisHidden && !chart.catAxisLineHidden;
  if (drawHorizontal) {
    applyThreeDStroke(ctx, threeDAxisStroke(
      orientation === 'vertical' ? chart.catAxisLineColor : chart.valAxisLineColor,
      orientation === 'vertical' ? chart.catAxisLineWidthEmu : chart.valAxisLineWidthEmu,
      orientation === 'vertical' ? chart.catAxisLineDash : chart.valAxisLineDash,
      ptToPx,
    ));
    ctx.beginPath();
    ctx.moveTo(geometry.horizontalStart.x, geometry.horizontalStart.y);
    ctx.lineTo(geometry.horizontalEnd.x, geometry.horizontalEnd.y);
    ctx.stroke();
  }
  if (drawVertical) {
    applyThreeDStroke(ctx, threeDAxisStroke(
      orientation === 'vertical' ? chart.valAxisLineColor : chart.catAxisLineColor,
      orientation === 'vertical' ? chart.valAxisLineWidthEmu : chart.catAxisLineWidthEmu,
      orientation === 'vertical' ? chart.valAxisLineDash : chart.catAxisLineDash,
      ptToPx,
    ));
    ctx.beginPath();
    ctx.moveTo(geometry.verticalStart.x, geometry.verticalStart.y);
    ctx.lineTo(geometry.verticalEnd.x, geometry.verticalEnd.y);
    ctx.stroke();
  }
  ctx.setLineDash([]);
}

/** Draw a tick in screen space from the normal of its already-projected axis.
 * The anchor comes from the same 3-D point as the axis/grid/data geometry; only
 * the physical tick length is a CSS/print dimension. This keeps 6pt/4pt tick
 * lengths readable without reintroducing an unprojected horizontal/vertical
 * segment that would disagree with a rotated axis. */
function screenAnnotationOutward(
  axisStart: Point,
  axisEnd: Point,
  projectedSceneCenter: Point,
  screenDirection: 'horizontal' | 'vertical',
): Point {
  const axisMidpoint = {
    x: (axisStart.x + axisEnd.x) / 2,
    y: (axisStart.y + axisEnd.y) / 2,
  };
  return screenDirection === 'horizontal'
    ? { x: axisMidpoint.x <= projectedSceneCenter.x ? -1 : 1, y: 0 }
    : { x: 0, y: axisMidpoint.y <= projectedSceneCenter.y ? -1 : 1 };
}

function projectedAxisTick(
  ctx: CanvasRenderingContext2D,
  mode: string | null | undefined,
  anchor: Point,
  axisStart: Point,
  axisEnd: Point,
  projectedSceneCenter: Point,
  screenDirection: 'horizontal' | 'vertical',
  level: 'major' | 'minor',
  ptToPx: number,
): void {
  if (!mode || mode === 'none') return;
  // Office paints 3-D ticks as screen annotations: value-axis ticks remain
  // horizontal and category-axis ticks vertical even when their projected
  // axis rule is sloped. Only the outward sign follows the selected front edge.
  const normal = screenAnnotationOutward(
    axisStart, axisEnd, projectedSceneCenter, screenDirection,
  );
  const length = projectedAxisTickLengthPx(level, ptToPx);
  const sideLength = mode === 'cross' ? length / 2 : length;
  const outer = mode === 'out' || mode === 'cross' ? sideLength : 0;
  const inner = mode === 'in' || mode === 'cross' ? sideLength : 0;
  ctx.beginPath();
  ctx.moveTo(anchor.x + normal.x * outer, anchor.y + normal.y * outer);
  ctx.lineTo(anchor.x - normal.x * inner, anchor.y - normal.y * inner);
  ctx.stroke();
}

function projectedAxisTickLengthPx(level: 'major' | 'minor', ptToPx: number): number {
  return (level === 'minor' ? 4 : 6) * ptToPx;
}

function projectedAxisTickOutwardExtentPx(
  mode: string | null | undefined,
  level: 'major' | 'minor',
  ptToPx: number,
): number {
  if (mode !== 'out' && mode !== 'cross') return 0;
  const length = projectedAxisTickLengthPx(level, ptToPx);
  return mode === 'cross' ? length / 2 : length;
}

function projectedAxisTickLabelOffsetPx(
  mode: string | null | undefined,
  lineHidden: boolean | null | undefined,
  ptToPx: number,
  previousMinimum: number,
): number {
  if (lineHidden) return previousMinimum;
  const tickOutset = projectedAxisTickOutwardExtentPx(mode, 'major', ptToPx);
  return Math.max(previousMinimum, tickOutset + 3 * ptToPx);
}

interface ThreeDAxisGeometry {
  axisX: number;
  axisY: number;
  depth: number;
  horizontalStart: Point;
  horizontalEnd: Point;
  verticalStart: Point;
  verticalEnd: Point;
}

/** Resolve authored axis crossing in model space before projection.
 *
 * `catAxisCrosses*` locates the category-axis rule on the numeric value axis;
 * `valAxisCrosses*` locates the value-axis rule on the category axis. Keeping
 * this choice in model space makes axis rules, ticks and labels share the same
 * homogeneous camera instead of moving a projected line after the fact.
 */
function threeDAxisGeometry(
  chart: ChartModel,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  orientation: 'vertical' | 'horizontal',
): ThreeDAxisGeometry {
  const { front } = projection;
  const clampFraction = (value: number): number => Math.max(0, Math.min(1, value));
  const topologyX = projection.topology.axisX === 'min' ? front.x : front.x + front.w;
  const topologyY = projection.topology.axisY === 'min' ? front.y : front.y + front.h;

  const numericCrossFraction = (): number => {
    if (chart.catAxisCrossesAt != null && Number.isFinite(chart.catAxisCrossesAt)) {
      // Percent-stacked render geometry uses the internal -100..100 scale,
      // while OOXML authors crossesAt in the underlying 0..1 value domain.
      const authored = chart.chartType.endsWith('Pct')
        ? chart.catAxisCrossesAt * 100
        : chart.catAxisCrossesAt;
      return clampFraction(axis.fraction(authored));
    }
    const crosses = chart.catAxisCrosses ?? 'autoZero';
    if (crosses === 'min') return clampFraction(axis.fraction(axis.min));
    if (crosses === 'max') return clampFraction(axis.fraction(axis.max));
    return clampFraction(axis.fraction(0));
  };
  const categoryCrossFraction = (): number | null => {
    const explicit = chart.valAxisCrossesAt;
    if (explicit != null && Number.isFinite(explicit)) {
      // Category-axis crossesAt is one-based in classic chart markup. A value
      // of 1 selects the first category position; clamp malformed/public-model
      // values before they enter camera geometry.
      return categoryPositionFraction(
        explicit - 1, categoryCount, categoryBetween, categoryReversed,
      );
    }
    const crosses = chart.valAxisCrosses;
    if (crosses !== 'min' && crosses !== 'max') return null;
    const fraction = crosses === 'max' ? 1 : 0;
    return categoryReversed ? 1 - fraction : fraction;
  };

  const numericFraction = numericCrossFraction();
  const categoryFraction = categoryCrossFraction();
  const axisX = orientation === 'horizontal'
    ? front.x + numericFraction * front.w
    : categoryFraction == null ? topologyX : front.x + categoryFraction * front.w;
  const axisY = orientation === 'vertical'
    ? front.y + front.h - numericFraction * front.h
    : categoryFraction == null ? topologyY : front.y + categoryFraction * front.h;
  const depth = projection.topology.nearDepth;
  return {
    axisX,
    axisY,
    depth,
    horizontalStart: projection.project(front.x, axisY, depth),
    horizontalEnd: projection.project(front.x + front.w, axisY, depth),
    verticalStart: projection.project(axisX, front.y + front.h, depth),
    verticalEnd: projection.project(axisX, front.y, depth),
  };
}

function cartesianAxisTicks(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  projection: ChartThreeDProjection,
  axis: NumericValueAxisPlan,
  categoryCount: number,
  categoryBetween: boolean,
  categoryReversed: boolean,
  orientation: 'vertical' | 'horizontal',
  barFamily: boolean,
  ptToPx: number,
): void {
  const { front } = projection;
  const geometry = threeDAxisGeometry(
    chart, projection, axis, categoryCount, categoryBetween, categoryReversed, orientation,
  );
  const { axisX, axisY, depth } = geometry;
  const projectedCenter = projection.project(
    front.x + front.w / 2,
    front.y + front.h / 2,
    depth,
  );
  const valueMinorTickMark = chart.valAxisMinorTickMark ?? 'none';
  const axisTickStroke = barFamily ? threeDAxisTickStroke : threeDAxisStroke;
  if (!chart.valAxisHidden && !chart.valAxisLineHidden) {
    applyThreeDStroke(ctx, axisTickStroke(
      chart.valAxisLineColor, chart.valAxisLineWidthEmu, chart.valAxisLineDash,
      ptToPx,
    ));
    const valueAnchor = (value: number) => orientation === 'horizontal'
      ? projection.project(front.x + axis.fraction(value) * front.w, axisY, depth)
      : projection.project(axisX, front.y + front.h - axis.fraction(value) * front.h, depth);
    const valueStart = orientation === 'horizontal'
      ? geometry.horizontalStart : geometry.verticalStart;
    const valueEnd = orientation === 'horizontal'
      ? geometry.horizontalEnd : geometry.verticalEnd;
    for (const value of axis.majorTicks) {
      projectedAxisTick(
        ctx, chart.valAxisMajorTickMark, valueAnchor(value), valueStart, valueEnd,
        projectedCenter, orientation === 'vertical' ? 'horizontal' : 'vertical', 'major', ptToPx,
      );
    }
    for (const value of axis.minorTicks) {
      projectedAxisTick(
        ctx, valueMinorTickMark, valueAnchor(value), valueStart, valueEnd,
        projectedCenter, orientation === 'vertical' ? 'horizontal' : 'vertical', 'minor', ptToPx,
      );
    }
  }
  if (!chart.catAxisHidden && !chart.catAxisLineHidden) {
    applyThreeDStroke(ctx, axisTickStroke(
      chart.catAxisLineColor, chart.catAxisLineWidthEmu, chart.catAxisLineDash,
      ptToPx,
    ));
    const categoryStart = orientation === 'vertical'
      ? geometry.horizontalStart : geometry.verticalStart;
    const categoryEnd = orientation === 'vertical'
      ? geometry.horizontalEnd : geometry.verticalEnd;
    const tickSkip = Math.max(1, Math.floor(chart.catAxisTickMarkSkip ?? 1));
    const drawCategoryTick = (fraction: number): void => {
      const anchor = orientation === 'vertical'
        ? projection.project(front.x + fraction * front.w, axisY, depth)
        : projection.project(axisX, front.y + fraction * front.h, depth);
      projectedAxisTick(
        ctx, chart.catAxisMajorTickMark, anchor, categoryStart, categoryEnd,
        projectedCenter, orientation === 'vertical' ? 'vertical' : 'horizontal', 'major', ptToPx,
      );
    };
    if (barFamily) {
      const boundaries = categoryGridlineFractions(categoryCount, categoryBetween);
      for (let index = 0; index < boundaries.length; index += tickSkip) {
        drawCategoryTick(categoryReversed ? 1 - boundaries[index] : boundaries[index]);
      }
    } else {
      for (let index = 0; index < categoryCount; index += tickSkip) {
        drawCategoryTick(categoryPositionFraction(
          index, categoryCount, categoryBetween, categoryReversed,
        ));
      }
    }
    const minorUnit = chart.catAxisMinorUnit;
    if (chart.catAxisMinorTickMark && chart.catAxisMinorTickMark !== 'none'
      && minorUnit != null && Number.isFinite(minorUnit) && minorUnit > 0) {
      const majorUnit = chart.catAxisMajorUnit != null
        && Number.isFinite(chart.catAxisMajorUnit) && chart.catAxisMajorUnit > 0
        ? chart.catAxisMajorUnit
        : tickSkip;
      const boundedMinorCount = Math.min(512, Math.ceil(categoryCount / minorUnit));
      for (let index = 1; index < boundedMinorCount; index++) {
        const categoryPosition = index * minorUnit;
        if (!(categoryPosition < categoryCount)) break;
        if (Math.abs(categoryPosition / majorUnit - Math.round(categoryPosition / majorUnit)) < 1e-9) continue;
        const fraction = categoryPositionFraction(
          categoryPosition, categoryCount, categoryBetween, categoryReversed,
        );
        const anchor = orientation === 'vertical'
          ? projection.project(front.x + fraction * front.w, axisY, depth)
          : projection.project(axisX, front.y + fraction * front.h, depth);
        projectedAxisTick(
          ctx, chart.catAxisMinorTickMark, anchor, categoryStart, categoryEnd,
          projectedCenter, orientation === 'vertical' ? 'vertical' : 'horizontal', 'minor', ptToPx,
        );
      }
    }
  }
  ctx.setLineDash([]);
}

function renderCartesian(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
): boolean {
  if (!chart.threeD) return false;
  const bars = chart.chartType === 'clusteredBar'
    || chart.chartType === 'clusteredBarH'
    || chart.chartType.startsWith('stackedBar');
  const horizontal = chart.chartType.endsWith('H') || chart.chartType.includes('BarH');
  const stacked = chart.chartType.startsWith('stacked');
  const depthArranged = bars && !stacked && chart.threeD.barGrouping === 'standard';
  const { plot, legend, legendMeasure } = titleAndPlot(
    ctx, chart, rect, ptToPx, horizontal ? 'horizontal' : 'vertical', shapeRotationDeg,
  );
  let projection = planChartThreeDProjection(chart.threeD, plot, {
    // Office places bar/column prisms in a compact depth box, while line and
    // area series planes span a substantially deeper Z scene. Projection and
    // view angles remain identical; only the authored family geometry differs.
    // A clustered/stacked bar group is one compact Z slab because its series
    // are separated on the category axis (or accumulated on the value axis).
    // A `standard` bar group instead uses `<c:serAx>` as a real third axis.
    // ECMA-376 §21.2.2.41 defines the authored model percentage but does not
    // prescribe Office's final projected auto-fit. For the default
    // h=100/depth=100/elevation=15/rotation=20/perspective=30 view, measured
    // Office axes are width:height:depth = 8.1:8.1:2.6. Standard Bar uses the
    // authored ECMA FOV directly; the stronger line/area compatibility gain
    // made its convergence visibly too severe.
    sceneDepthScale: bars ? (depthArranged ? 0.65 : 0.10) : 0.40,
    perspectiveTangentGain: depthArranged ? 1 : 2,
    // MS-OE376 §2.1.1501(a) assigns omitted hPercent to Office automatic
    // scaling. Repeated Excel/PDF line/area probes resolve that automatic
    // model-space height to one third of the scene width; an authored
    // hPercent continues to own the ratio in planChartThreeDProjection.
    sceneHeightScale: !bars
      && !(chart.threeD.heightPercentAuthored ?? chart.threeD.heightPercent != null)
      ? 1 / 3
      : undefined,
  });
  if (!projection) return true;
  projection = fitChartThreeDProjectionToWallThickness(projection, chart.threeD, plot);
  const percent = chart.chartType.endsWith('Pct');
  const categories = chart.series.find(series => (series.categories?.length ?? 0) > 0)?.categories
    ?? chart.categories;
  const categoryCount = Math.max(1, categories.length, ...chart.series.map(series => series.values.length));
  const percentScales = percent
    ? planThreeDPercentCategoryScales(chart, categoryCount)
    : undefined;
  const categoryReversed = chart.catAxisOrientation === 'maxMin';
  const categoryBetween = chart.catAxisCrossBetween === 'between';
  const dispBlanks = chart.dispBlanksAs ?? 'gap';
  const logarithmic = chart.valAxisLogBase != null
    && Number.isFinite(chart.valAxisLogBase)
    && chart.valAxisLogBase >= 2;
  const hasFiniteValue = (seriesIndex: number, categoryIndex: number): boolean => {
    const raw = chart.series[seriesIndex]?.values[categoryIndex];
    return raw != null && Number.isFinite(raw) && (!logarithmic || raw > 0)
      || (raw == null && (stacked || dispBlanks === 'zero'));
  };
  const valueAt = (seriesIndex: number, categoryIndex: number): number => {
    return threeDPlottedValue(chart, seriesIndex, categoryIndex, percentScales);
  };
  let dataMin = 0;
  // Percent-stacked extents follow the data signs. Starting every percent
  // chart at +100 makes an all-negative stack reserve an empty positive half
  // of the scene and moves the zero plane away from its authored edge.
  let dataMax = 0;
  if (stacked) {
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
      let positive = 0;
      let negative = 0;
      for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
        const value = valueAt(seriesIndex, categoryIndex);
        if (value >= 0) positive = threeDSafeAdd(positive, value);
        else negative = threeDSafeAdd(negative, value);
      }
      dataMin = Math.min(dataMin, negative);
      dataMax = Math.max(dataMax, positive);
    }
    if (percent) {
      // Match the shared 2-D percent-stacked contract: each occupied sign gets
      // a complete 100% side of the axis, not merely the largest category's
      // signed share. Authored min/max still win in axisPlan().
      dataMin = dataMin < 0 ? -100 : 0;
      dataMax = dataMax > 0 ? 100 : 0;
      if (dataMin === 0 && dataMax === 0) dataMax = 1;
    }
  } else {
    const finiteValues = chart.series.flatMap(series => series.values).filter(
      (value): value is number => value != null
        && Number.isFinite(value)
        && (!logarithmic || value > 0),
    );
    const extent = finiteDataExtent(
      finiteValues,
      logarithmic ? { min: 1, max: 10 } : { min: 0, max: 1 },
    );
    dataMin = logarithmic ? extent.min : Math.min(0, extent.min);
    dataMax = logarithmic ? extent.max : Math.max(0, extent.max);
  }
  const projectedValueAxisStart = horizontal
    ? projection.project(
        projection.front.x,
        projection.topology.axisY === 'min'
          ? projection.front.y
          : projection.front.y + projection.front.h,
        projection.topology.nearDepth,
      )
    : projection.project(
        projection.topology.axisX === 'min'
          ? projection.front.x
          : projection.front.x + projection.front.w,
        projection.front.y,
        projection.topology.nearDepth,
      );
  const projectedValueAxisEnd = horizontal
    ? projection.project(
        projection.front.x + projection.front.w,
        projection.topology.axisY === 'min'
          ? projection.front.y
          : projection.front.y + projection.front.h,
        projection.topology.nearDepth,
      )
    : projection.project(
        projection.topology.axisX === 'min'
          ? projection.front.x
          : projection.front.x + projection.front.w,
        projection.front.y + projection.front.h,
        projection.topology.nearDepth,
      );
  const projectedValueAxisLenPt = Math.hypot(
    projectedValueAxisEnd.x - projectedValueAxisStart.x,
    projectedValueAxisEnd.y - projectedValueAxisStart.y,
  ) / ptToPx;
  const axis = axisPlan(
    chart, dataMin, dataMax, projectedValueAxisLenPt,
    percent, horizontal ? 'horizontal' : 'vertical',
  );
  // A column/bar datum is zero only while zero belongs to the value-axis
  // domain. Logarithmic axes and explicitly positive/negative-only bounds
  // place the visible datum at the nearest authored edge; projecting zero in
  // log space otherwise creates a many-decades-off-screen cuboid.
  const visibleAxisValue = (value: number): number => {
    if (!Number.isFinite(value)) return axis.min;
    return Math.max(axis.min, Math.min(axis.max, value));
  };
  if (!bars) {
    walls(
      ctx, chart, projection, axis, horizontal ? 'horizontal' : 'vertical',
      categoryCount, categoryBetween, categoryReversed, ptToPx,
    );
  }
  const { front } = projection;
  const seriesCount = Math.max(1, chart.series.length);
  const deferredDataLabels: Array<() => void> = [];
  // Classic parser caches may carry thousands of indexed overrides. Resolve
  // them once per series; point-by-point Array.find would turn a bounded 10k
  // Canvas paint into quadratic work.
  const pointOverrides = chart.series.map(series => new Map<number, ChartDataPointOverride>(
    series.dataPointOverrides?.map(override => [override.idx, override]) ?? [],
  ));
  const labelOverrides = chart.series.map(series => new Map<number, ChartDataLabelOverride>(
    series.dataLabelOverrides?.map(override => [override.idx, override]) ?? [],
  ));
  if (bars) {
    // ECMA-376 §21.2.2.77 keeps two non-stacked 3-D arrangements distinct:
    // `clustered` series are adjacent on the category axis and share one depth
    // interval, while `standard` series reuse one category footprint and
    // occupy separate slots on `<c:serAx>`. Public models that omit the new
    // carrier retain the pre-existing clustered fallback.
    const clusterDepth = projection.prismInterval(0, 1, true);
    const primitives: Array<{
      x: number; y: number; width: number; height: number;
      nearDepth: number; farDepth: number;
      categoryIndex: number; seriesIndex: number; color: string;
      fillPaint: Fill | null | undefined;
      lineFill: ChartModel['plotAreaLineFill'] | null | undefined;
      shape: string; baseCoord: number; endCoord: number; endScale: number;
      baseScale: number; omitBaseCap: boolean; omitEndCap: boolean;
      outline: boolean; outlineColor: string; outlineWidth: number;
      outlineDash: string; outlineCustomDash: ChartModel['plotAreaLineCustomDash'];
      outlineCap: CanvasLineCap; outlineJoin: CanvasLineJoin;
      labelValue: number;
      plottedLabelValue: number;
    }> = [];
    const positiveBase = new Array(categoryCount).fill(0) as number[];
    const negativeBase = new Array(categoryCount).fill(0) as number[];
    const interval = (horizontal ? front.h : front.w) / categoryCount;
    const gapWidth = chart.barGapWidth != null
      && Number.isFinite(chart.barGapWidth)
      && chart.barGapWidth >= 0
      ? chart.barGapWidth
      : 150;
    for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
      const series = chart.series[seriesIndex];
      const depthInterval = depthArranged
        ? projection.prismInterval(seriesIndex, seriesCount, false)
        : clusterDepth;
      const clusterSlot = planThreeDBarClusterSlot(
        interval,
        gapWidth,
        depthArranged ? 0 : seriesIndex,
        depthArranged ? 1 : seriesCount,
        stacked || depthArranged,
      );
      for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
        if (!hasFiniteValue(seriesIndex, categoryIndex)) continue;
        const value = valueAt(seriesIndex, categoryIndex);
        const pointOverride = pointOverrides[seriesIndex].get(categoryIndex);
        const paint = threeDDatumPaint(
          chart, series, pointOverride, categoryIndex, seriesIndex,
        );
        const color = paint.fill === null ? 'transparent' : paint.color;
        const outline = paint.lineFill !== null
          && (paint.lineColor != null || paint.lineFill !== undefined);
        const base = stacked ? (value >= 0 ? positiveBase[categoryIndex] : negativeBase[categoryIndex]) : 0;
        const end = threeDSafeAdd(base, value);
        if (stacked) {
          if (value >= 0) positiveBase[categoryIndex] = end;
          else negativeBase[categoryIndex] = end;
        }
        const visibleBase = visibleAxisValue(base);
        const visibleEnd = visibleAxisValue(end);
        const shape = series.threeDShape ?? chart.threeD.shape ?? 'box';
        const tapered = shape === 'cone' || shape === 'pyramid';
        const toMax = shape === 'coneToMax' || shape === 'pyramidToMax';
        const ordinaryTaperScale = (coordinate: number): number => {
          if (!tapered) return 1;
          // A taper is linear in the same model coordinate used to paint the
          // value axis. Raw-value interpolation is only equivalent on a linear
          // axis; on a logarithmic axis it turns a clipped cone into the wrong
          // frustum. Zero is outside a log domain, so the authored datum begins
          // with a full ring at the first visible positive axis value.
          const modelBaseValue = logarithmic && !(base > 0) ? axis.min : base;
          const modelBase = axis.fraction(modelBaseValue);
          const modelEnd = axis.fraction(end);
          const modelCoordinate = axis.fraction(coordinate);
          const span = modelEnd - modelBase;
          if (span === 0 || !Number.isFinite(span)
            || !Number.isFinite(modelCoordinate - modelBase)) {
            return coordinate === modelBaseValue ? 1 : 0;
          }
          return Math.max(0, Math.min(1, 1 - (modelCoordinate - modelBase) / span));
        };
        const toMaxScale = (coordinate: number): number => {
          if (!toMax) return 1;
          const bound = end >= base ? axis.max : axis.min;
          const baseline = logarithmic ? axis.min : 0;
          const modelBound = axis.fraction(bound);
          const modelBaseline = axis.fraction(baseline);
          const modelCoordinate = axis.fraction(coordinate);
          const span = Math.abs(modelBound - modelBaseline);
          if (!(span > 0) || ![modelBound, modelBaseline, modelCoordinate].every(Number.isFinite)) {
            return threeDToMaxScale(coordinate, axis.min, axis.max);
          }
          return Math.max(0, Math.min(1, Math.abs(modelBound - modelCoordinate) / span));
        };
        // An authored axis clips the original solid. Preserve its parametric
        // cross-section at each clip plane instead of regenerating a full cone
        // or pyramid inside the visible interval.
        const baseScale = toMax
          ? toMaxScale(visibleBase)
          : ordinaryTaperScale(visibleBase);
        const endScale = toMax
          ? toMaxScale(visibleEnd)
          : ordinaryTaperScale(visibleEnd);
        if (horizontal) {
          const x0 = front.x + axis.fraction(visibleBase) * front.w;
          const x1 = front.x + axis.fraction(visibleEnd) * front.w;
          const categorySlot = categoryReversed ? categoryCount - 1 - categoryIndex : categoryIndex;
          const top = front.y + categorySlot * interval + clusterSlot.offset;
          primitives.push({
            x: Math.min(x0, x1), y: top,
            width: Math.abs(x1 - x0), height: clusterSlot.size,
            nearDepth: depthInterval.near, farDepth: depthInterval.far,
            categoryIndex, seriesIndex, color, fillPaint: paint.fill,
            lineFill: paint.lineFill,
            shape,
            baseCoord: x0, endCoord: x1,
            baseScale, endScale, omitBaseCap: false, omitEndCap: false,
            outline,
            outlineColor: paint.lineColor ?? 'rgba(0,0,0,0.42)',
            outlineWidth: paint.lineWidthEmu != null
              ? threeDMeshOutlineWidthPx(paint.lineWidthEmu, ptToPx)
              : 0.75 * ptToPx / PT_TO_PX,
            outlineDash: paint.lineDash ?? 'solid',
            outlineCustomDash: paint.lineCustomDash,
            outlineCap: paint.lineCap,
            outlineJoin: paint.lineJoin,
            labelValue: percent ? value / 100 : value,
            plottedLabelValue: end,
          });
        } else {
          const y0 = front.y + front.h - axis.fraction(visibleBase) * front.h;
          const y1 = front.y + front.h - axis.fraction(visibleEnd) * front.h;
          const categorySlot = categoryReversed ? categoryCount - 1 - categoryIndex : categoryIndex;
          const left = front.x + categorySlot * interval + clusterSlot.offset;
          primitives.push({
            x: left, y: Math.min(y0, y1),
            width: clusterSlot.size, height: Math.abs(y1 - y0),
            nearDepth: depthInterval.near, farDepth: depthInterval.far,
            categoryIndex, seriesIndex, color, fillPaint: paint.fill,
            lineFill: paint.lineFill,
            shape,
            baseCoord: y0, endCoord: y1,
            baseScale, endScale, omitBaseCap: false, omitEndCap: false,
            outline,
            outlineColor: paint.lineColor ?? 'rgba(0,0,0,0.42)',
            outlineWidth: paint.lineWidthEmu != null
              ? threeDMeshOutlineWidthPx(paint.lineWidthEmu, ptToPx)
              : 0.75 * ptToPx / PT_TO_PX,
            outlineDash: paint.lineDash ?? 'solid',
            outlineCustomDash: paint.lineCustomDash,
            outlineCap: paint.lineCap,
            outlineJoin: paint.lineJoin,
            labelValue: percent ? value / 100 : value,
            plottedLabelValue: end,
          });
        }
      }
    }
    if (stacked) {
      // Remove only genuinely shared internal cross-sections. This converts a
      // compatible stack into one continuous mesh while preserving caps when
      // shapes/scales differ or an authored axis clips away the neighbour.
      const primitivesByCategory = bucketThreeDStackItems(primitives, categoryCount);
      for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
        const categoryPrimitives = primitivesByCategory[categoryIndex];
        for (const sign of [-1, 1] as const) {
          const segments = categoryPrimitives
            .filter(item => Math.sign(item.labelValue) === sign
              && !isTransparentPaint(item.color)
              && Math.abs(item.endCoord - item.baseCoord) > 1e-9)
            .sort((a, b) => a.seriesIndex - b.seriesIndex);
          for (let index = 0; index + 1 < segments.length; index++) {
            const lower = segments[index];
            const upper = segments[index + 1];
            const tolerance = 1e-8 * Math.max(
              1, Math.abs(lower.endCoord), Math.abs(upper.baseCoord),
            );
            if (lower.shape !== upper.shape
              || Math.abs(lower.endCoord - upper.baseCoord) > tolerance
              || Math.abs(lower.endScale - upper.baseScale) > 1e-9
              || lower.nearDepth !== upper.nearDepth
              || lower.farDepth !== upper.farDepth) continue;
            lower.omitEndCap = true;
            upper.omitBaseCap = true;
          }
        }
        // The first positive and negative slabs meet at the zero plane. When
        // both are opaque and have the same cross-section this plane is inside
        // one continuous signed stack, so neither solid may paint a cap there.
        // A noFill neighbour is deliberately excluded: it cannot occlude the
        // opaque slab's exposed cap.
        const positive = categoryPrimitives.find(item =>
          item.labelValue > 0 && !isTransparentPaint(item.color));
        const negative = categoryPrimitives.find(item =>
          item.labelValue < 0 && !isTransparentPaint(item.color));
        if (positive && negative) {
          const tolerance = 1e-8 * Math.max(
            1, Math.abs(positive.baseCoord), Math.abs(negative.baseCoord),
          );
          if (positive.shape === negative.shape
            && Math.abs(positive.baseCoord - negative.baseCoord) <= tolerance
            && Math.abs(positive.baseScale - negative.baseScale) <= 1e-9
            && positive.nearDepth === negative.nearDepth
            && positive.farDepth === negative.farDepth) {
            positive.omitBaseCap = true;
            negative.omitBaseCap = true;
          }
        }
      }
    }
    walls(
      ctx, chart, projection, axis, horizontal ? 'horizontal' : 'vertical',
      categoryCount, categoryBetween, categoryReversed, ptToPx,
    );
    // Build one scene list and sort every visible face by camera-space depth.
    // A logical series index is not a view depth after rotation, and painting
    // a complete cuboid at a time can still put one of its far faces over a
    // nearer neighbour. The compact face sort is the shared painter contract.
    const meshBudget: ScenePrimitiveBudget = {
      remaining: MAX_PROJECTED_STROKE_PRIMITIVES,
      exceeded: false,
    };
    const sceneFaces = primitives.flatMap(item => {
      const faces = shapeMeshFaces(
        projection,
        item.shape,
        horizontal,
        item.x, item.y, item.width, item.height,
        item.baseCoord, item.endCoord,
        item.nearDepth, item.farDepth, item.color,
        item.baseScale,
        item.endScale,
        item.omitBaseCap,
        item.omitEndCap,
        item.outline && item.outlineColor ? {
          color: item.outlineColor,
          width: item.outlineWidth,
          dash: drawingmlLineDashArray(
            item.outlineCustomDash, item.outlineDash, item.outlineWidth,
          ),
          cap: item.outlineCap,
          join: item.outlineJoin,
        } : undefined,
        meshBudget,
      );
      applyScenePaint(ctx, faces, 'fill', item.fillPaint, shapeRotationDeg);
      applyScenePaint(ctx, faces, 'outline', item.lineFill, shapeRotationDeg);
      return faces;
    });
    if (meshBudget.exceeded) {
      paintThreeDTooManyDataPoints(ctx, rect);
      return true;
    }
    for (const item of sortProjectedSceneFaces(sceneFaces)) paintSceneFace(ctx, item);
    for (const item of primitives) {
      const series = chart.series[item.seriesIndex];
      const anchor = horizontal
        ? projection.project(
          item.endCoord,
          item.y + item.height / 2,
          (item.nearDepth + item.farDepth) / 2,
        )
        : projection.project(
          item.x + item.width / 2,
          item.endCoord,
          (item.nearDepth + item.farDepth) / 2,
        );
      const labelOverride = labelOverrides[item.seriesIndex].get(item.categoryIndex);
      if (wantsThreeDDataLabel(chart, series, labelOverride)) {
        deferredDataLabels.push(() => drawThreeDDataLabel(
          ctx, chart, series, item.seriesIndex, item.categoryIndex,
          item.labelValue, anchor, rect, ptToPx, 0, undefined, labelOverride,
          't', anchor, true, chart.valAxisDisplayUnits, axis.max, item.plottedLabelValue,
          shapeRotationDeg,
        ));
      }
    }
  } else {
    // Stacked line/area geometry is derived once in source-series order. The
    // lower and upper boundaries then live in the same x/y/z scene as every
    // other primitive; painting a raw series from zero would not be a stacked
    // surface even if its projection were geometrically coherent.
    const stackedLower = chart.series.map(() => new Array<number>(categoryCount).fill(0));
    const stackedUpper = chart.series.map(() => new Array<number>(categoryCount).fill(0));
    if (stacked) {
      const positive = new Array<number>(categoryCount).fill(0);
      const negative = new Array<number>(categoryCount).fill(0);
      for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
        for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
          const value = valueAt(seriesIndex, categoryIndex);
          const base = value >= 0 ? positive[categoryIndex] : negative[categoryIndex];
          stackedLower[seriesIndex][categoryIndex] = base;
          stackedUpper[seriesIndex][categoryIndex] = threeDSafeAdd(base, value);
          if (value >= 0) positive[categoryIndex] = threeDSafeAdd(positive[categoryIndex], value);
          else negative[categoryIndex] = threeDSafeAdd(negative[categoryIndex], value);
        }
      }
    }
    const seriesAverageDepth = (seriesIndex: number): number => {
      const z = projection.seriesDepth(seriesIndex, seriesCount, stacked);
      let sum = 0;
      let count = 0;
      for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
        if (!hasFiniteValue(seriesIndex, categoryIndex)) continue;
        const upperValue = stacked
          ? stackedUpper[seriesIndex][categoryIndex]
          : valueAt(seriesIndex, categoryIndex);
        const x = front.x + categoryPositionFraction(
          categoryIndex, categoryCount, categoryBetween, categoryReversed,
        ) * front.w;
        const y = front.y + front.h - axis.fraction(visibleAxisValue(upperValue)) * front.h;
        sum += projection.cameraDepth(x, y, z);
        count++;
      }
      return count > 0 ? sum / count : Number.NEGATIVE_INFINITY;
    };
    const seriesOrder = chart.series.map((_, seriesIndex) => seriesIndex).sort((a, b) =>
      seriesAverageDepth(a) - seriesAverageDepth(b) || b - a
    );
    const seriesPaints = chart.series.map((series, seriesIndex) =>
      threeDDatumPaint(chart, series, undefined, seriesIndex, seriesIndex)
    );
    const sceneCommands: Array<ProjectedSceneFace & {
      layer: 0 | 1;
      paint: () => void;
    }> = [];
    const foregroundMarkers: Array<() => void> = [];
    let strokeBudgetExceeded = false;
    for (const seriesIndex of seriesOrder) {
      if (strokeBudgetExceeded) break;
      const series = chart.series[seriesIndex];
      const seriesPaint = seriesPaints[seriesIndex];
      const color = seriesPaint.fill === null ? 'transparent' : seriesPaint.color;
      const depthInterval = stacked
        ? projection.prismInterval(0, 1, true)
        : projection.prismInterval(seriesIndex, seriesCount, false);
      const depthProbeX = front.x + front.w / 2;
      const depthProbeY = front.y + front.h / 2;
      // Lines and annotations sit on the camera-facing ridge of the solid
      // area slab.  Ordinary line charts use the same plane without painting
      // the slab faces.
      const z = projection.cameraDepth(depthProbeX, depthProbeY, depthInterval.near)
        >= projection.cameraDepth(depthProbeX, depthProbeY, depthInterval.far)
        ? depthInterval.near : depthInterval.far;
      const points: Array<Point | null> = [];
      const lowerPoints: Array<Point | null> = [];
      const pointDepths: Array<number | null> = [];
      const lowerDepths: Array<number | null> = [];
      const sourceUpperValues: Array<number | null> = [];
      const sceneXs: Array<number | null> = [];
      const upperYs: Array<number | null> = [];
      const lowerYs: Array<number | null> = [];
      const upperFractions: Array<number | null> = [];
      const lowerFractions: Array<number | null> = [];
      for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
        if (!hasFiniteValue(seriesIndex, categoryIndex)) {
          points.push(null);
          lowerPoints.push(null);
          pointDepths.push(null);
          lowerDepths.push(null);
          sourceUpperValues.push(null);
          sceneXs.push(null);
          upperYs.push(null);
          lowerYs.push(null);
          upperFractions.push(null);
          lowerFractions.push(null);
          continue;
        }
        const upperValue = stacked
          ? stackedUpper[seriesIndex][categoryIndex]
          : valueAt(seriesIndex, categoryIndex);
        const lowerValue = stacked
          ? stackedLower[seriesIndex][categoryIndex]
          : 0;
        const sceneX = front.x + categoryPositionFraction(
          categoryIndex, categoryCount, categoryBetween, categoryReversed,
        ) * front.w;
        const upperFraction = axis.fraction(upperValue);
        // Zero is not part of a logarithmic domain.  An ordinary area datum
        // still has a visible base at the lower plot edge, so represent that
        // clipped model boundary as fraction zero rather than asking log(0).
        const rawLowerFraction = axis.fraction(lowerValue);
        const lowerFraction = Number.isFinite(rawLowerFraction)
          ? rawLowerFraction
          : lowerValue <= axis.min ? 0 : 1;
        const upperInside = Number.isFinite(upperFraction)
          && upperFraction >= 0 && upperFraction <= 1;
        const visibleUpperFraction = Number.isFinite(upperFraction)
          ? Math.max(0, Math.min(1, upperFraction))
          : upperValue <= axis.min ? 0 : 1;
        const upperY = front.y + front.h
          - visibleUpperFraction * front.h;
        const lowerY = front.y + front.h
          - Math.max(0, Math.min(1, lowerFraction)) * front.h;
        points.push(
          chart.chartType.toLowerCase().includes('area') || upperInside
            ? projection.project(sceneX, upperY, z)
            : null,
        );
        lowerPoints.push(projection.project(sceneX, lowerY, z));
        pointDepths.push(upperInside
          ? projection.cameraDepth(sceneX, upperY, z)
          : null);
        lowerDepths.push(projection.cameraDepth(sceneX, lowerY, z));
        sourceUpperValues.push(upperValue);
        sceneXs.push(sceneX);
        upperYs.push(upperY);
        lowerYs.push(lowerY);
        upperFractions.push(upperFraction);
        lowerFractions.push(lowerFraction);
      }
      const runs: Array<{
        upper: Point[];
        lower: Point[];
        upperDepths: number[];
        lowerDepths: number[];
        indices: number[];
        sceneXs: number[];
        upperYs: number[];
        lowerYs: number[];
        upperFractions: number[];
        lowerFractions: number[];
      }> = [];
      let run: (typeof runs)[number] | null = null;
      for (let categoryIndex = 0; categoryIndex < points.length; categoryIndex++) {
        const point = points[categoryIndex];
        const lower = lowerPoints[categoryIndex];
        if (!point || !lower) {
          if (dispBlanks === 'gap') {
            if (run) runs.push(run);
            run = null;
          }
          continue;
        }
        run ??= {
          upper: [], lower: [], upperDepths: [], lowerDepths: [], indices: [],
          sceneXs: [], upperYs: [], lowerYs: [], upperFractions: [], lowerFractions: [],
        };
        run.upper.push(point);
        run.lower.push(lower);
        run.upperDepths.push(pointDepths[categoryIndex] ?? 0);
        run.lowerDepths.push(lowerDepths[categoryIndex] ?? 0);
        run.indices.push(categoryIndex);
        run.sceneXs.push(sceneXs[categoryIndex] ?? 0);
        run.upperYs.push(upperYs[categoryIndex] ?? 0);
        run.lowerYs.push(lowerYs[categoryIndex] ?? 0);
        run.upperFractions.push(upperFractions[categoryIndex] ?? 0);
        run.lowerFractions.push(lowerFractions[categoryIndex] ?? 0);
      }
      if (run) runs.push(run);
      const areaStrokeRuns: ProjectedStrokePoint[][] = [];
      const areaPaintFaces: SceneFace[] = [];
      if (chart.chartType.toLowerCase().includes('area')) {
        for (const item of runs) {
          let strokeRun: ProjectedStrokePoint[] | null = null;
          for (let index = 0; index + 1 < item.upper.length; index++) {
            const pieces = clipAxisFractionBand(
              item.lowerFractions[index], item.lowerFractions[index + 1],
              item.upperFractions[index], item.upperFractions[index + 1],
            );
            for (let pieceIndex = 0; pieceIndex < pieces.length; pieceIndex++) {
              const piece = pieces[pieceIndex];
              const sceneX0 = item.sceneXs[index]
                + (item.sceneXs[index + 1] - item.sceneXs[index]) * piece.startT;
              const sceneX1 = item.sceneXs[index]
                + (item.sceneXs[index + 1] - item.sceneXs[index]) * piece.endT;
              const lowerY0 = front.y + front.h - piece.lowerStart * front.h;
              const lowerY1 = front.y + front.h - piece.lowerEnd * front.h;
              const upperY0 = front.y + front.h - piece.upperStart * front.h;
              const upperY1 = front.y + front.h - piece.upperEnd * front.h;
              const ridgeStart = {
                ...projection.project(sceneX0, upperY0, z),
                cameraDepth: projection.cameraDepth(sceneX0, upperY0, z),
                cameraWeight: projection.cameraProjectionWeight(sceneX0, upperY0, z),
              };
              const ridgeEnd = {
                ...projection.project(sceneX1, upperY1, z),
                cameraDepth: projection.cameraDepth(sceneX1, upperY1, z),
                cameraWeight: projection.cameraProjectionWeight(sceneX1, upperY1, z),
              };
              const continues = strokeRun != null && Math.hypot(
                strokeRun.at(-1)!.x - ridgeStart.x,
                strokeRun.at(-1)!.y - ridgeStart.y,
              ) <= 1e-8;
              if (!continues) {
                if (strokeRun && strokeRun.length >= 2) areaStrokeRuns.push(strokeRun);
                strokeRun = [ridgeStart, ridgeEnd];
              } else {
                strokeRun!.push(ridgeEnd);
              }
              const faces = areaStripFaces(
                projection,
                sceneX0, sceneX1,
                lowerY0, lowerY1,
                upperY0, upperY1,
                depthInterval.near, depthInterval.far, color,
                index === 0 && pieceIndex === 0 && piece.startT === 0,
                index + 2 === item.upper.length
                  && pieceIndex + 1 === pieces.length && piece.endT === 1,
              );
              for (const areaFace of faces) {
                if (sceneCommands.length >= MAX_PROJECTED_STROKE_PRIMITIVES) {
                  strokeBudgetExceeded = true;
                  break;
                }
                sceneCommands.push({
                  points: areaFace.points,
                  cameraDepth: areaFace.cameraDepth,
                  cameraDepths: areaFace.cameraDepths,
                  cameraWeights: areaFace.cameraWeights,
                  layer: 0,
                  paint: () => paintSceneFace(ctx, areaFace),
                });
                areaPaintFaces.push(areaFace);
              }
              if (strokeBudgetExceeded) break;
            }
            if (strokeBudgetExceeded) break;
          }
          if (strokeRun && strokeRun.length >= 2) areaStrokeRuns.push(strokeRun);
          if (strokeBudgetExceeded) break;
        }
        applyScenePaint(ctx, areaPaintFaces, 'fill', seriesPaint.fill, shapeRotationDeg);
      }
      const lineStrokeRuns: Array<{
        path: ProjectedStrokePoint[];
        modelPath: Point[];
        ownerIndex: number;
        startClipped: boolean;
        endClipped: boolean;
        dashOffset: number;
      }> = [];
      if (!chart.chartType.toLowerCase().includes('area')) {
        type ModelLinePoint = { x: number; fraction: number; ownerIndex: number };
        const modelRuns: ModelLinePoint[][] = [];
        let modelRun: ModelLinePoint[] | null = null;
        for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
          const value = sourceUpperValues[categoryIndex];
          const fraction = value == null ? Number.NaN : axis.fraction(value);
          if (value == null || !Number.isFinite(fraction)) {
            if (dispBlanks === 'gap' && modelRun) {
              modelRuns.push(modelRun);
              modelRun = null;
            }
            continue;
          }
          modelRun ??= [];
          modelRun.push({
            x: front.x + categoryPositionFraction(
              categoryIndex, categoryCount, categoryBetween, categoryReversed,
            ) * front.w,
            fraction,
            ownerIndex: categoryIndex,
          });
        }
        if (modelRun) modelRuns.push(modelRun);
        const projectModelPoint = (point: ModelLinePoint): ProjectedStrokePoint => {
          const visibleFraction = Math.max(0, Math.min(1, point.fraction));
          const sceneY = front.y + front.h - visibleFraction * front.h;
          return {
            ...projection.project(point.x, sceneY, z),
            cameraDepth: projection.cameraDepth(point.x, sceneY, z),
            cameraWeight: projection.cameraProjectionWeight(point.x, sceneY, z),
          };
        };
        const finishVisibleRun = (current: (typeof lineStrokeRuns)[number] | null) => {
          if (current && current.path.length >= 2) lineStrokeRuns.push(current);
        };
        for (const source of modelRuns) {
          if (source.length < 2) continue;
          const sampled: ModelLinePoint[] = [source[0]];
          for (let index = 0; index + 1 < source.length; index++) {
            const p0 = source[index - 1] ?? source[index];
            const p1 = source[index];
            const p2 = source[index + 1];
            const p3 = source[index + 2] ?? p2;
            if (series.smooth !== true || source.length < 3) {
              sampled.push(p2);
              continue;
            }
            const c1 = {
              x: p1.x + (p2.x - p0.x) / 6,
              fraction: p1.fraction + (p2.fraction - p0.fraction) / 6,
            };
            const c2 = {
              x: p2.x - (p3.x - p1.x) / 6,
              fraction: p2.fraction - (p3.fraction - p1.fraction) / 6,
            };
            for (let step = 1; step <= 12; step++) {
              const t = step / 12;
              const u = 1 - t;
              sampled.push({
                x: u * u * u * p1.x + 3 * u * u * t * c1.x
                  + 3 * u * t * t * c2.x + t * t * t * p2.x,
                fraction: u * u * u * p1.fraction
                  + 3 * u * u * t * c1.fraction
                  + 3 * u * t * t * c2.fraction + t * t * t * p2.fraction,
                ownerIndex: p2.ownerIndex,
              });
            }
          }
          let visible: (typeof lineStrokeRuns)[number] | null = null;
          let traversedLength = 0;
          for (let index = 0; index + 1 < sampled.length; index++) {
            const startModel = sampled[index];
            const endModel = sampled[index + 1];
            const projectUnclipped = (point: ModelLinePoint): Point => projection.project(
              point.x,
              front.y + front.h - point.fraction * front.h,
              z,
            );
            const projectedStart = projectUnclipped(startModel);
            const projectedEnd = projectUnclipped(endModel);
            const projectedLength = Math.hypot(
              projectedEnd.x - projectedStart.x,
              projectedEnd.y - projectedStart.y,
            );
            const modelLength = Number.isFinite(projectedLength)
              ? projectedLength
              : Math.hypot(
                endModel.x - startModel.x,
                (endModel.fraction - startModel.fraction) * front.h,
              );
            const clipped = clipAxisFractionSegment(
              startModel.fraction, endModel.fraction,
            );
            if (!clipped || clipped.endT - clipped.startT <= 1e-12) {
              finishVisibleRun(visible);
              visible = null;
              traversedLength += modelLength;
              continue;
            }
            const at = (t: number): ModelLinePoint => ({
              x: startModel.x + (endModel.x - startModel.x) * t,
              fraction: startModel.fraction
                + (endModel.fraction - startModel.fraction) * t,
              ownerIndex: endModel.ownerIndex,
            });
            const projectedClipStart = projectUnclipped(at(clipped.startT));
            const clippedPrefixLength = Math.hypot(
              projectedClipStart.x - projectedStart.x,
              projectedClipStart.y - projectedStart.y,
            );
            const start = projectModelPoint(at(clipped.startT));
            const end = projectModelPoint(at(clipped.endT));
            const modelStart = at(clipped.startT);
            const modelEnd = at(clipped.endT);
            const toModelPoint = (point: ModelLinePoint): Point => ({
              x: point.x,
              y: front.y + front.h
                - Math.max(0, Math.min(1, point.fraction)) * front.h,
            });
            const continues = visible != null && Math.hypot(
              visible.path.at(-1)!.x - start.x,
              visible.path.at(-1)!.y - start.y,
            ) <= 1e-8 && visible.ownerIndex === endModel.ownerIndex;
            if (!continues) {
              finishVisibleRun(visible);
              visible = {
                path: [start, end],
                modelPath: [toModelPoint(modelStart), toModelPoint(modelEnd)],
                ownerIndex: endModel.ownerIndex,
                startClipped: index > 0 || clipped.startT > 0,
                endClipped: index + 1 < sampled.length - 1 || clipped.endT < 1,
                dashOffset: traversedLength + (Number.isFinite(clippedPrefixLength)
                  ? clippedPrefixLength : modelLength * clipped.startT),
              };
            } else {
              visible!.path.push(end);
              visible!.modelPath.push(toModelPoint(modelEnd));
              visible!.endClipped = index + 1 < sampled.length - 1 || clipped.endT < 1;
            }
            traversedLength += modelLength;
          }
          finishVisibleRun(visible);
        }
      }
      const areaFamily = chart.chartType.toLowerCase().includes('area');
      const authoredAreaLine = series.lineHidden != null
        || series.lineColor != null
        || series.lineWidthEmu != null
        || series.chartexStyle?.lineHidden != null
        || series.chartexStyle?.lineColors?.some(Boolean)
        || series.chartexStyle?.lineWidthEmu != null
        || series.chartexStyle?.lineDash != null
        || series.chartexStyle?.lineCap != null
        || series.chartexStyle?.lineJoin != null
        || seriesPaint.lineColor != null
        || seriesPaint.lineFill !== undefined;
      if (areaFamily ? seriesPaint.lineFill !== null && authoredAreaLine : true) {
        const pointLinePaints = new Map<number, ThreeDDatumPaint>();
        const paintForLineRun = (ownerIndex: number): ThreeDDatumPaint => {
          const point = pointOverrides[seriesIndex].get(ownerIndex);
          if (!threeDPointOwnsLineSegment(point)) return seriesPaint;
          let paint = pointLinePaints.get(ownerIndex);
          if (!paint) {
            paint = threeDDatumPaint(chart, series, point, ownerIndex, seriesIndex);
            pointLinePaints.set(ownerIndex, paint);
          }
          return paint;
        };
        const strokeFaceGroups = new Map<ThreeDDatumPaint, SceneFace[]>();
        const lineRibbonBudget: ScenePrimitiveBudget = {
          remaining: MAX_PROJECTED_STROKE_PRIMITIVES,
          exceeded: false,
        };
        const pushStrokeGeometry = (
          path: readonly ProjectedStrokePoint[],
          paint: ThreeDDatumPaint,
          startCap: CanvasLineCap,
          endCap: CanvasLineCap = startCap,
          dashOffset = 0,
          modelPath?: readonly Point[],
        ) => {
          if (paint.lineFill === null) return;
          // Automatic 3-D line/area edges use the darker Chart Style line
          // role. Direct point line paint owns only its incoming segment.
          const automaticStrokeBase = paint === seriesPaint ? color : paint.color;
          const strokeStyle = paint.lineColor ?? scaleHexColor(automaticStrokeBase, 0.70);
          const lineWidth = paint.lineWidthEmu
            ? Math.max(0.5, paint.lineWidthEmu / EMU_PER_PT) * ptToPx
            : areaFamily
              ? 0.75 * ptToPx
              : Math.max(1, 2 * ptToPx);
          const lineDash = drawingmlLineDashArray(
            paint.lineCustomDash, paint.lineDash, lineWidth,
          );
          const strokeOptions = {
            width: lineWidth,
            dash: lineDash,
            dashOffset,
            lineCap: paint.lineCap,
            startCap,
            endCap,
            lineJoin: paint.lineJoin,
          };
          // ECMA Line3D is a depth-bearing ribbon. Tessellate the same bounded
          // dash/cap/join polygon in the chart's model X/Y plane, extrude it
          // through the series interval, then hand every face to the existing
          // camera/depth/material pipeline. Area ridges remain ordinary edge
          // strokes because their slab already supplies the depth faces.
          if (!areaFamily && modelPath && modelPath.length >= 2) {
            const modelStrokes = buildProjectedStrokePrimitives(
              modelPath.map(point => ({ ...point, cameraDepth: 0, cameraWeight: 1 })),
              strokeOptions,
            );
            if (modelStrokes == null) {
              strokeBudgetExceeded = true;
              return;
            }
            const group = strokeFaceGroups.get(paint) ?? [];
            for (const stroke of modelStrokes) {
              const mesh = buildThreeDLineRibbonMesh({
                outline: stroke.points,
                nearDepth: depthInterval.near,
                farDepth: depthInterval.far,
              });
              if (!mesh) continue;
              const faces = projectThreeDMesh(
                projection, mesh, strokeStyle, undefined, lineRibbonBudget,
              ).map(face => ({ ...face, paintRole: 'outline' as const }));
              if (lineRibbonBudget.exceeded) {
                strokeBudgetExceeded = true;
                return;
              }
              group.push(...faces);
              for (const face of faces) {
                sceneCommands.push({
                  points: face.points,
                  cameraDepth: face.cameraDepth,
                  cameraDepths: face.cameraDepths,
                  cameraWeights: face.cameraWeights,
                  layer: 1,
                  paint: () => paintSceneFace(ctx, face),
                });
              }
            }
            strokeFaceGroups.set(paint, group);
            return;
          }
          const strokes = buildProjectedStrokePrimitives(path, strokeOptions);
          if (strokes == null) {
            strokeBudgetExceeded = true;
            return;
          }
          if (sceneCommands.length + strokes.length > MAX_PROJECTED_STROKE_PRIMITIVES) {
            strokeBudgetExceeded = true;
            return;
          }
          for (const stroke of strokes) {
            const strokeFace: SceneFace = {
              points: stroke.points,
              color: strokeStyle,
              paintRole: 'outline',
              shade: 0,
              cameraDepth: stroke.cameraDepth,
              cameraDepths: stroke.cameraDepths,
              cameraWeights: stroke.cameraWeights,
              outline: false,
            };
            const group = strokeFaceGroups.get(paint) ?? [];
            group.push(strokeFace);
            strokeFaceGroups.set(paint, group);
            sceneCommands.push({
              points: stroke.points,
              cameraDepth: stroke.cameraDepth,
              cameraDepths: stroke.cameraDepths,
              cameraWeights: stroke.cameraWeights,
              layer: 1,
              paint: () => paintSceneFace(ctx, strokeFace),
            });
          }
        };
        if (areaFamily) {
          for (const path of areaStrokeRuns) {
            pushStrokeGeometry(path, seriesPaint, seriesPaint.lineCap);
          }
        } else {
          for (const item of lineStrokeRuns) {
            const paint = paintForLineRun(item.ownerIndex);
            pushStrokeGeometry(
              item.path,
              paint,
              item.startClipped ? 'butt' : paint.lineCap,
              item.endClipped ? 'butt' : paint.lineCap,
              item.dashOffset,
              item.modelPath,
            );
          }
        }
        for (const [paint, faces] of strokeFaceGroups) {
          applyScenePaint(ctx, faces, 'outline', paint.lineFill, shapeRotationDeg);
        }
      }
      const seriesMarkersVisible = (areaFamily
        ? series.showMarker === true || seriesHasMarkerDetail(series)
        : series.showMarker === true) && series.markerSymbol !== 'none';
      if ((chart.chartType.toLowerCase().includes('line') || areaFamily)
        && (seriesMarkersVisible || hasVisiblePointMarkerOverride(series))) {
        for (let categoryIndex = 0; categoryIndex < points.length; categoryIndex++) {
          const point = points[categoryIndex];
          if (!point) continue;
          const override = pointOverrides[seriesIndex].get(categoryIndex);
          const symbol = effectiveMarkerSymbol(
            series, override, 'circle', seriesMarkersVisible,
          );
          if (symbol === 'none') continue;
          const sizePt = override?.markerSize ?? series.markerSize ?? 5;
          const markerFill = markerFillColorFor(
            series,
            override,
            categoryIndex,
            series.color ?? PALETTE[seriesIndex % PALETTE.length],
          );
          const markerFillPaint = markerFillPaintFor(series, override, categoryIndex);
          const markerLine = override?.markerLine ?? series.markerLine ?? series.lineColor ?? series.color
            ?? PALETTE[seriesIndex % PALETTE.length];
          const markerLineWidth = (override?.markerLineWidthEmu ?? series.markerLineWidthEmu) != null
            ? Math.max(0.25, (override?.markerLineWidthEmu ?? series.markerLineWidthEmu ?? 0) / EMU_PER_PT * ptToPx)
            : Math.max(0.75, series.lineWidthEmu != null
              ? series.lineWidthEmu / EMU_PER_PT * ptToPx
              : ptToPx);
          foregroundMarkers.push(() => paintThreeDMarker(
              ctx, point, symbol, Math.max(2, sizePt) * ptToPx,
              markerFill === '00000000' ? 'transparent' : `#${markerFill}`,
              `#${markerLine}`, markerLineWidth,
              markerFillPaint, shapeRotationDeg, ptToPx,
            ));
        }
      }
      for (let categoryIndex = 0; categoryIndex < points.length; categoryIndex++) {
        const point = points[categoryIndex];
        if (!point) continue;
        const sourceValue = valueAt(seriesIndex, categoryIndex);
        const override = pointOverrides[seriesIndex].get(categoryIndex);
        const markerSize = override?.markerSize ?? series.markerSize ?? 5;
        const labelOverride = labelOverrides[seriesIndex].get(categoryIndex);
        if (wantsThreeDDataLabel(chart, series, labelOverride)) {
          deferredDataLabels.push(() => drawThreeDDataLabel(
            ctx, chart, series, seriesIndex, categoryIndex,
            percent ? sourceValue / 100 : sourceValue,
            point, rect, ptToPx,
            series.showMarker === true || override?.markerSymbol != null
              ? markerSize * ptToPx / 2
              : 0,
            undefined,
            labelOverride,
            't', point, true, chart.valAxisDisplayUnits, axis.max,
            stacked ? stackedUpper[seriesIndex][categoryIndex] : sourceValue,
            shapeRotationDeg,
          ));
          }
        }
      }
    // Area faces and explicit stroke polygons participate in one local-overlap
    // depth graph. Average series depth is insufficient when two projected
    // surfaces cross at only one side of the chart.
    if (strokeBudgetExceeded) {
      // The source/mesh preflight rejects ordinary oversized charts before
      // any paint. This late guard handles adversarial sub-pixel dash
      // expansion. Do not cover a noFill chart with an invented white panel;
      // preserve the host background and report the bounded rejection.
      paintThreeDTooManyDataPoints(ctx, rect);
      return true;
    }
    for (const command of sortProjectedSceneFaces(sceneCommands)) command.paint();
    // Markers are annotations on the completed series geometry. Keeping them
    // out of the face/segment depth list prevents a later adjacent segment
    // from cutting through its own endpoint marker.
    for (const paintMarker of foregroundMarkers) paintMarker();
  }
  frontAxes(
    ctx, chart, projection, axis, categoryCount, categoryBetween, categoryReversed,
    ptToPx, horizontal ? 'horizontal' : 'vertical',
  );
  cartesianAxisTicks(
    ctx, chart, projection, axis, categoryCount, categoryBetween, categoryReversed,
    horizontal ? 'horizontal' : 'vertical', bars, ptToPx,
  );
  drawThreeDSeriesAxis(
    ctx,
    chart,
    rect,
    projection,
    axis,
    categoryCount,
    categoryBetween,
    categoryReversed,
    horizontal ? 'horizontal' : 'vertical',
    bars,
    ptToPx,
  );
  const resolvedAxisGeometry = threeDAxisGeometry(
    chart,
    projection,
    axis,
    categoryCount,
    categoryBetween,
    categoryReversed,
    horizontal ? 'horizontal' : 'vertical',
  );
  const edgeGeometry = (
    kind: 'value' | 'category',
    position: string | null | undefined,
  ): ThreeDAxisGeometry => {
    if (position !== 'low' && position !== 'high') return resolvedAxisGeometry;
    const depth = projection.topology.nearDepth;
    let axisX = resolvedAxisGeometry.axisX;
    let axisY = resolvedAxisGeometry.axisY;
    if ((kind === 'value') === horizontal) {
      const minY = front.y;
      const maxY = front.y + front.h;
      const minPoint = projection.project(front.x + front.w / 2, minY, depth);
      const maxPoint = projection.project(front.x + front.w / 2, maxY, depth);
      const lowY = minPoint.y >= maxPoint.y ? minY : maxY;
      const highY = lowY === minY ? maxY : minY;
      axisY = position === 'low' ? lowY : highY;
    } else {
      const minX = front.x;
      const maxX = front.x + front.w;
      const minPoint = projection.project(minX, front.y + front.h / 2, depth);
      const maxPoint = projection.project(maxX, front.y + front.h / 2, depth);
      const lowX = minPoint.x <= maxPoint.x ? minX : maxX;
      const highX = lowX === minX ? maxX : minX;
      axisX = position === 'low' ? lowX : highX;
    }
    return {
      axisX, axisY, depth,
      horizontalStart: projection.project(front.x, axisY, depth),
      horizontalEnd: projection.project(front.x + front.w, axisY, depth),
      verticalStart: projection.project(axisX, front.y + front.h, depth),
      verticalEnd: projection.project(axisX, front.y, depth),
    };
  };
  const valueLabelGeometry = edgeGeometry('value', chart.valAxisTickLabelPos);
  const categoryLabelGeometry = edgeGeometry('category', chart.catAxisTickLabelPos);
  const valueFontPx = chartTextFontSizePx(chart.valAxisFontSizeHpt, ptToPx) ?? 9 * ptToPx;
  ctx.font = `${chart.valAxisFontItalic ? 'italic ' : ''}${chart.valAxisFontBold ? 'bold ' : ''}${valueFontPx}px ${fontFamily(chart.valAxisFontFace)}`;
  ctx.fillStyle = chart.valAxisFontColor ? `#${chart.valAxisFontColor}` : '#595959';
  ctx.textAlign = horizontal ? 'center' : 'right';
  ctx.textBaseline = horizontal ? 'top' : 'middle';
  if (!chart.valAxisHidden && chart.valAxisTickLabelPos !== 'none') {
    const { axisX, axisY, depth: axisDepth } = valueLabelGeometry;
    const sceneMidpoint = projection.project(front.x + front.w / 2, front.y + front.h / 2, axisDepth);
    const valueAxisStart = horizontal
      ? valueLabelGeometry.horizontalStart : valueLabelGeometry.verticalStart;
    const valueAxisEnd = horizontal
      ? valueLabelGeometry.horizontalEnd : valueLabelGeometry.verticalEnd;
    const labelNormal = screenAnnotationOutward(
      valueAxisStart,
      valueAxisEnd,
      sceneMidpoint,
      horizontal ? 'vertical' : 'horizontal',
    );
    ctx.textAlign = Math.abs(labelNormal.x) < 0.2 ? 'center' : labelNormal.x < 0 ? 'right' : 'left';
    ctx.textBaseline = Math.abs(labelNormal.y) < 0.2 ? 'middle' : labelNormal.y < 0 ? 'bottom' : 'top';
    const labelOffset = projectedAxisTickLabelOffsetPx(
      chart.valAxisMajorTickMark, chart.valAxisLineHidden, ptToPx, 5,
    );
    const displayUnitDivisor = chart.valAxisDisplayUnits?.divisor;
    for (const value of axis.majorTicks) {
      const point = horizontal
        ? projection.project(front.x + axis.fraction(value) * front.w, axisY, axisDepth)
        : projection.project(axisX, front.y + front.h - axis.fraction(value) * front.h, axisDepth);
      ctx.fillText(formatChartValWithCode(
        percent
          ? value / 100
          : displayUnitDivisor != null
              && Number.isFinite(displayUnitDivisor)
              && displayUnitDivisor > 0
            ? value / displayUnitDivisor
            : value,
        percent ? chart.valAxisFormatCode ?? '0%' : chart.valAxisFormatCode,
        chart.date1904,
      ), point.x + labelNormal.x * labelOffset, point.y + labelNormal.y * labelOffset);
    }
  }
  const categoryFontPx = chartTextFontSizePx(chart.catAxisFontSizeHpt, ptToPx) ?? 9 * ptToPx;
  ctx.font = `${chart.catAxisFontItalic ? 'italic ' : ''}${chart.catAxisFontBold ? 'bold ' : ''}${categoryFontPx}px ${fontFamily(chart.catAxisFontFace)}`;
  ctx.fillStyle = chart.catAxisFontColor ? `#${chart.catAxisFontColor}` : '#595959';
  if (!chart.catAxisHidden && chart.catAxisTickLabelPos !== 'none') {
    const labelOffset = categoryLabelOffsetPx(
      projectedAxisTickLabelOffsetPx(
        chart.catAxisMajorTickMark, chart.catAxisLineHidden, ptToPx, 6,
      ),
      chart.catAxisLabelOffsetPercent,
    );
    const formattedCategories = Array.from({ length: categoryCount }, (_, index) =>
      formatCategoryLabel(
        String(categories[index] ?? index + 1),
        chart.catAxisFormatCode,
        chart.date1904,
      )
    );
    let rotation = authoredCategoryRotation(chart);
    if (rotation == null) {
      rotation = 0;
      if (!horizontal && categoryCount > 1) {
        let minimumSpacing = Number.POSITIVE_INFINITY;
        let previous: Point | null = null;
        const categoryAxisY = categoryLabelGeometry.axisY;
        for (let index = 0; index < categoryCount; index++) {
          const fraction = categoryPositionFraction(index, categoryCount, categoryBetween, categoryReversed);
          const current = projection.project(
            front.x + fraction * front.w,
            categoryAxisY,
            projection.topology.nearDepth,
          );
          if (previous) minimumSpacing = Math.min(
            minimumSpacing,
            Math.hypot(current.x - previous.x, current.y - previous.y),
          );
          previous = current;
        }
        const widest = Math.max(0, ...formattedCategories.map(label => ctx.measureText(label).width));
        // Office automatically angles dense cartesian 3-D category labels.
        // The boundary corpus repeats a -45° choice once a horizontal label
        // no longer fits between adjacent projected tick positions. Explicit
        // bodyPr@rot, including an authored zero, remains authoritative.
        if (widest > minimumSpacing * 0.9) rotation = -Math.PI / 4;
      }
    }
    const labelInterval = Math.max(1, Math.floor(chart.catAxisTickLabelSkip ?? 1));
    for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex += labelInterval) {
      const fraction = categoryPositionFraction(
        categoryIndex, categoryCount, categoryBetween, categoryReversed,
      );
      const { axisX, axisY, depth: axisDepth } = categoryLabelGeometry;
      const point = horizontal
        ? projection.project(
          axisX,
          front.y + categoryPositionFraction(
            categoryIndex, categoryCount, categoryBetween, categoryReversed,
          ) * front.h,
          axisDepth,
        )
        : projection.project(front.x + fraction * front.w, axisY, axisDepth);
      if (horizontal) {
        const axisMidpoint = projection.project(axisX, front.y + front.h / 2, axisDepth);
        const sceneMidpoint = projection.project(front.x + front.w / 2, front.y + front.h / 2, axisDepth);
        const onScreenLeft = axisMidpoint.x <= sceneMidpoint.x;
        ctx.textAlign = onScreenLeft ? 'right' : 'left';
        ctx.textBaseline = 'middle';
        ctx.fillText(
          formattedCategories[categoryIndex],
          point.x + (onScreenLeft ? -labelOffset : labelOffset),
          point.y,
        );
      } else {
        const sceneMidpoint = projection.project(
          front.x + front.w / 2,
          front.y + front.h / 2,
          axisDepth,
        );
        const outward = screenAnnotationOutward(
          categoryLabelGeometry.horizontalStart,
          categoryLabelGeometry.horizontalEnd,
          sceneMidpoint,
          'vertical',
        );
        drawAngledCategoryLabel(
          ctx,
          formattedCategories[categoryIndex],
          point,
          rotation,
          horizontal,
          outward.y < 0 ? -1 : 1,
          labelOffset,
        );
      }
    }
  }
  drawThreeDAxisTitles(ctx, chart, rect, plot, horizontal, ptToPx);
  // Data labels are chart foreground annotations. Flush them only after all
  // walls, faces, lines, markers, axes, ticks and tick labels, so later scene
  // primitives cannot cross a label belonging to an earlier series.
  for (const drawLabel of deferredDataLabels) drawLabel();
  simpleLegend(ctx, chart, legend, ptToPx, false, legendMeasure, shapeRotationDeg);
  return true;
}

function renderPie(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg: number,
): boolean {
  if (!chart.threeD || chart.chartType !== 'pie') return false;
  const series = chart.series[0];
  if (!series) return true;
  const values = series.values.flatMap((value, index) =>
    value != null && Number.isFinite(value) ? [{ index, value: Math.abs(value) }] : []
  );
  let maxMagnitude = 0;
  for (const item of values) maxMagnitude = Math.max(maxMagnitude, item.value);
  if (!(maxMagnitude > 0)) return true;
  const scaledTotal = values.reduce((sum, item) => sum + item.value / maxMagnitude, 0);
  if (!(scaledTotal > 0) || !Number.isFinite(scaledTotal)) return true;
  const { plot, legend, legendMeasure } = titleAndPlot(
    ctx, chart, rect, ptToPx, 'radial', shapeRotationDeg,
  );
  // Pie is a cylindrical sector in X/Z with thickness on Y. Its authored
  // depth/gap fields do not control the radial solid (the depth=100/2000
  // Office references are pixel-identical), but rotX/rotY/perspective must use
  // the same homogeneous camera as every cartesian mesh.
  let projection = planChartThreeDProjection({
    ...chart.threeD,
    // Unlike the cartesian compatibility view, an omitted radial yaw starts
    // at zero: the first-slice ray remains at twelve o'clock. An authored
    // rotY still rotates the complete solid and its slice seams.
    rotationY: chart.threeD.rotationY ?? 0,
    // MS-OE376 §2.1.1501(b): pie3D hPercent scales the default pie
    // thickness. It is not the cartesian scene-height ratio used by the
    // other 3-D families.
    heightPercent: undefined,
    depthPercent: 100,
  }, plot, {
    sceneDepthScale: 1,
    // radius=.45W and thickness=.30r occupy .135W on model Y. A .15W scene
    // box tightly fits the actual cylinder when hPercent is omitted, instead
    // of shrinking it to fit an unused full-height cartesian cuboid.
    sceneHeightScale: 0.15,
  });
  if (!projection) return true;
  const pointOverrides = new Map<number, ChartDataPointOverride>(
    series.dataPointOverrides?.map(override => [override.idx, override]) ?? [],
  );
  let maxExplosion = 0;
  for (const override of pointOverrides.values()) {
    if (override.explosion != null && Number.isFinite(override.explosion)) {
      maxExplosion = Math.max(maxExplosion, Math.max(0, Math.min(100, override.explosion)) / 100);
    }
  }
  const { scene } = projection;
  const radius = Math.min(
    scene.w * 0.45 / (1 + maxExplosion),
    projection.modelDepth * 0.45 / (1 + maxExplosion),
    scene.h / 0.45,
  );
  if (!(radius > 0)) return true;
  const centerX = scene.x + scene.w / 2;
  const centerY = scene.y + scene.h / 2;
  const centerDepth = 0.5;
  // Office's default pie wall is about .29r. The camera now decides which
  // side is visible; no screen-down wall or abs(rotX) special case remains.
  const thickness = radius * 0.30 * pieThreeDThicknessMultiplier(chart.threeD);
  // rotY is already part of the camera and must not be added to the seam a
  // second time. Slice angles are resolved in the X/Z model plane below.
  let cumulativeFraction = 0;
  const slices: Array<{
    index: number;
    start: number;
    end: number;
    color: string;
    fillPaint: Fill | null | undefined;
    value: number;
    percentValue: number;
    centerX: number;
    centerDepth: number;
    segments: number;
    mesh: ThreeDMesh;
    lineColor: string | null;
    lineFill: ChartModel['plotAreaLineFill'] | null | undefined;
    lineWidthEmu: number | null;
    lineDash: string;
    lineCustomDash: ChartModel['plotAreaLineCustomDash'];
    lineCap: CanvasLineCap;
    lineJoin: CanvasLineJoin;
  }> = [];
  const labelOverrides = new Map<number, ChartDataLabelOverride>(
    series.dataLabelOverrides?.map(override => [override.idx, override]) ?? [],
  );
  // Keep the projected circumference below roughly four pixels per facet,
  // bounded to 128 segments for synchronous Canvas work. A fixed 32-facet
  // ring is visibly polygonal on large/zoomed charts and makes an otherwise
  // constant-width outline appear stepped at every facet boundary.
  const pieRoundSegments = Math.max(48, Math.min(
    128, Math.ceil(Math.PI * 2 * radius / 4),
  ));
  for (const item of values) {
    const percentValue = (item.value / maxMagnitude) / scaledTotal;
    const angles = threeDPieSliceAngles(
      chart.firstSliceAngle, cumulativeFraction, percentValue,
    );
    const pointOverride = pointOverrides.get(item.index);
    const paint = threeDDatumPaint(chart, series, pointOverride, item.index, 0);
    const middle = angles.middle;
    const explosion = pointOverride?.explosion != null && Number.isFinite(pointOverride.explosion)
      ? Math.max(0, Math.min(100, pointOverride.explosion)) / 100
      : 0;
    const sliceCenterX = centerX + Math.cos(middle) * radius * explosion;
    const sliceCenterDepth = centerDepth
      + Math.sin(middle) * radius * explosion / projection.modelDepth;
    const segments = Math.max(2, Math.ceil(pieRoundSegments * percentValue));
    const mesh = buildThreeDPieSectorMesh({
      centerX: sliceCenterX,
      centerY,
      centerDepth: sliceCenterDepth,
      radius,
      modelDepth: projection.modelDepth,
      thickness,
      startAngle: angles.start,
      endAngle: angles.end,
      segments,
    });
    if (!mesh) {
      cumulativeFraction += percentValue;
      continue;
    }
    slices.push({
      index: item.index,
      start: angles.start,
      end: angles.end,
      color: paint.fill === null ? 'transparent' : paint.color,
      fillPaint: paint.fill,
      value: item.value,
      percentValue,
      centerX: sliceCenterX,
      centerDepth: sliceCenterDepth,
      segments,
      mesh,
      lineColor: paint.lineColor,
      lineFill: paint.lineFill,
      lineWidthEmu: paint.lineWidthEmu ?? null,
      lineDash: paint.lineDash ?? 'solid',
      lineCustomDash: paint.lineCustomDash,
      lineCap: paint.lineCap,
      lineJoin: paint.lineJoin,
    });
    cumulativeFraction += percentValue;
  }
  projection = fitChartThreeDProjectionToPoints(
    projection,
    slices.flatMap(slice => slice.mesh.vertices),
    plot,
    0.08,
  );
  const deferredLabels: Array<() => void> = [];
  const pieBudget: ScenePrimitiveBudget = {
    remaining: MAX_PROJECTED_STROKE_PRIMITIVES,
    exceeded: false,
  };
  const outlineStyleForSlice = (slice: (typeof slices)[number]): MeshOutlineStyle | undefined => {
    const width = slice.lineWidthEmu != null
      ? Math.max(0.25, slice.lineWidthEmu / EMU_PER_PT * ptToPx)
      : 0.75 * ptToPx;
    return slice.lineFill !== null
      && (slice.lineColor != null || slice.lineFill !== undefined) ? {
      color: slice.lineColor ?? 'rgba(0,0,0,0.42)',
      width,
      dash: drawingmlLineDashArray(slice.lineCustomDash, slice.lineDash, width),
      cap: slice.lineCap,
      join: slice.lineJoin,
    } : undefined;
  };
  // Fill solids remain independently colored and depth-sorted. A uniform,
  // non-exploded pie uses semantic continuous outline paths; differently
  // styled or exploded points retain independent authored solid outlines.
  const pieFillFaces = slices.flatMap(slice => {
    const faces = projectThreeDMesh(
      projection, slice.mesh, slice.color, undefined, pieBudget,
    );
    applyScenePaint(ctx, faces, 'fill', slice.fillPaint, shapeRotationDeg);
    return faces;
  });
  const pieOutlineFaces: SceneFace[] = [];
  const outlineStyles = slices.map(outlineStyleForSlice);
  const firstOutline = outlineStyles[0];
  const outlineKey = (style: MeshOutlineStyle | undefined) => style == null ? null : [
    style.color, style.width, style.dash.join(','), style.cap, style.join,
  ].join('|');
  const uniformOutline = firstOutline != null
    && slices.every(slice => slice.lineFill === undefined)
    && outlineStyles.every(style => outlineKey(style) === outlineKey(firstOutline))
    && slices.every(slice => Math.abs(slice.centerX - centerX) < 1e-9
      && Math.abs(slice.centerDepth - centerDepth) < 1e-9);
  if (uniformOutline) {
    pieOutlineFaces.push(...projectThreeDPieOutline(
      projection, slices, centerY, radius, thickness, firstOutline, pieBudget,
    ));
  } else {
    for (let index = 0; index < slices.length; index++) {
      const style = outlineStyles[index];
      if (!style) continue;
      const faces = projectThreeDMesh(
        projection, slices[index].mesh, 'transparent', style, pieBudget, true,
      );
      applyScenePaint(ctx, faces, 'outline', slices[index].lineFill, shapeRotationDeg);
      pieOutlineFaces.push(...faces);
    }
  }
  if (pieBudget.exceeded) {
    paintThreeDTooManyDataPoints(ctx, rect);
    return true;
  }
  if (uniformOutline) {
    for (const item of sortProjectedSceneFaces(pieFillFaces)) paintSceneFace(ctx, item);
    // Semantic paths already contain only visible cap/wall boundaries. Paint
    // them after all colored faces so neither adjacent coplanar slice can hide
    // half the authored screen-space width.
    for (const item of sortProjectedSceneFaces(pieOutlineFaces)) paintSceneFace(ctx, item);
  } else {
    // Exploded or independently styled slices still require inter-solid depth
    // ordering because their authored outlines can genuinely occlude peers.
    for (const item of sortProjectedSceneFaces([
      ...pieFillFaces, ...pieOutlineFaces,
    ])) paintSceneFace(ctx, item);
  }
  for (const slice of slices) {
    const middle = (slice.start + slice.end) / 2;
    const labelOverride = labelOverrides.get(slice.index);
    if (wantsThreeDDataLabel(chart, series, labelOverride)) {
      const defaults = series.seriesDataLabels;
      const fontPx = chartTextFontSizePx(
        labelOverride?.fontSizeHpt ?? defaults?.fontSizeHpt ?? chart.dataLabelFontSizeHpt,
        ptToPx,
      ) ?? 9 * ptToPx;
      ctx.font = `${labelOverride?.fontBold ?? defaults?.fontBold
        ?? chart.dataLabelFontBold ? 'bold ' : ''}${fontPx}px ${chartFontFamily(
        chart, labelOverride?.fontFace ?? defaults?.fontFace ?? chart.dataLabelFontFace,
      )}`;
      const labelText = effectiveDataLabelText({
        customText: labelOverride?.text,
        showCategory: labelOverride?.showCatName ?? defaults?.showCatName ?? false,
        showSeries: labelOverride?.showSerName ?? defaults?.showSerName ?? false,
        showValue: labelOverride?.showVal ?? defaults?.showVal ?? chart.showDataLabels,
        showPercent: labelOverride?.showPercent ?? defaults?.showPercent ?? false,
        category: series.categories?.[slice.index]
          ?? chart.categories[slice.index] ?? `${slice.index + 1}`,
        seriesName: series.name || 'Series 1',
        sourceValue: slice.value,
        percentRatio: slice.percentValue,
        formatCode: labelOverride?.formatCode
          ?? defaults?.formatCode ?? chart.dataLabelFormatCode ?? series.valFormatCode,
        separator: labelOverride?.separator ?? defaults?.separator,
        date1904: chart.date1904,
      });
      const fallbackFamily = chartFontFamily(
        chart, labelOverride?.fontFace ?? defaults?.fontFace ?? chart.dataLabelFontFace,
      );
      const richMeasure = labelOverride?.text && labelOverride.richRuns?.length
        ? resolveRichDataLabelBlock(ctx, {
          runs: labelOverride.richRuns,
          ptToPx,
          fontFamily: fallbackFamily,
          fallbackBold: labelOverride.fontBold ?? defaults?.fontBold
            ?? chart.dataLabelFontBold ?? false,
          fontFamilyForFace: face => chartFontFamily(chart, face),
        }, fontPx, `#${labelOverride.fontColor ?? defaults?.fontColor
          ?? series.labelColor ?? chart.dataLabelFontColor ?? '111111'}`)
        : null;
      const authoredPosition = labelOverride?.position ?? defaults?.position ?? chart.dataLabelPosition;
      const surfaceTop = centerY - thickness / 2;
      const surfaceBottom = centerY + thickness / 2;
      const surfaceY = projection.cameraDepth(
        slice.centerX, surfaceTop, slice.centerDepth,
      ) >= projection.cameraDepth(slice.centerX, surfaceBottom, slice.centerDepth)
        ? surfaceTop : surfaceBottom;
      let arcCapacity = 0;
      let previousArcPoint: Point | null = null;
      for (let step = 0; step <= 12; step++) {
        const arcAngle = slice.start + (slice.end - slice.start) * step / 12;
        const current = projection.project(
          slice.centerX + Math.cos(arcAngle) * radius * 0.64,
          surfaceY,
          slice.centerDepth + Math.sin(arcAngle) * radius * 0.64 / projection.modelDepth,
        );
        if (previousArcPoint) arcCapacity += Math.hypot(
          current.x - previousArcPoint.x, current.y - previousArcPoint.y,
        );
        previousArcPoint = current;
      }
      const outside = (authoredPosition == null || authoredPosition === 'bestFit')
        && (slice.percentValue === 0 || arcCapacity < (richMeasure?.width
          ?? ctx.measureText(labelText).width));
      const labelOutside = outside || authoredPosition === 'outEnd';
      const labelRadius = radius * (labelOutside ? 1.12 : 0.64);
      const label = projection.project(
        slice.centerX + Math.cos(middle) * labelRadius,
        surfaceY,
        slice.centerDepth + Math.sin(middle) * labelRadius / projection.modelDepth,
      );
      const leader = projection.project(
        slice.centerX + Math.cos(middle) * radius,
        surfaceY,
        slice.centerDepth + Math.sin(middle) * radius / projection.modelDepth,
      );
      deferredLabels.push(() => drawThreeDDataLabel(
        ctx, chart, series, 0, slice.index, slice.value, label, plot, ptToPx,
        0, slice.percentValue, labelOverride, 'ctr', leader,
        labelOutside,
        undefined, undefined, undefined, shapeRotationDeg,
      ));
    }
  }
  for (const drawLabel of deferredLabels) drawLabel();
  const categories = series.categories?.length ? series.categories : chart.categories;
  const legendPointColors = Array.from({ length: categories.length }, (_, index) => {
    const pointOverride = pointOverrides.get(index);
    const authored = pointOverride?.fillHidden === true
      ? '00000000'
      : pointOverride?.color ?? series.dataPointColors?.[index] ?? series.color;
    if (authored === '00000000') return '00000000';
    return scaleHexColor(authored ? `#${authored}` : colorFor(index), 0.80).replace(/^#/, '');
  });
  simpleLegend(ctx, {
    ...chart,
    categories,
    series: [{ ...series, categories, dataPointColors: legendPointColors }],
  }, legend, ptToPx, true, legendMeasure, shapeRotationDeg);
  return true;
}

/** Render a bounded compatibility projection for classic 3-D chart groups.
 * Returns false for ordinary 2-D charts so existing family paths remain exact. */
export function renderSimpleThreeDChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): boolean {
  if (!chart.threeD) return false;
  const supported = chart.chartType === 'pie'
    || SUPPORTED_CARTESIAN_THREE_D_TYPES.has(chart.chartType);
  if (!supported) return false;
  if (!threeDChartPaintsWithinBudget(chart)) {
    paintThreeDTooManyDataPoints(ctx, rect);
    return true;
  }
  if (renderPie(ctx, chart, rect, ptToPx, shapeRotationDeg)) return true;
  return renderCartesian(ctx, chart, rect, ptToPx, shapeRotationDeg);
}
