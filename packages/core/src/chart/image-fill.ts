import type { ImageFill } from '../types/common.js';
import type { ChartExElementStyle, ChartModel } from '../types/chart.js';
import { drawImageCropped, imageNaturalSize, srcRectHasVisibleArea } from '../image/crop.js';
import { EMU_PER_PT, PT_TO_PX } from '../units.js';
import { computeBoxWhiskerStats } from './box-whisker.js';
import {
  classicMarkerPointIsPainted,
  chartDataTableFamilyIsPainted,
  dataLabelLegendKeyCount,
  deletedLegendEntryIndices,
  effectiveMarkerSymbol,
  hasVisiblePointMarkerOverride,
  markersSuppressedByChartStyle,
  markerFillPaintFor,
  markerSymbolConsumesFill,
  seriesHasMarkerDetail,
  seriesLegendMarkerIsVisible,
  visibleBubbleSize,
} from './marker-style.js';
import {
  MAX_CANVAS_CHART_POINTS,
  MAX_CHART_IMAGE_FILL_TILES,
  MAX_CHART_MARKER_IMAGE_SOURCES,
  sourceChartStructureCount,
} from './resource-limits.js';
import { indexChartPlotGroups, markerChartTypeForPlotGroup } from './plot-groups.js';
import { chartThreeDSurfacePaint } from './style-paint.js';
import { planChartThreeDSurfacePicture } from './three-d-surface-picture-plan.js';

const SURFACE_PICTURE_FAMILIES = new Set([
  'line', 'stackedLine', 'stackedLinePct',
  'area', 'stackedArea', 'stackedAreaPct',
  'clusteredBar', 'clusteredBarH',
  'stackedBar', 'stackedBarH', 'stackedBarPct', 'stackedBarHPct',
  'surface', 'surface3D',
]);

export type ChartImageLookup = (fill: ImageFill) => CanvasImageSource | null | undefined;

/** Stable document-cache key for one decoded chart picture source. */
export function chartImageFillKey(fill: ImageFill): string {
  return JSON.stringify([
    fill.imagePath,
    fill.svgImagePath ?? null,
    fill.duotone?.clr1 ?? null,
    fill.duotone?.clr2 ?? null,
  ]);
}

let activeLookup: ChartImageLookup | undefined;

/** Install one document-owned synchronous image lookup for the duration of a
 * chart paint. Hosts warm their bounded decoded-image cache before calling the
 * synchronous chart renderer; the chart layer never fetches or decodes per
 * point. */
export function withChartImageLookup<T>(
  lookup: ChartImageLookup | undefined,
  paint: () => T,
): T {
  const previous = activeLookup;
  activeLookup = lookup;
  try {
    return paint();
  } finally {
    activeLookup = previous;
  }
}

function alignmentOffset(
  alignment: string,
  boxW: number,
  boxH: number,
  tileW: number,
  tileH: number,
): { x: number; y: number } {
  const value = alignment;
  const x = value.endsWith('r') || value === 'r'
    ? boxW - tileW
    : value === 't' || value === 'ctr' || value === 'b'
      ? (boxW - tileW) / 2
      : 0;
  const y = value.startsWith('b') || value === 'b'
    ? boxH - tileH
    : value === 'l' || value === 'ctr' || value === 'r'
      ? (boxH - tileH) / 2
      : 0;
  return { x, y };
}

const TILE_ALIGNMENTS = new Set(['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br']);
const TILE_FLIPS = new Set(['none', 'x', 'y', 'xy']);

export interface ChartImageTileMetrics {
  alignment: string;
  tileW: number;
  tileH: number;
  offsetX: number;
  offsetY: number;
  flipX: boolean;
  flipY: boolean;
}

interface TileGeometry extends ChartImageTileMetrics {
  columns: number;
  rows: number;
  repetitions: number;
}

/** EG_FillModeProperties is a choice with no schema default. Exactly one
 * authored mode must be present; omitting it must not invent stretch. */
function imageFillModeIsPaintable(fill: ImageFill): boolean {
  return (fill.tile != null) !== (fill.stretch === true)
    && srcRectHasVisibleArea(fill.srcRect);
}

/** Return the synchronously preloaded source for a validated chart image fill.
 * Kept internal to chart modules; hosts still own all fetch/decode work. */
export function chartImageFillSource(fill: ImageFill): CanvasImageSource | null {
  if (!imageFillModeIsPaintable(fill)) return null;
  return activeLookup?.(fill) ?? null;
}

/** Resolve the shared DrawingML tile size, offset, alignment and mirroring.
 * Destination-specific repetition counts remain with each consumer. */
export function chartImageTileMetrics(
  fill: ImageFill,
  image: CanvasImageSource,
  ptToPx = PT_TO_PX,
): ChartImageTileMetrics | null {
  const tile = fill.tile;
  // CT_TileInfoProperties@flip defaults to none. Its remaining placement
  // attributes and CT_BlipFillProperties@dpi have no usable schema default.
  // Do not invent Office compatibility semantics when those facts are absent.
  // A zero dpi requests embedded image metadata, which Canvas image sources do
  // not expose, so that case also remains fail-closed.
  if (!tile) return null;
  const { algn, tx, ty, sx, sy } = tile;
  const flip = tile.flip ?? 'none';
  if (!algn || !TILE_ALIGNMENTS.has(algn) || !TILE_FLIPS.has(flip)
    || !Number.isFinite(tx) || !Number.isFinite(ty)
    || !(fill.dpi != null && Number.isFinite(fill.dpi) && fill.dpi > 0)
    || !(Number.isFinite(sx) && (sx as number) > 0)
    || !(Number.isFinite(sy) && (sy as number) > 0)) return null;
  const natural = imageNaturalSize(image);
  if (!(Number.isFinite(natural.w) && natural.w > 0)
    || !(Number.isFinite(natural.h) && natural.h > 0)) return null;
  const cssPixelsPerImagePixel = 96 / fill.dpi * (ptToPx / PT_TO_PX);
  const tileW = natural.w * (sx as number) * cssPixelsPerImagePixel;
  const tileH = natural.h * (sy as number) * cssPixelsPerImagePixel;
  if (!(tileW > 0) || !(tileH > 0)) return null;
  return {
    alignment: algn,
    tileW,
    tileH,
    offsetX: (tx as number) / EMU_PER_PT * ptToPx,
    offsetY: (ty as number) / EMU_PER_PT * ptToPx,
    flipX: flip === 'x' || flip === 'xy',
    flipY: flip === 'y' || flip === 'xy',
  };
}

/** Place the authored tile-grid origin in one destination's local coordinates. */
export function chartImageTileOrigin(
  metrics: ChartImageTileMetrics,
  width: number,
  height: number,
): { x: number; y: number } {
  const anchor = alignmentOffset(
    metrics.alignment,
    width,
    height,
    metrics.tileW,
    metrics.tileH,
  );
  return { x: anchor.x + metrics.offsetX, y: anchor.y + metrics.offsetY };
}

function imageTileGeometry(
  fill: ImageFill,
  image: CanvasImageSource,
  w: number,
  h: number,
  ptToPx: number,
): TileGeometry | null {
  const metrics = chartImageTileMetrics(fill, image, ptToPx);
  if (!metrics) return null;
  const { tileW, tileH } = metrics;
  const columns = Math.ceil(w / tileW) + 2;
  const rows = Math.ceil(h / tileH) + 2;
  const repetitions = columns * rows;
  if (!Number.isSafeInteger(repetitions)) return null;
  return { ...metrics, columns, rows, repetitions };
}

/** Exact Canvas image-draw work for one picture fill at the destination size.
 * A missing/unsupported authored recipe paints nothing and therefore costs 0. */
export function chartImageFillPaintWork(
  fill: ImageFill,
  lookup: ChartImageLookup | undefined,
  w: number,
  h: number,
  ptToPx = PT_TO_PX,
): number {
  if (!(w > 0) || !(h > 0)) return 0;
  if (!imageFillModeIsPaintable(fill)) return 0;
  const image = lookup?.(fill);
  if (!image) return 0;
  if (!fill.tile) return 1;
  const geometry = imageTileGeometry(fill, image, w, h, ptToPx);
  return geometry && geometry.repetitions <= MAX_CHART_IMAGE_FILL_TILES
    ? geometry.repetitions
    : 0;
}

/** Monotonic work upper bound for a consumer whose destination rectangle is
 * conservatively estimated. If the estimate exceeds the per-marker tile cap,
 * smaller real consumers can still paint, so charge the cap rather than zero. */
export function chartImageFillPaintWorkUpperBound(
  fill: ImageFill,
  lookup: ChartImageLookup | undefined,
  w: number,
  h: number,
  ptToPx = PT_TO_PX,
): number {
  if (!(w > 0) || !(h > 0)) return 0;
  if (!imageFillModeIsPaintable(fill)) return 0;
  const image = lookup?.(fill);
  if (!image) return 0;
  if (!fill.tile) return 1;
  const geometry = imageTileGeometry(fill, image, w, h, ptToPx);
  return geometry ? Math.min(geometry.repetitions, MAX_CHART_IMAGE_FILL_TILES) : 0;
}

/** Unique picture fills reachable by marker consumers. Family suppression,
 * point validity and direct precedence match the painters so hosts never fetch
 * images for markers that cannot be drawn. */
interface ChartMarkerImageFillResult {
  fills: ImageFill[];
  sourceLimitExceeded: boolean;
}

function collectChartMarkerImageFillResult(chart: ChartModel): ChartMarkerImageFillResult {
  const sourceCount = sourceChartStructureCount(chart);
  if (sourceCount > MAX_CANVAS_CHART_POINTS) {
    return { fills: [], sourceLimitExceeded: false };
  }
  const fills = new Map<string, ImageFill>();
  let sourceLimitExceeded = false;
  const add = (fill: unknown) => {
    if (!fill || typeof fill !== 'object' || (fill as ImageFill).fillType !== 'image') return;
    const image = fill as ImageFill;
    if (!imageFillModeIsPaintable(image)) return;
    const key = chartImageFillKey(image);
    if (fills.has(key)) return;
    if (fills.size >= MAX_CHART_MARKER_IMAGE_SOURCES) {
      sourceLimitExceeded = true;
      return;
    }
    fills.set(key, image);
  };
  const selectedStylePaint = (
    style: ChartExElementStyle | null | undefined,
    index: number,
  ) => {
    const paints = style?.fillPaints;
    if (!paints?.length) return undefined;
    return paints[(style?.fillColorIndex ?? index) % paints.length];
  };
  const styleImageDecision = (
    style: ChartExElementStyle | null | undefined,
    index: number,
  ): ImageFill | null | undefined => {
    if (!style) return undefined;
    if (style.fillHidden === true) return style.fillNoStyle ? undefined : null;
    const paint = selectedStylePaint(style, index);
    if (paint?.fillType === 'image') return paint;
    if (paint != null || style.fillPaintAuthored === true) return null;
    const colors = style.fillColors;
    return colors?.length && colors[(style.fillColorIndex ?? index) % colors.length]
      ? null
      : undefined;
  };
  const frameImageDecision = (
    directFill: ChartModel['chartFill'] | ChartModel['plotAreaFill'],
    directColor: string | null | undefined,
    hidden: boolean | null | undefined,
    paintAuthored: boolean | null | undefined,
    linked: ChartExElementStyle | null | undefined,
  ): ImageFill | null | undefined => {
    if (hidden === true) return null;
    if (directFill?.fillType === 'image') return directFill;
    if (directFill != null || directColor != null || paintAuthored === true) return null;
    if (linked?.fillNoStyle === true) return undefined;
    return styleImageDecision(linked, 0);
  };
  const chartAreaImage = frameImageDecision(
    chart.chartFill,
    // `chartStyleRoleChartArea` treats the structured/provenance fields as
    // authoritative; retain that exact ownership rule for prefetch too.
    undefined,
    chart.chartFillHidden,
    chart.chartFillPaintAuthored,
    chart.chartStyleRoles?.chartArea,
  );
  if (chartAreaImage) add(chartAreaImage);
  const plotAreaImage = frameImageDecision(
    chart.plotAreaFill,
    chart.plotAreaBg,
    chart.plotAreaFillHidden,
    chart.plotAreaFillPaintAuthored,
    chart.threeD ? chart.chartStyleRoles?.plotArea3D : chart.chartStyleRoles?.plotArea,
  );
  if (plotAreaImage) add(plotAreaImage);
  const finiteSurfaceValue = chart.series.some(series =>
    series.values.some(value => value != null && Number.isFinite(value))
  );
  const surfaceColumnCount = Math.max(
    chart.categories.length,
    ...chart.series.map(series => series.categories?.length ?? series.values.length),
  );
  const surfaceGeometryCanPaint = chart.chartType === 'surface' || chart.chartType === 'surface3D'
    ? chart.series.length >= 2 && surfaceColumnCount >= 2 && finiteSurfaceValue
    : finiteSurfaceValue;
  if (chart.threeD && SURFACE_PICTURE_FAMILIES.has(chart.chartType) && surfaceGeometryCanPaint) {
    const explicitSpan = chart.valMin != null && Number.isFinite(chart.valMin)
      && chart.valMax != null && Number.isFinite(chart.valMax)
      ? chart.valMax - chart.valMin
      : undefined;
    for (const [kind, role] of [
      ['floor', 'floor'], ['sideWall', 'wall'], ['backWall', 'wall'],
    ] as const) {
      const surface = chart.threeD[kind];
      const fill = chartThreeDSurfacePaint(chart, surface, role).fill;
      if (fill?.fillType === 'image'
        && planChartThreeDSurfacePicture(fill, surface, kind, explicitSpan)) add(fill);
    }
  }
  const plotGroupBySeries = indexChartPlotGroups(chart);
  const scatterHasNumericX = chart.series.some((series, seriesIndex) => {
    const group = plotGroupBySeries[seriesIndex];
    const family = group?.kind === 'bubble' || group?.kind === 'scatter'
      ? 'scatter'
      : series.seriesType ?? (chart.chartType === 'bubble' ? 'scatter' : chart.chartType);
    return family === 'scatter' && (series.categories ?? chart.categories).some(category =>
      Number.isFinite(Number.parseFloat(category))
    );
  });
  const linkedMarkerStyle = chart.chartStyleRoles?.dataPointMarker;
  const deletedLegendEntries = deletedLegendEntryIndices(chart);
  const chartHasCategories = chart.categories.length > 0
    || (chart.series[0]?.categories?.length ?? 0) > 0
    || chart.series.some(series => series.values.length > 0);
  for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
    const series = chart.series[seriesIndex];
    const group = plotGroupBySeries[seriesIndex];
    const isBubble = group?.kind === 'bubble'
      || (group == null && chart.chartType === 'bubble');
    const family = group?.kind === 'bubble' || group?.kind === 'scatter'
      ? 'scatter'
      : series.seriesType ?? (chart.chartType === 'bubble' ? 'scatter' : chart.chartType);
    const effectiveChartType = markerChartTypeForPlotGroup(chart.chartType, group);
    const effectiveScatterStyle = group?.scatterStyle ?? chart.scatterStyle;
    const effectiveRadarStyle = group?.radarStyle ?? chart.radarStyle;
    const markerContext = {
      chartType: effectiveChartType,
      bubbleScale: group?.bubbleScale ?? chart.bubbleScale,
      showNegativeBubbles: group?.showNegativeBubbles ?? chart.showNegativeBubbles,
    };
    const markerFamily = family === 'line' || family === 'stackedLine'
      || family === 'stackedLinePct' || family === 'area'
      || family === 'stackedArea' || family === 'stackedAreaPct'
      || family === 'scatter' || family === 'radar' || family === 'stock';
    if (!markerFamily || markersSuppressedByChartStyle(
      family, effectiveChartType, effectiveScatterStyle, effectiveRadarStyle,
    )) continue;
    const areaFamily = family === 'area' || family === 'stackedArea'
      || family === 'stackedAreaPct';
    const seriesVisible = areaFamily
      ? (series.showMarker === true || seriesHasMarkerDetail(series))
        && series.markerSymbol !== 'none'
      : family === 'stock'
        ? series.markerSymbol != null && series.markerSymbol !== 'none'
        : series.showMarker !== false && series.markerSymbol !== 'none';
    const legendDeleted = deletedLegendEntries.has(seriesIndex);
    const pointCount = Math.max(
      series.values.length, series.categories?.length ?? 0, chart.categories.length,
    );
    const labelKeyVisible = dataLabelLegendKeyCount(
      chart, series, family, pointCount, scatterHasNumericX, markerContext,
    ) > 0;
    const seriesKeyVisible = seriesLegendMarkerIsVisible(
      effectiveChartType, effectiveScatterStyle, series, effectiveRadarStyle,
    ) && ((chart.showLegend && !legendDeleted)
      || (chart.dataTable?.showKeys === true && chartDataTableFamilyIsPainted(chart.chartType)
        && chartHasCategories)
      || labelKeyVisible);
    const seriesKeySymbol = series.markerSymbol ?? (family === 'stock' ? 'none' : 'circle');
    if (seriesKeyVisible && markerSymbolConsumesFill(seriesKeySymbol)) {
      if (isBubble) {
        const seriesShape = styleImageDecision(series.chartexStyle, seriesIndex);
        if (seriesShape) add(seriesShape);
        if (seriesShape === undefined && series.color == null) {
          const linkedShape = styleImageDecision(chart.chartStyleRoles?.dataPoint, seriesIndex);
          if (linkedShape) add(linkedShape);
        }
      } else add(series.markerFillPaint);
      if (!isBubble
        && series.markerFillPaint === undefined && series.markerFill == null
        && series.markerFillPaintAuthored !== true) {
        add(selectedStylePaint(linkedMarkerStyle, seriesIndex));
      }
    }
    if (!seriesVisible && !hasVisiblePointMarkerOverride(series)) continue;
    const overrides = new Map((series.dataPointOverrides ?? []).map(point => [point.idx, point]));
    for (let index = 0; index < pointCount; index++) {
      if (!classicMarkerPointIsPainted(
        chart, series, family, index, scatterHasNumericX, markerContext,
      )) continue;
      const point = overrides.get(index);
      const symbol = effectiveMarkerSymbol(series, point, 'circle', seriesVisible);
      if (!markerSymbolConsumesFill(symbol)) continue;
      if (isBubble) {
        const effectiveBubbleSettings = {
          showNegativeBubbles: group?.showNegativeBubbles ?? chart.showNegativeBubbles,
        };
        if ((group?.bubbleScale ?? chart.bubbleScale ?? 100) <= 0
          || visibleBubbleSize(effectiveBubbleSettings, series.bubbleSizes?.[index]) == null) continue;
        // MS-OE376 §2.1.1504(b): a visible negative bubble uses the
        // application's inverted fill, not its positive direct/linked image.
        // The current wire model's alternate bubble default is solid/noFill,
        // so none of these positive image sources is reachable.
        if ((series.bubbleSizes?.[index] ?? 0) < 0) continue;
        const pointShape = styleImageDecision(point?.chartexStyle, index);
        if (pointShape) add(pointShape);
        if (pointShape !== undefined || point?.fillHidden === true || point?.color != null) continue;
        if (series.dataPointColors?.[index] != null) continue;
        const seriesShape = styleImageDecision(series.chartexStyle, index);
        if (seriesShape) add(seriesShape);
        if (seriesShape !== undefined || series.color != null) continue;
        const linkedShape = styleImageDecision(chart.chartStyleRoles?.dataPoint, index);
        if (linkedShape) add(linkedShape);
        continue;
      }
      const paint = markerFillPaintFor(series, point, index);
      add(paint);
      if (paint === undefined && point?.markerFill == null && point?.color == null
        && series.dataPointColors?.[index] == null && series.markerFill == null
        && point?.markerFillPaintAuthored !== true
        && series.markerFillPaintAuthored !== true) {
        add(selectedStylePaint(linkedMarkerStyle, seriesIndex));
      }
    }
  }
  for (let seriesIndex = 0; seriesIndex < (chart.chartexBox?.series.length ?? 0); seriesIndex++) {
    const series = chart.chartexBox!.series[seriesIndex];
    const symbol = chart.chartStyleMarkerSymbol ?? chart.chartexMarkerSymbol ?? 'circle';
    if (!markerSymbolConsumesFill(symbol)
      || !(series.showNonoutliers || series.showOutliers)) continue;
    let markerCount = 0;
    for (const values of series.valuesByCategory) {
      const stats = computeBoxWhiskerStats(values, series.quartileMethod);
      if (!stats) continue;
      if (series.showNonoutliers) markerCount += stats.inner.length;
      if (series.showOutliers) markerCount += stats.outliers.length;
    }
    if (markerCount === 0) continue;
    const styleIndex = series.chartexFormatIdx ?? seriesIndex;
    const localStyle = series.chartexStyle;
    const local = styleImageDecision(localStyle, styleIndex);
    if (local) add(local);
    if (local !== undefined || series.color != null) continue;
    const linked = styleImageDecision(
      chart.chartexDataPointMarkerStyle ?? chart.chartexDataPointStyle ?? undefined,
      styleIndex,
    );
    if (linked) add(linked);
  }
  return {
    fills: sourceLimitExceeded ? [] : [...fills.values()],
    sourceLimitExceeded,
  };
}

export function collectChartMarkerImageFills(chart: ChartModel): ImageFill[] {
  return collectChartMarkerImageFillResult(chart).fills;
}

/** Collect chart picture fills for one host render pass. Hosts retain the decoded
 * sources until that page/slide/viewport paint completes, so the count ceiling
 * applies to the aggregate rather than independently to each chart. */
export function collectChartMarkerImageFillsForCharts(
  charts: readonly ChartModel[],
): ImageFill[] {
  const fills = new Map<string, ImageFill>();
  for (const chart of charts) {
    const result = collectChartMarkerImageFillResult(chart);
    if (result.sourceLimitExceeded) return [];
    for (const fill of result.fills) {
      const key = chartImageFillKey(fill);
      if (fills.has(key)) continue;
      if (fills.size >= MAX_CHART_MARKER_IMAGE_SOURCES) return [];
      fills.set(key, fill);
    }
  }
  return [...fills.values()];
}

export function paintChartImageFill(
  ctx: CanvasRenderingContext2D,
  fill: ImageFill,
  x: number,
  y: number,
  w: number,
  h: number,
  ptToPx = PT_TO_PX,
  shapeRotationDeg = 0,
): boolean {
  const image = activeLookup?.(fill);
  if (!image || !(w > 0) || !(h > 0)) return false;
  if (!imageFillModeIsPaintable(fill)) return false;
  if (shapeRotationDeg !== 0 && fill.rotWithShape == null) return false;
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  if (fill.rotWithShape === false && shapeRotationDeg !== 0) {
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate(-shapeRotationDeg * Math.PI / 180);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }
  if (fill.alpha != null) ctx.globalAlpha *= Math.max(0, Math.min(1, fill.alpha));
  if (!fill.tile) {
    const rect = fill.fillRect;
    const dx = x + (rect?.l ?? 0) * w;
    const dy = y + (rect?.t ?? 0) * h;
    const dw = (1 - (rect?.l ?? 0) - (rect?.r ?? 0)) * w;
    const dh = (1 - (rect?.t ?? 0) - (rect?.b ?? 0)) * h;
    if (dw > 0 && dh > 0) drawImageCropped(ctx, image, fill.srcRect, dx, dy, dw, dh);
    ctx.restore();
    return dw > 0 && dh > 0;
  }

  const geometry = imageTileGeometry(fill, image, w, h, ptToPx);
  if (!geometry || geometry.repetitions > MAX_CHART_IMAGE_FILL_TILES) {
    ctx.restore();
    return false;
  }
  const {
    tileW, tileH, flipX, flipY, columns, rows,
  } = geometry;
  const origin = chartImageTileOrigin(geometry, w, h);
  const originX = x + origin.x;
  const originY = y + origin.y;
  const firstColumn = Math.floor((x - originX) / tileW) - 1;
  const firstRow = Math.floor((y - originY) / tileH) - 1;
  for (let row = firstRow; row < firstRow + rows; row++) {
    for (let column = firstColumn; column < firstColumn + columns; column++) {
      const dx = originX + column * tileW;
      const dy = originY + row * tileH;
      const mirrorX = flipX && Math.abs(column) % 2 === 1;
      const mirrorY = flipY && Math.abs(row) % 2 === 1;
      ctx.save();
      ctx.translate(dx + (mirrorX ? tileW : 0), dy + (mirrorY ? tileH : 0));
      ctx.scale(mirrorX ? -1 : 1, mirrorY ? -1 : 1);
      drawImageCropped(ctx, image, fill.srcRect, 0, 0, tileW, tileH);
      ctx.restore();
    }
  }
  ctx.restore();
  return true;
}
