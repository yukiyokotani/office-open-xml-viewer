import type { ImageFill } from '../types/common.js';
import type {
  ChartDataPointOverride, ChartExElementStyle, ChartModel, ChartSeries, ChartStockBarPaint,
} from '../types/chart.js';
import { drawImageCropped, imageNaturalSize, srcRectHasVisibleArea } from '../image/crop.js';
import { fillCanProduceVisiblePixels } from '../shape/paint.js';
import { EMU_PER_PT, PT_TO_PX } from '../units.js';
import { computeBoxWhiskerStats } from './box-whisker.js';
import {
  classicDataLabelPointIsPainted,
  classicMarkerPointIsPainted,
  chartDataTableFamilyIsPainted,
  dataLabelLegendKeyCount,
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
  deletedLegendEntryIndices,
  legendEntryIsVisible,
  legendEntryRanges,
  legendSeriesHasVisibleEntry,
} from './legend-entry-plan.js';
import {
  MAX_CANVAS_CHART_POINTS,
  MAX_CHART_IMAGE_FILL_TILES,
  MAX_CHART_MARKER_IMAGE_SOURCES,
  sourceChartStructureCount,
} from './resource-limits.js';
import { indexChartPlotGroups, markerChartTypeForPlotGroup } from './plot-groups.js';
import {
  chartStyleDirectFillDecision,
  chartStyleDirectNoFillDecision,
  chartStyleFillCascade,
  chartStockBarFillDecision,
  chartStyleFillDecision,
  chartThreeDSurfacePaint,
} from './style-paint.js';
import {
  chartDataPointStyleRole,
  chartSeriesVariesByPoint,
  rawLinkedChartStyleRole,
  withChartStyleIndexCache,
  withEffectiveChartStyleRoles,
} from './effective-style.js';
import { withSparseStyleIndexCache } from './sparse-style-index.js';
import { planChartThreeDSurfacePicture } from './three-d-surface-picture-plan.js';
import { classicDataPointFillDecision } from './classic-data-point-style.js';
import { dataLabelIsDeleted } from './data-label-style.js';
import {
  chartLabelBoxHasVisiblePaint,
  effectiveChartLabelBoxFill,
  mergeChartLabelBoxes,
} from './label-box.js';
import { effectiveLegendFrameFill } from './legend-frame.js';
import { applyPlotVisibleOnly } from './source-visibility.js';
import { withActiveChartImageFillPainter } from './image-fill-context.js';
import {
  visitChartExHierarchyBodySites,
  visitChartExHierarchyLabelSites,
} from './chart-ex-hierarchy-labels.js';
import { planWaterfallPaintSites } from './waterfall-plan.js';

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
    return withActiveChartImageFillPainter(paintChartImageFill, paint);
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

interface ChartImageTileStaticMetrics {
  alignment: string;
  tx: number;
  ty: number;
  sx: number;
  sy: number;
  dpi: number;
  flipX: boolean;
  flipY: boolean;
}

/** EG_FillModeProperties is a choice with no schema default. Exactly one
 * authored mode must be present; omitting it must not invent stretch. */
function imageFillModeIsPaintable(fill: ImageFill): boolean {
  return (fill.tile != null) !== (fill.stretch === true)
    && srcRectHasVisibleArea(fill.srcRect);
}

/** Validate the authored facts needed to derive tile geometry before an image
 * has been decoded. Hosts use the same predicate when deciding whether a tiled
 * occurrence really requires native source dimensions. */
function chartImageTileStaticMetrics(fill: ImageFill): ChartImageTileStaticMetrics | null {
  const tile = fill.tile;
  if (!tile) return null;
  const { algn, tx, ty, sx, sy } = tile;
  const flip = tile.flip ?? 'none';
  const dpi = fill.dpi;
  if (!algn || !TILE_ALIGNMENTS.has(algn) || !TILE_FLIPS.has(flip)
    || !Number.isFinite(tx) || !Number.isFinite(ty)
    || !(dpi != null && Number.isFinite(dpi) && dpi > 0)
    || !(Number.isFinite(sx) && (sx as number) > 0)
    || !(Number.isFinite(sy) && (sy as number) > 0)) return null;
  return {
    alignment: algn,
    tx: tx as number,
    ty: ty as number,
    sx: sx as number,
    sy: sy as number,
    dpi,
    flipX: flip === 'x' || flip === 'xy',
    flipY: flip === 'y' || flip === 'xy',
  };
}

/** Return the synchronously preloaded source for a validated chart image fill.
 * Kept internal to chart modules; hosts still own all fetch/decode work. */
export function chartImageFillSource(fill: ImageFill): CanvasImageSource | null {
  if (!fillCanProduceVisiblePixels(fill) || !imageFillModeIsPaintable(fill)) return null;
  return activeLookup?.(fill) ?? null;
}

/** Resolve the shared DrawingML tile size, offset, alignment and mirroring.
 * Destination-specific repetition counts remain with each consumer. */
export function chartImageTileMetrics(
  fill: ImageFill,
  image: CanvasImageSource,
  ptToPx = PT_TO_PX,
): ChartImageTileMetrics | null {
  // CT_TileInfoProperties@flip defaults to none. Its remaining placement
  // attributes and CT_BlipFillProperties@dpi have no usable schema default.
  // Do not invent Office compatibility semantics when those facts are absent.
  // A zero dpi requests embedded image metadata, which Canvas image sources do
  // not expose, so that case also remains fail-closed.
  const authored = chartImageTileStaticMetrics(fill);
  if (!authored) return null;
  const natural = imageNaturalSize(image);
  if (!(Number.isFinite(natural.w) && natural.w > 0)
    || !(Number.isFinite(natural.h) && natural.h > 0)) return null;
  const cssPixelsPerImagePixel = 96 / authored.dpi * (ptToPx / PT_TO_PX);
  const tileW = natural.w * authored.sx * cssPixelsPerImagePixel;
  const tileH = natural.h * authored.sy * cssPixelsPerImagePixel;
  if (!(tileW > 0) || !(tileH > 0)) return null;
  return {
    alignment: authored.alignment,
    tileW,
    tileH,
    offsetX: authored.tx / EMU_PER_PT * ptToPx,
    offsetY: authored.ty / EMU_PER_PT * ptToPx,
    flipX: authored.flipX,
    flipY: authored.flipY,
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
  if (!fillCanProduceVisiblePixels(fill) || !imageFillModeIsPaintable(fill)) return 0;
  const image = (lookup ?? activeLookup)?.(fill);
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
  if (!fillCanProduceVisiblePixels(fill) || !imageFillModeIsPaintable(fill)) return 0;
  const image = (lookup ?? activeLookup)?.(fill);
  if (!image) return 0;
  if (!fill.tile) return 1;
  const geometry = imageTileGeometry(fill, image, w, h, ptToPx);
  return geometry ? Math.min(geometry.repetitions, MAX_CHART_IMAGE_FILL_TILES) : 0;
}

/** Per-source facts retained from every reachable chart picture-fill
 * occurrence before identical decoded sources are deduplicated. Frame-relative
 * factors let DOCX/PPTX/XLSX hosts apply their own display scale and DPR. */
export interface ChartImageFillUsage {
  /** Stable representative carrying the source path, MIME, SVG twin and effects. */
  readonly fill: ImageFill;
  /** A statically valid `<a:tile>` occurrence needs native decoded dimensions. */
  readonly preserveNaturalSize: boolean;
  /** Any authored `<a:srcRect>` occurrence forces the raster SVG fallback. */
  readonly hasSourceCrop: boolean;
  /** Largest stretched destination/source ratio on each axis. */
  readonly targetWidthFactor: number;
  readonly targetHeightFactor: number;
  /** Largest post-crop full metafile-frame ratio on each axis. */
  readonly metafileWidthFactor: number;
  readonly metafileHeightFactor: number;
}

export interface ChartImageFillUsageSize {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly targetWidthPx?: number;
  readonly targetHeightPx?: number;
}

type ChartImageFillOccurrence = Omit<ChartImageFillUsage, 'fill'>;

function isPositiveFiniteFactor(value: number): boolean {
  return Number.isFinite(value) && value > 0;
}

/** Apply one validated usage to a host chart frame without allowing finite
 * inputs whose products overflow (or underflow to zero) to become an absent
 * decode target. Hosts use this before aggregate source gating. */
export function chartImageFillUsageSize(
  usage: ChartImageFillUsage,
  frame: Readonly<{
    widthPt: number;
    heightPt: number;
    targetWidthPx?: number;
    targetHeightPx?: number;
  }>,
): ChartImageFillUsageSize | null {
  if (!isPositiveFiniteFactor(frame.widthPt)
    || !isPositiveFiniteFactor(frame.heightPt)
    || !isPositiveFiniteFactor(usage.metafileWidthFactor)
    || !isPositiveFiniteFactor(usage.metafileHeightFactor)
    || !Number.isFinite(usage.targetWidthFactor) || usage.targetWidthFactor < 0
    || !Number.isFinite(usage.targetHeightFactor) || usage.targetHeightFactor < 0) return null;
  const widthPt = frame.widthPt * usage.metafileWidthFactor;
  const heightPt = frame.heightPt * usage.metafileHeightFactor;
  if (!isPositiveFiniteFactor(widthPt) || !isPositiveFiniteFactor(heightPt)) return null;
  const targetFrameWidthPx = frame.targetWidthPx;
  const targetFrameHeightPx = frame.targetHeightPx;
  const hasTargetWidth = targetFrameWidthPx != null;
  const hasTargetHeight = targetFrameHeightPx != null;
  if (hasTargetWidth !== hasTargetHeight) return null;
  if (targetFrameWidthPx == null || targetFrameHeightPx == null) return { widthPt, heightPt };
  if (!isPositiveFiniteFactor(targetFrameWidthPx)
    || !isPositiveFiniteFactor(targetFrameHeightPx)) return null;
  const rawTargetWidthPx = targetFrameWidthPx * usage.targetWidthFactor;
  const rawTargetHeightPx = targetFrameHeightPx * usage.targetHeightFactor;
  if (!Number.isFinite(rawTargetWidthPx) || rawTargetWidthPx < 0
    || (usage.targetWidthFactor > 0 && rawTargetWidthPx === 0)
    || !Number.isFinite(rawTargetHeightPx) || rawTargetHeightPx < 0
    || (usage.targetHeightFactor > 0 && rawTargetHeightPx === 0)) return null;
  const targetWidthPx = Math.ceil(rawTargetWidthPx);
  const targetHeightPx = Math.ceil(rawTargetHeightPx);
  if (!Number.isFinite(targetWidthPx) || !Number.isFinite(targetHeightPx)) return null;
  return { widthPt, heightPt, targetWidthPx, targetHeightPx };
}

function chartImageFillOccurrence(fill: ImageFill): ChartImageFillOccurrence | null {
  if (!fillCanProduceVisiblePixels(fill) || !imageFillModeIsPaintable(fill)) return null;
  const logicalWidth = fill.srcRect ? 1 - fill.srcRect.l - fill.srcRect.r : 1;
  const logicalHeight = fill.srcRect ? 1 - fill.srcRect.t - fill.srcRect.b : 1;
  if (!(Number.isFinite(logicalWidth) && logicalWidth > 0)
    || !(Number.isFinite(logicalHeight) && logicalHeight > 0)) return null;
  const hasSourceCrop = fill.srcRect != null;
  if (fill.tile) {
    if (!chartImageTileStaticMetrics(fill)) return null;
    const metafileWidthFactor = 1 / logicalWidth;
    const metafileHeightFactor = 1 / logicalHeight;
    if (!isPositiveFiniteFactor(metafileWidthFactor)
      || !isPositiveFiniteFactor(metafileHeightFactor)) return null;
    return {
      preserveNaturalSize: true,
      hasSourceCrop,
      targetWidthFactor: 0,
      targetHeightFactor: 0,
      metafileWidthFactor,
      metafileHeightFactor,
    };
  }
  const rect = fill.fillRect;
  const left = rect?.l ?? 0;
  const top = rect?.t ?? 0;
  const right = rect?.r ?? 0;
  const bottom = rect?.b ?? 0;
  if (![left, top, right, bottom].every(Number.isFinite)) return null;
  const destinationWidth = 1 - left - right;
  const destinationHeight = 1 - top - bottom;
  if (!isPositiveFiniteFactor(destinationWidth)
    || !isPositiveFiniteFactor(destinationHeight)) return null;
  const widthFactor = destinationWidth / logicalWidth;
  const heightFactor = destinationHeight / logicalHeight;
  if (!isPositiveFiniteFactor(widthFactor)
    || !isPositiveFiniteFactor(heightFactor)) return null;
  return {
    preserveNaturalSize: false,
    hasSourceCrop,
    targetWidthFactor: widthFactor,
    targetHeightFactor: heightFactor,
    metafileWidthFactor: widthFactor,
    metafileHeightFactor: heightFactor,
  };
}

function mergeChartImageFillUsages(
  left: ChartImageFillUsage,
  right: ChartImageFillUsage,
): ChartImageFillUsage {
  return {
    fill: left.fill,
    preserveNaturalSize: left.preserveNaturalSize || right.preserveNaturalSize,
    hasSourceCrop: left.hasSourceCrop || right.hasSourceCrop,
    targetWidthFactor: Math.max(left.targetWidthFactor, right.targetWidthFactor),
    targetHeightFactor: Math.max(left.targetHeightFactor, right.targetHeightFactor),
    metafileWidthFactor: Math.max(left.metafileWidthFactor, right.metafileWidthFactor),
    metafileHeightFactor: Math.max(left.metafileHeightFactor, right.metafileHeightFactor),
  };
}

/** Unique picture fills reachable by marker consumers. Family suppression,
 * point validity and direct precedence match the painters so hosts never fetch
 * images for markers that cannot be drawn. */
interface ChartMarkerImageFillResult {
  usages: ChartImageFillUsage[];
  sourceLimitExceeded: boolean;
  usageRejected: boolean;
}

function collectChartMarkerImageFillResult(
  chart: ChartModel,
  acceptUsage?: (usage: ChartImageFillUsage) => boolean,
): ChartMarkerImageFillResult {
  const sourceCount = sourceChartStructureCount(chart);
  if (sourceCount > MAX_CANVAS_CHART_POINTS) {
    return { usages: [], sourceLimitExceeded: false, usageRejected: false };
  }
  // Match the renderer's first semantic projection. Hidden source points must
  // not consume the decoded-image source ceiling or suppress a visible peer.
  chart = applyPlotVisibleOnly(chart);
  // The host preflights image resources before the synchronous renderer runs.
  // Resolve the same numeric-plus-linked role cascade here so the image that is
  // warmed is exactly the one the renderer can select.
  chart = withEffectiveChartStyleRoles(chart);
  const usages = new Map<string, ChartImageFillUsage>();
  let sourceLimitExceeded = false;
  let usageRejected = false;
  const add = (fill: unknown) => {
    if (usageRejected) return;
    if (!fill || typeof fill !== 'object' || (fill as ImageFill).fillType !== 'image') return;
    const image = fill as ImageFill;
    const occurrence = chartImageFillOccurrence(image);
    if (!occurrence) return;
    const key = chartImageFillKey(image);
    const usage: ChartImageFillUsage = { fill: image, ...occurrence };
    if (acceptUsage && !acceptUsage(usage)) {
      usageRejected = true;
      return;
    }
    const prior = usages.get(key);
    if (prior) {
      usages.set(key, mergeChartImageFillUsages(prior, usage));
      return;
    }
    if (usages.size >= MAX_CHART_MARKER_IMAGE_SOURCES) {
      sourceLimitExceeded = true;
      return;
    }
    usages.set(key, usage);
  };
  const styleImageDecision = (
    style: ChartExElementStyle | null | undefined,
    index: number,
  ): ImageFill | null | undefined => {
    const fill = chartStyleFillDecision(style, index);
    return fill?.fillType === 'image' ? fill : fill == null ? fill : null;
  };
  const directStyleImageDecision = (
    style: ChartExElementStyle | null | undefined,
    rawLinked: ChartExElementStyle | null | undefined,
    index: number,
  ): ImageFill | null | undefined => {
    const fill = chartStyleDirectFillDecision(style, rawLinked, index);
    return fill?.fillType === 'image' ? fill : fill == null ? fill : null;
  };
  const dataPointImageDecision = (
    local: ChartExElementStyle | null | undefined,
    legacyColor: string | null | undefined,
    linked: ChartExElementStyle | null | undefined,
    rawLinked: ChartExElementStyle | null | undefined,
    index: number,
  ): ImageFill | null | undefined => {
    const fill = local?.fillHidden
      ? chartStyleDirectNoFillDecision(rawLinked)
      : chartStyleDirectFillDecision(local, rawLinked, index);
    if (fill !== undefined) return fill?.fillType === 'image' ? fill : null;
    if (legacyColor) return null;
    const fallback = chartStyleFillDecision(linked, index);
    return fallback?.fillType === 'image' ? fallback : fallback == null ? fallback : null;
  };
  const waterfallBodyImageDecision = (
    point: ChartDataPointOverride | undefined,
    series: ChartSeries | undefined,
    semanticIndex: number,
  ): ImageFill | null | undefined => {
    const pointAuthors = point?.fillHidden === true || point?.color != null
      || point?.chartexStyle?.fillPaintAuthored === true
      || point?.chartexStyle?.fillHidden != null
      || point?.chartexStyle?.fillColors?.some(color => color != null) === true
      || point?.chartexStyle?.fillPaints?.some(paint => paint != null) === true;
    if (pointAuthors) {
      const pointStyle = point?.fillHidden === true
        ? { ...point.chartexStyle, fillHidden: true, fillPaintAuthored: true }
        : point?.chartexStyle;
      return dataPointImageDecision(
        pointStyle,
        point?.color,
        chart.chartexDataPointStyle,
        rawLinkedChartStyleRole(chart, 'dataPoint')
          ?? (chart.classicChartStyleRoles == null ? chart.chartexDataPointStyle : undefined),
        semanticIndex,
      );
    }
    if (series?.chartexStyle?.fillPaintAuthored === true) {
      return dataPointImageDecision(
        series.chartexStyle,
        series.color,
        chart.chartexDataPointStyle,
        rawLinkedChartStyleRole(chart, 'dataPoint')
          ?? (chart.classicChartStyleRoles == null ? chart.chartexDataPointStyle : undefined),
        semanticIndex,
      );
    }
    return dataPointImageDecision(
      series?.chartexStyle, series?.color, chart.chartexDataPointStyle,
      rawLinkedChartStyleRole(chart, 'dataPoint'), semanticIndex,
    );
  };
  const frameImageDecision = (
    directFill: ChartModel['chartFill'] | ChartModel['plotAreaFill'],
    directColor: string | null | undefined,
    hidden: boolean | null | undefined,
    paintAuthored: boolean | null | undefined,
    linked: ChartExElementStyle | null | undefined,
    rawLinked: ChartExElementStyle | null | undefined,
    directStyle: ChartExElementStyle | null | undefined,
  ): ImageFill | null | undefined => {
    if (hidden === true) {
      const noFill = chartStyleDirectNoFillDecision(rawLinked);
      if (noFill !== undefined) return noFill;
    }
    if (directFill?.fillType === 'image') return directFill;
    if (directFill != null || directColor != null
      || paintAuthored === true && hidden !== true) return null;
    if (linked?.fillNoStyle === true) return undefined;
    const fill = chartStyleFillCascade(linked, rawLinked, 0, directStyle);
    return fill?.fillType === 'image' ? fill : fill == null ? fill : null;
  };
  const chartAreaImage = frameImageDecision(
    chart.chartFill,
    // `chartStyleRoleChartArea` treats the structured/provenance fields as
    // authoritative; retain that exact ownership rule for prefetch too.
    undefined,
    chart.chartFillHidden,
    chart.chartFillPaintAuthored,
    chart.chartStyleRoles?.chartArea,
    rawLinkedChartStyleRole(chart, 'chartArea'),
    chart.chartAreaStyle,
  );
  if (chartAreaImage) add(chartAreaImage);
  const plotAreaImage = frameImageDecision(
    chart.plotAreaFill,
    chart.plotAreaBg,
    chart.plotAreaFillHidden,
    chart.plotAreaFillPaintAuthored,
    chart.threeD ? chart.chartStyleRoles?.plotArea3D : chart.chartStyleRoles?.plotArea,
    rawLinkedChartStyleRole(chart, chart.threeD ? 'plotArea3D' : 'plotArea'),
    chart.plotAreaStyle,
  );
  if (plotAreaImage) add(plotAreaImage);
  // Legend layout and its frame remain visible even when every legend entry
  // is deleted (or no series exists), so prefetch the frame by the same
  // `showLegend` gate used by paint/effect preflight rather than entry count.
  if (chart.showLegend) add(effectiveLegendFrameFill(chart));
  const addLabelBoxFill = (
    direct: Parameters<typeof effectiveChartLabelBoxFill>[0],
    linked: Parameters<typeof effectiveChartLabelBoxFill>[1],
    rawLinked: Parameters<typeof effectiveChartLabelBoxFill>[2],
    linkedIndex = 0,
  ): void => {
    add(effectiveChartLabelBoxFill(
      direct, linked, rawLinked, true, 0, linkedIndex,
    )?.fillPaint);
  };
  if (chart.title) {
    addLabelBoxFill(
      chart.titleStyle ? { style: chart.titleStyle } : undefined,
      chart.chartStyleRoles?.title,
      rawLinkedChartStyleRole(chart, 'title'),
    );
  }
  if (chart.catAxisTitle) {
    addLabelBoxFill(
      chart.catAxisTitleStyle ? { style: chart.catAxisTitleStyle } : undefined,
      chart.chartStyleRoles?.axisTitle,
      rawLinkedChartStyleRole(chart, 'axisTitle'),
    );
  }
  if (chart.valAxisTitle) {
    addLabelBoxFill(
      chart.valAxisTitleStyle ? { style: chart.valAxisTitleStyle } : undefined,
      chart.chartStyleRoles?.axisTitle,
      rawLinkedChartStyleRole(chart, 'axisTitle'),
    );
  }
  for (const axis of [chart.secondaryValAxis, chart.secondaryCatAxis]) {
    if (axis?.title) {
      addLabelBoxFill(
        axis.titleStyle ? { style: axis.titleStyle } : undefined,
        chart.chartStyleRoles?.axisTitle,
        rawLinkedChartStyleRole(chart, 'axisTitle'),
      );
    }
    if (axis?.displayUnits?.label) {
      addLabelBoxFill(
        axis.displayUnits.label.boxStyle,
        chart.chartStyleRoles?.axisTitle,
        rawLinkedChartStyleRole(chart, 'axisTitle'),
      );
    }
  }
  if (chart.threeD?.seriesAxis?.title) {
    addLabelBoxFill(
      chart.threeD.seriesAxis.titleStyle
        ? { style: chart.threeD.seriesAxis.titleStyle } : undefined,
      chart.chartStyleRoles?.axisTitle,
      rawLinkedChartStyleRole(chart, 'axisTitle'),
    );
  }
  for (const units of [chart.valAxisDisplayUnits, chart.catAxisDisplayUnits]) {
    if (units?.label) {
      addLabelBoxFill(
        units.label.boxStyle,
        chart.chartStyleRoles?.axisTitle,
        rawLinkedChartStyleRole(chart, 'axisTitle'),
      );
    }
  }
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
  const upDownBarImage = (
    direct: ChartStockBarPaint,
    role: 'upBar' | 'downBar',
  ): ImageFill | null | undefined => {
    const fill = chartStockBarFillDecision(chart, direct, role);
    return fill?.fillType === 'image' ? fill : fill == null ? fill : null;
  };
  const collectUpDownDirections = (
    start: ChartModel['series'][number] | undefined,
    end: ChartModel['series'][number] | undefined,
    direct: NonNullable<ChartModel['stockUpDownBarStyle']>,
  ) => {
    if (!start || !end) return;
    let hasUp = false;
    let hasDown = false;
    const count = Math.max(start.values.length, end.values.length);
    for (let index = 0; index < count && !(hasUp && hasDown); index++) {
      const startValue = start.values[index];
      const endValue = end.values[index];
      if (startValue == null || endValue == null
        || !Number.isFinite(startValue) || !Number.isFinite(endValue)
        || startValue === endValue) continue;
      if (endValue > startValue) hasUp = true;
      else hasDown = true;
    }
    if (hasUp) add(upDownBarImage(direct.up, 'upBar'));
    if (hasDown) add(upDownBarImage(direct.down, 'downBar'));
  };
  for (const decoration of chart.lineGroupDecorations ?? []) {
    if (!decoration.upDownBars) continue;
    let members = chart.series.filter(series => series.lineGroupIndex === decoration.groupIndex);
    if (members.length === 0 && decoration.groupIndex === 0
      && ['line', 'stackedLine', 'stackedLinePct'].includes(chart.chartType)) {
      members = chart.series.filter(series => series.seriesType == null || series.seriesType === 'line');
    }
    collectUpDownDirections(
      members[0], members.at(-1), decoration.upDownBars,
    );
  }
  if (chart.stockUpDownBars) {
    const stockGroup = chart.plotGroups?.find(group => group.kind === 'stock');
    const stockSeries = stockGroup
      ? chart.series.slice(stockGroup.seriesStart, stockGroup.seriesStart + stockGroup.seriesCount)
      : chart.series;
    collectUpDownDirections(
      stockSeries[0],
      stockSeries.at(-1),
      chart.stockUpDownBarStyle ?? { gapWidthPercent: 150, up: {}, down: {} },
    );
  }
  const scatterHasNumericX = chart.series.some((series, seriesIndex) => {
    const group = plotGroupBySeries[seriesIndex];
    const family = group?.kind === 'bubble' || group?.kind === 'scatter'
      ? 'scatter'
      : series.seriesType ?? (chart.chartType === 'bubble' ? 'scatter' : chart.chartType);
    return family === 'scatter' && (series.categories ?? chart.categories).some(category =>
      Number.isFinite(Number.parseFloat(category))
    );
  });
  const deletedLegendEntries = deletedLegendEntryIndices(chart);
  const legendRanges = legendEntryRanges(chart, true);
  const chartHasCategories = chart.categories.length > 0
    || (chart.series[0]?.categories?.length ?? 0) > 0
    || chart.series.some(series => series.values.length > 0);
  for (let seriesIndex = 0; seriesIndex < chart.series.length; seriesIndex++) {
    const series = chart.series[seriesIndex];
    const sourceSeriesIndex = series.chartexFormatIdx ?? seriesIndex;
    const variesByPoint = chartSeriesVariesByPoint(chart, seriesIndex);
    const linkedMarkerStyle = chartDataPointStyleRole(chart, 'dataPointMarker', seriesIndex);
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
    const areaFamily = family === 'area' || family === 'stackedArea'
      || family === 'stackedAreaPct';
    const seriesVisible = areaFamily
      ? (series.showMarker === true || seriesHasMarkerDetail(series))
        && series.markerSymbol !== 'none'
      : family === 'stock'
        ? series.markerSymbol != null && series.markerSymbol !== 'none'
        : series.showMarker !== false && series.markerSymbol !== 'none';
    const seriesLegendVisible = legendSeriesHasVisibleEntry(
      legendRanges, deletedLegendEntries, seriesIndex,
    );
    const pointCount = Math.max(
      series.values.length, series.categories?.length ?? 0, chart.categories.length,
    );
    const dataLabelOverrides = new Map(
      (series.dataLabelOverrides ?? []).map(label => [label.idx, label]),
    );
    for (let index = 0; index < pointCount; index++) {
      const label = dataLabelOverrides.get(index);
      if (dataLabelIsDeleted(series.seriesDataLabels, label)) continue;
      const hasContent = Boolean(
        label?.text
        || (label?.showVal ?? series.seriesDataLabels?.showVal ?? chart.showDataLabels)
        || (label?.showCatName ?? series.seriesDataLabels?.showCatName)
        || (label?.showSerName ?? series.seriesDataLabels?.showSerName)
        || (label?.showPercent ?? series.seriesDataLabels?.showPercent)
        || (label?.showBubbleSize ?? series.seriesDataLabels?.showBubbleSize)
        || (label?.showLegendKey ?? series.seriesDataLabels?.showLegendKey)
      );
      if (!hasContent || !classicDataLabelPointIsPainted(
        chart, series, family, index, scatterHasNumericX, seriesIndex,
      )) continue;
      const direct = mergeChartLabelBoxes(label?.labelBox, series.seriesDataLabels?.labelBox);
      const usesCalloutRole = chartLabelBoxHasVisiblePaint(direct);
      const linked = usesCalloutRole
        ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
        : chart.chartStyleRoles?.dataLabel;
      const rawLinked = usesCalloutRole
        ? rawLinkedChartStyleRole(chart, 'dataLabelCallout')
          ?? rawLinkedChartStyleRole(chart, 'dataLabel')
        : rawLinkedChartStyleRole(chart, 'dataLabel');
      addLabelBoxFill(direct, linked, rawLinked, label ? index : sourceSeriesIndex);
    }
    for (const trendline of series.trendLines ?? []) {
      const hasLabelContent = trendline.dispEq === true || trendline.dispRSqr === true
        || Boolean(trendline.labelText)
        || trendline.labelRichRuns?.some(run => run.text.length > 0) === true;
      if (hasLabelContent) {
        addLabelBoxFill(
          trendline.labelBox,
          chart.chartStyleRoles?.trendlineLabel,
          rawLinkedChartStyleRole(chart, 'trendlineLabel'),
          sourceSeriesIndex,
        );
      }
    }
    const labelKeyVisible = dataLabelLegendKeyCount(
      chart, series, family, pointCount, scatterHasNumericX, markerContext, seriesIndex,
    ) > 0;
    const labelOverrides = new Map(
      (series.dataLabelOverrides ?? []).map(label => [label.idx, label]),
    );
    const pointLabelKeyVisible = (pointIndex: number): boolean => {
      const label = labelOverrides.get(pointIndex);
      return !dataLabelIsDeleted(series.seriesDataLabels, label)
        && (label?.showLegendKey ?? series.seriesDataLabels?.showLegendKey ?? false) === true
        && classicDataLabelPointIsPainted(
          chart, series, family, pointIndex, scatterHasNumericX, seriesIndex,
        );
    };
    const pointDrivenLegend = legendRanges[seriesIndex]?.pointDriven === true;
    const seriesKeyVisible = !pointDrivenLegend && seriesLegendMarkerIsVisible(
      effectiveChartType, effectiveScatterStyle, series, effectiveRadarStyle,
    ) && ((chart.showLegend && seriesLegendVisible)
      || (chart.dataTable?.showKeys === true && chartDataTableFamilyIsPainted(chart.chartType)
        && chartHasCategories)
      || labelKeyVisible);
    const seriesKeySymbol = series.markerSymbol ?? (family === 'stock' ? 'none' : 'circle');
    const classicThreeDGroup = group?.kind === 'bar3D' || group?.kind === 'pie3D'
      || group?.kind === 'area3D' || group?.kind === 'line3D'
      || group?.kind === 'surface3D';
    // Parsed classic bar series retain the canonical `seriesType: "bar"`;
    // group.kind owns the 2-D/3-D distinction. Preflight only the flat body
    // painters that can actually consume these pictures.
    const classicFilledFamily = !classicThreeDGroup
      && (group?.kind === 'bar' || family === 'bar'
      || family === 'clusteredBar' || family === 'clusteredBarH'
      || family === 'stackedBar' || family === 'stackedBarH'
      || family === 'stackedBarPct' || family === 'stackedBarHPct'
      || family === 'pie' || family === 'doughnut'
      || family === 'ofPie' || areaFamily
      || (family === 'radar' && effectiveRadarStyle === 'filled'));
    if (classicFilledFamily) {
      const pointOverrides = new Map(
        (series.dataPointOverrides ?? []).map(point => [point.idx, point]),
      );
      const bodyHasGeometry = areaFamily
        ? series.values.length > 0
        : family === 'radar'
          ? pointCount > 2 && Array.from({ length: pointCount }, (_, index) =>
              series.values[index] != null && Number.isFinite(series.values[index])
            ).every(Boolean)
          : family === 'pie' || family === 'pie3D' || family === 'doughnut' || family === 'ofPie'
            ? series.values.some(value => value != null && Number.isFinite(value) && value !== 0)
            : series.values.some(value => value != null && Number.isFinite(value) && value !== 0);
      const keyCanPaint = (chart.showLegend && seriesLegendVisible)
        || (chart.dataTable?.showKeys === true && chartDataTableFamilyIsPainted(chart.chartType)
          && chartHasCategories)
        || labelKeyVisible;
      const pointDrivenKeys = variesByPoint
        || family === 'pie' || family === 'doughnut' || family === 'ofPie';
      if (bodyHasGeometry || keyCanPaint) {
        if (areaFamily || (family === 'radar' && !variesByPoint)) {
          const decision = classicDataPointFillDecision(
            chart, series, undefined, sourceSeriesIndex,
          );
          if (decision?.fillType === 'image') add(decision);
        } else if (family === 'radar') {
          const point = pointOverrides.get(0);
          const decision = classicDataPointFillDecision(chart, series, point, 0, 0);
          if (decision?.fillType === 'image') add(decision);
        } else {
          // A series-driven legend/table/label key is painted once without a
          // dPt override. Collect it independently from body reachability.
          if (keyCanPaint && !pointDrivenKeys) {
            const keyDecision = classicDataPointFillDecision(
              chart, series, undefined, sourceSeriesIndex,
            );
            if (keyDecision?.fillType === 'image') add(keyDecision);
          }
          for (let pointIndex = 0; pointIndex < pointCount; pointIndex++) {
            const value = series.values[pointIndex];
            const bodyPointCanPaint = value != null
              && Number.isFinite(value) && value !== 0;
            const pointKeyCanPaint = pointDrivenKeys && (
              (chart.showLegend && legendEntryIsVisible(
                legendRanges, deletedLegendEntries, seriesIndex, pointIndex,
              ))
              || (chart.dataTable?.showKeys === true
                && chartDataTableFamilyIsPainted(chart.chartType)
                && chartHasCategories && pointIndex === 0)
              || pointLabelKeyVisible(pointIndex)
            );
            if (!bodyPointCanPaint && !pointKeyCanPaint) continue;
            const point = pointOverrides.get(pointIndex);
            const styleIndex = variesByPoint ? pointIndex : sourceSeriesIndex;
            const decision = classicDataPointFillDecision(
              chart, series, point, styleIndex, pointIndex,
            );
            if (decision?.fillType === 'image') add(decision);
          }
        }
      }
    }
    if (!markerFamily || markersSuppressedByChartStyle(
      family, effectiveChartType, effectiveScatterStyle, effectiveRadarStyle,
    )) continue;
    if (seriesKeyVisible && markerSymbolConsumesFill(seriesKeySymbol)) {
      if (isBubble) {
        const rawDataPointRole = rawLinkedChartStyleRole(chart, 'dataPoint');
        const seriesShape = directStyleImageDecision(
          series.chartexStyle, rawDataPointRole, sourceSeriesIndex,
        );
        if (seriesShape) add(seriesShape);
        if (seriesShape === undefined && series.color == null) {
          const linkedShape = styleImageDecision(
            chartDataPointStyleRole(chart, 'dataPoint', seriesIndex),
            sourceSeriesIndex,
          );
          if (linkedShape) add(linkedShape);
        }
      } else add(series.markerFillPaint);
      if (!isBubble
        && series.markerFillPaint === undefined && series.markerFill == null
        && !(series.markerFillPaintAuthored === true
          && series.markerStyle?.fillHidden !== true)) {
        if (!variesByPoint) {
          const linkedFill = chartStyleFillCascade(
            linkedMarkerStyle,
            rawLinkedChartStyleRole(chart, 'dataPointMarker'),
            sourceSeriesIndex,
            series.markerStyle,
          );
          if (linkedFill?.fillType === 'image') add(linkedFill);
        }
      }
    }
    if (!seriesVisible && !hasVisiblePointMarkerOverride(series)) continue;
    const overrides = new Map((series.dataPointOverrides ?? []).map(point => [point.idx, point]));
    for (let index = 0; index < pointCount; index++) {
      const plotPointVisible = classicMarkerPointIsPainted(
        chart, series, family, index, scatterHasNumericX, markerContext,
      ) && (!isBubble
        || ((group?.bubbleScale ?? chart.bubbleScale ?? 100) > 0
          && visibleBubbleSize(
            { showNegativeBubbles: group?.showNegativeBubbles ?? chart.showNegativeBubbles },
            series.bubbleSizes?.[index],
          ) != null));
      const legendKeyVisible = pointDrivenLegend && chart.showLegend
        && legendEntryIsVisible(legendRanges, deletedLegendEntries, seriesIndex, index);
      const tableKeyVisible = pointDrivenLegend
        && chart.dataTable?.showKeys === true
        && chartDataTableFamilyIsPainted(chart.chartType)
        && chartHasCategories && index === 0;
      const labelPointKeyVisible = pointDrivenLegend && pointLabelKeyVisible(index);
      if (!plotPointVisible && !legendKeyVisible && !tableKeyVisible && !labelPointKeyVisible) {
        continue;
      }
      const point = overrides.get(index);
      const symbol = effectiveMarkerSymbol(series, point, 'circle', seriesVisible);
      if (!markerSymbolConsumesFill(symbol)) continue;
      if (isBubble) {
        // MS-OE376 §2.1.1504(b): a visible negative bubble uses the
        // application's inverted fill, not its positive direct/linked image.
        if ((series.bubbleSizes?.[index] ?? 0) < 0) {
          continue;
        }
        const linkedPointRole = chartDataPointStyleRole(chart, 'dataPoint', seriesIndex);
        const linkedPointIndex = chartSeriesVariesByPoint(chart, seriesIndex)
          ? index : sourceSeriesIndex;
        const rawDataPointRole = rawLinkedChartStyleRole(chart, 'dataPoint');
        const pointDecision = chartStyleDirectFillDecision(
          point?.chartexStyle, rawDataPointRole, index,
        );
        const pointShape = pointDecision?.fillType === 'image'
          ? pointDecision : pointDecision == null ? pointDecision : null;
        if (pointShape) add(pointShape);
        const pointNoFill = point?.fillHidden === true
          ? chartStyleDirectNoFillDecision(rawDataPointRole)
          : undefined;
        if (pointShape !== undefined || pointNoFill !== undefined
          || point?.color != null) continue;
        if (series.dataPointColors?.[index] != null) continue;
        const seriesShape = directStyleImageDecision(
          series.chartexStyle, rawDataPointRole, sourceSeriesIndex,
        );
        if (seriesShape) add(seriesShape);
        if (seriesShape !== undefined || series.color != null) continue;
        const linkedShape = styleImageDecision(linkedPointRole, linkedPointIndex);
        if (linkedShape) add(linkedShape);
        continue;
      }
      const paint = markerFillPaintFor(series, point, index);
      add(paint);
      if (paint === undefined && point?.markerFill == null && point?.color == null
        && series.dataPointColors?.[index] == null && series.markerFill == null
        && !(point?.markerFillPaintAuthored === true
          && point.markerStyle?.fillHidden !== true)
        && !(series.markerFillPaintAuthored === true
          && series.markerStyle?.fillHidden !== true)) {
        const directStyle = point?.markerStyle?.shapePropertiesPresent === true
          ? point.markerStyle
          : series.markerStyle;
        const linkedFill = chartStyleFillCascade(
          linkedMarkerStyle,
          rawLinkedChartStyleRole(chart, 'dataPointMarker'),
          variesByPoint ? index : sourceSeriesIndex,
          directStyle,
        );
        if (linkedFill?.fillType === 'image') add(linkedFill);
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
    const linkedStyle = chart.chartexDataPointMarkerStyle
      ?? chart.chartexDataPointStyle ?? undefined;
    const rawLinkedStyle = chart.chartexDataPointMarkerStyle != null
      ? rawLinkedChartStyleRole(chart, 'dataPointMarker')
        ?? (chart.classicChartStyleRoles == null ? linkedStyle : undefined)
      : rawLinkedChartStyleRole(chart, 'dataPoint')
        ?? (chart.classicChartStyleRoles == null ? linkedStyle : undefined);
    const local = directStyleImageDecision(localStyle, rawLinkedStyle, styleIndex);
    if (local) add(local);
    if (local !== undefined || series.color != null) continue;
    const linked = styleImageDecision(linkedStyle, styleIndex);
    if (linked) add(linked);
  }
  if (chart.chartType === 'waterfall') {
    const series = chart.series[0];
    const overrides = new Map((series?.dataPointOverrides ?? []).map(point => [point.idx, point]));
    const plan = planWaterfallPaintSites(
      series?.values ?? [], chart.categories.length, chart.subtotalIndices,
    );
    if (!plan.cumulativeOverflow && plan.rawMax > plan.rawMin) {
      for (let index = 0; index < plan.bars.length; index++) {
        const bar = plan.bars[index]!;
        if (!bar.paintSlot) continue;
        add(waterfallBodyImageDecision(overrides.get(index), series, bar.semanticIndex));
      }
    }
  } else if (chart.chartType === 'funnel') {
    const series = chart.series[0];
    if ((series?.values ?? []).some(value => value != null && value > 0)) {
      add(dataPointImageDecision(
        series?.chartexStyle, series?.color, chart.chartexDataPointStyle,
        rawLinkedChartStyleRole(chart, 'dataPoint'), 0,
      ));
    }
  }
  for (let seriesIndex = 0; seriesIndex < (chart.chartexBox?.series.length ?? 0); seriesIndex++) {
    const series = chart.chartexBox!.series[seriesIndex]!;
    if (!series.valuesByCategory.some(values =>
      computeBoxWhiskerStats(values, series.quartileMethod) != null)) continue;
    add(dataPointImageDecision(
      series.chartexStyle, series.color, chart.chartexDataPointStyle,
      rawLinkedChartStyleRole(chart, 'dataPoint'),
      series.chartexFormatIdx ?? seriesIndex,
    ));
  }
  const hierarchySeries = chart.series[0];
  visitChartExHierarchyBodySites(chart, ({ node, paintsBody }) => {
    if (!paintsBody) return;
    add(dataPointImageDecision(
      hierarchySeries?.chartexStyle, hierarchySeries?.color,
      chart.chartexDataPointStyle, rawLinkedChartStyleRole(chart, 'dataPoint'),
      node.branchIndex,
    ));
  });
  visitChartExHierarchyLabelSites(chart, ({ label, linkedStyleIndex }) => {
    const direct = label.labelBox;
    const usesCalloutRole = chartLabelBoxHasVisiblePaint(direct);
    const linked = usesCalloutRole
      ? chart.chartStyleRoles?.dataLabelCallout ?? chart.chartStyleRoles?.dataLabel
      : chart.chartStyleRoles?.dataLabel;
    const rawLinked = usesCalloutRole
      ? rawLinkedChartStyleRole(chart, 'dataLabelCallout')
        ?? rawLinkedChartStyleRole(chart, 'dataLabel')
      : rawLinkedChartStyleRole(chart, 'dataLabel');
    const effective = effectiveChartLabelBoxFill(
      direct, linked, rawLinked, linked != null, 0, linkedStyleIndex,
    );
    if (effective?.fillPaint?.fillType === 'image') add(effective.fillPaint);
  });
  return {
    usages: sourceLimitExceeded || usageRejected ? [] : [...usages.values()],
    sourceLimitExceeded,
    usageRejected,
  };
}

export function collectChartImageFillUsages(chart: ChartModel): ChartImageFillUsage[] {
  return withChartStyleIndexCache(() => withSparseStyleIndexCache(() =>
    collectChartMarkerImageFillResult(chart).usages
  ));
}

export function collectChartMarkerImageFills(chart: ChartModel): ImageFill[] {
  return collectChartImageFillUsages(chart).map(usage => usage.fill);
}

/** Collect chart picture fills for one host render pass. Hosts retain the decoded
 * sources until that page/slide/viewport paint completes, so the count ceiling
 * applies to the aggregate rather than independently to each chart. */
export function collectChartImageFillUsagesForCharts(
  charts: readonly ChartModel[],
  acceptUsage?: (usage: ChartImageFillUsage, chartIndex: number) => boolean,
): ChartImageFillUsage[] {
  return withChartStyleIndexCache(() => withSparseStyleIndexCache(() => {
    const usages = new Map<string, ChartImageFillUsage>();
    for (let chartIndex = 0; chartIndex < charts.length; chartIndex++) {
      const chart = charts[chartIndex]!;
      const result = collectChartMarkerImageFillResult(
        chart,
        acceptUsage ? usage => acceptUsage(usage, chartIndex) : undefined,
      );
      // Host frame validation rejects the whole owning chart before either its
      // per-chart source ceiling or the aggregate ceiling can suppress peers.
      if (result.usageRejected) continue;
      if (result.sourceLimitExceeded) return [];
      for (const usage of result.usages) {
        const key = chartImageFillKey(usage.fill);
        const prior = usages.get(key);
        if (prior) {
          usages.set(key, mergeChartImageFillUsages(prior, usage));
          continue;
        }
        if (usages.size >= MAX_CHART_MARKER_IMAGE_SOURCES) return [];
        usages.set(key, usage);
      }
    }
    return [...usages.values()];
  }));
}

export function collectChartMarkerImageFillsForCharts(
  charts: readonly ChartModel[],
): ImageFill[] {
  return collectChartImageFillUsagesForCharts(charts).map(usage => usage.fill);
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
  if (!fillCanProduceVisiblePixels(fill) || !(w > 0) || !(h > 0)) return false;
  if (!imageFillModeIsPaintable(fill)) return false;
  const image = activeLookup?.(fill);
  if (!image) return false;
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
