import type {
  ChartModel,
  ChartRect,
  ChartexRegionMap,
  ChartexRegionMapColors,
  ChartexValueColorStop,
} from '../types/chart.js';
import { chartTitleBand, resolveManualLayoutRect } from './layout.js';
import { paintLegendFrame } from './legend-frame.js';
import { paintPlotAreaFrame } from './plot-area-frame.js';
import { formatChartValWithCode } from './chart-number-format.js';
import { chartFontFamily } from './renderer.js';
import {
  NATURAL_EARTH_110M,
  type RegionMapFeature,
  type RegionMapPoint,
} from './region-map-natural-earth.generated.js';

export const REGION_MAP_IMPLEMENTATION_MARKER = 'ooxml-region-map-natural-earth-v1';

type ProjectionType = 'mercator' | 'miller' | 'robinson' | 'albers';
interface Point { x: number; y: number }

const ROBINSON_X = [1, .9986, .9954, .99, .9822, .973, .96, .9427, .9216, .8962, .8679, .835, .7986, .7597, .7186, .6732, .6213, .5722, .5322] as const;
const ROBINSON_Y = [0, .062, .124, .186, .248, .31, .372, .434, .4958, .5571, .6176, .6769, .7346, .7903, .8435, .8936, .9394, .9761, 1] as const;
const radians = (degrees: number) => degrees * Math.PI / 180;
const MAX_REGION_MAP_SOURCE_ROWS = 10_000;

/** Deterministic projection of one lon/lat pair. MS-ODRAWXML enumerates the
 * projection names but does not prescribe projection parameters. Albers uses
 * the conventional world-view standard parallels 20°/50° and central meridian
 * 0°. The schema leaves projectionType optional; Office-produced global maps
 * with the attribute omitted use the bounded Robinson world outline, which is
 * the compatibility fallback here. */
export function projectRegionMapPoint(
  point: RegionMapPoint,
  requested: string | null | undefined,
): Point {
  const projection: ProjectionType = requested === 'mercator' || requested === 'miller'
    || requested === 'albers' || requested === 'robinson'
    ? requested
    : 'robinson';
  const lon = radians(Math.max(-180, Math.min(180, point[0])));
  const latDegrees = Math.max(-89.999, Math.min(89.999, point[1]));
  const lat = radians(latDegrees);
  if (projection === 'mercator') {
    const safeLat = radians(Math.max(-85, Math.min(85, latDegrees)));
    return { x: lon, y: -Math.log(Math.tan(Math.PI / 4 + safeLat / 2)) };
  }
  if (projection === 'miller') {
    return { x: lon, y: -1.25 * Math.log(Math.tan(Math.PI / 4 + .4 * lat)) };
  }
  if (projection === 'albers') {
    const phi1 = radians(20);
    const phi2 = radians(50);
    const n = .5 * (Math.sin(phi1) + Math.sin(phi2));
    const c = Math.cos(phi1) ** 2 + 2 * n * Math.sin(phi1);
    const rho0 = Math.sqrt(c) / n;
    const rho = Math.sqrt(Math.max(0, c - 2 * n * Math.sin(lat))) / n;
    const theta = n * lon;
    return { x: rho * Math.sin(theta), y: rho0 - rho * Math.cos(theta) };
  }
  const absolute = Math.abs(latDegrees);
  const index = Math.min(17, Math.floor(absolute / 5));
  const fraction = Math.min(1, (absolute - index * 5) / 5);
  const interpolate = (table: readonly number[]) => table[index] + (table[index + 1] - table[index]) * fraction;
  return {
    x: .8487 * lon * interpolate(ROBINSON_X),
    y: -1.3523 * Math.sign(latDegrees) * interpolate(ROBINSON_Y),
  };
}

function normalizedKey(value: string): string {
  return value.normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

const FEATURE_BY_NAME = new Map<string, RegionMapFeature>();
const FEATURE_BY_A2 = new Map<string, RegionMapFeature>();
const FEATURE_BY_A3 = new Map<string, RegionMapFeature>();
const FEATURE_BY_POSTAL = new Map<string, RegionMapFeature>();
for (const feature of NATURAL_EARTH_110M) {
  FEATURE_BY_NAME.set(normalizedKey(feature.n), feature);
  if (feature.a2) FEATURE_BY_A2.set(normalizedKey(feature.a2), feature);
  if (feature.a3) FEATURE_BY_A3.set(normalizedKey(feature.a3), feature);
  // Postal codes are the least authoritative identity and collide with ISO
  // codes in Natural Earth (for example Northern Cyprus `p=CN`). Preserve the
  // first value and consult ISO/name maps before this compatibility fallback.
  if (feature.p) {
    const key = normalizedKey(feature.p);
    if (!FEATURE_BY_POSTAL.has(key)) FEATURE_BY_POSTAL.set(key, feature);
  }
}
const ALIASES: Readonly<Record<string, string>> = {
  unitedstatesofamerica: 'United States',
  usa: 'United States',
  uk: 'United Kingdom',
  russia: 'Russian Federation',
  southkorea: 'Republic of Korea',
  northkorea: 'Dem. Rep. Korea',
  czechia: 'Czech Republic',
  ivorycoast: "Côte d'Ivoire",
  drcongo: 'Democratic Republic of the Congo',
  congodemocraticrepublic: 'Democratic Republic of the Congo',
};

export function resolveRegionMapFeature(
  label: string,
  entityId?: string | null,
): RegionMapFeature | undefined {
  const lookup = (key: string): RegionMapFeature | undefined => FEATURE_BY_NAME.get(key)
    ?? FEATURE_BY_A2.get(key) ?? FEATURE_BY_A3.get(key) ?? FEATURE_BY_POSTAL.get(key);
  const identity = entityId ? lookup(normalizedKey(entityId)) : undefined;
  if (identity) return identity;
  const key = normalizedKey(label);
  return lookup(key) ?? (ALIASES[key] ? lookup(normalizedKey(ALIASES[key])) : undefined);
}

function hexRgb(value: string | null | undefined, fallback: string): [number, number, number] {
  const text = (value ?? '').replace(/^#/, '').slice(0, 6);
  if (!/^[0-9a-f]{6}$/i.test(text)) return hexRgb(fallback, '000000');
  return [0, 2, 4].map((offset) => Number.parseInt(text.slice(offset, offset + 2), 16)) as [number, number, number];
}

function rgbHex(rgb: readonly number[]): string {
  return `#${rgb.map((value) => Math.round(Math.max(0, Math.min(255, value))).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

function interpolateColor(a: readonly number[], b: readonly number[], t: number): string {
  return rgbHex(a.map((value, index) => value + (b[index] - value) * Math.max(0, Math.min(1, t))));
}

function rgbToHsl([red, green, blue]: readonly number[]): [number, number, number] {
  const r = red / 255;
  const g = green / 255;
  const b = blue / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  const delta = max - min;
  if (delta === 0) return [0, 0, lightness];
  const saturation = delta / (1 - Math.abs(2 * lightness - 1));
  const rawHue = max === r
    ? ((g - b) / delta) % 6
    : max === g
      ? (b - r) / delta + 2
      : (r - g) / delta + 4;
  return [((rawHue * 60) + 360) % 360, saturation, lightness];
}

function hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number] {
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const segment = ((hue % 360) + 360) % 360 / 60;
  const secondary = chroma * (1 - Math.abs(segment % 2 - 1));
  const [r1, g1, b1] = segment < 1 ? [chroma, secondary, 0]
    : segment < 2 ? [secondary, chroma, 0]
      : segment < 3 ? [0, chroma, secondary]
        : segment < 4 ? [0, secondary, chroma]
          : segment < 5 ? [secondary, 0, chroma]
            : [chroma, 0, secondary];
  const match = lightness - chroma / 2;
  return [r1 + match, g1 + match, b1 + match].map((value) => value * 255) as [number, number, number];
}

/** Excel's omitted Region Map ramp is the first theme accent with DrawingML
 * luminance transforms: min lumMod=20% + lumOff=80%, max lumMod=75%. The
 * transform was verified against Office vector output and keeps a custom theme
 * authoritative instead of freezing the observed blue RGB values. */
function defaultRegionMapRamp(themeAccent: string | null | undefined): { min: string; max: string } {
  const [hue, saturation, lightness] = rgbToHsl(hexRgb(themeAccent, '4472C4'));
  return {
    min: rgbHex(hslToRgb(hue, saturation, lightness * .2 + .8)),
    max: rgbHex(hslToRgb(hue, saturation, lightness * .75)),
  };
}

function finiteLerp(a: number, b: number, t: number): number {
  const clamped = Math.max(0, Math.min(1, t));
  const direct = a + (b - a) * clamped;
  if (Number.isFinite(direct)) return direct;
  const magnitude = Math.max(Math.abs(a), Math.abs(b));
  if (!(magnitude > 0) || !Number.isFinite(magnitude)) return 0;
  return ((a / magnitude) * (1 - clamped) + (b / magnitude) * clamped) * magnitude;
}

function finiteRatio(value: number, low: number, high: number): number {
  if (low === high) return .5;
  const direct = (value - low) / (high - low);
  if (Number.isFinite(direct)) return direct;
  const magnitude = Math.max(Math.abs(value), Math.abs(low), Math.abs(high));
  if (!(magnitude > 0) || !Number.isFinite(magnitude)) return .5;
  return (value / magnitude - low / magnitude) / (high / magnitude - low / magnitude);
}

function stopValue(
  stop: ChartexValueColorStop | null | undefined,
  fallback: number,
  min: number,
  max: number,
): number {
  if (!stop || stop.kind === 'extremeValue') return fallback;
  if (stop.kind === 'percent' && Number.isFinite(stop.value)) return finiteLerp(min, max, (stop.value as number) / 100);
  if (stop.kind === 'number' && Number.isFinite(stop.value)) return stop.value as number;
  return fallback;
}

interface ColorScale { color(value: number): string; min: number; max: number; minColor: string; midColor?: string; maxColor: string }

export function regionMapColorScale(
  values: readonly number[],
  authored?: ChartexRegionMapColors | null,
  themeAccent?: string | null,
): ColorScale {
  let min = Infinity;
  let max = -Infinity;
  for (const value of values) {
    if (!Number.isFinite(value)) continue;
    min = Math.min(min, value);
    max = Math.max(max, value);
  }
  if (!Number.isFinite(min) || !Number.isFinite(max)) {
    min = 0;
    max = 1;
  }
  const automatic = defaultRegionMapRamp(themeAccent);
  const minColor = rgbHex(hexRgb(authored?.minColor, automatic.min));
  // CT_ValueColorPositions defaults count to 2; a present midColor is not a
  // third stop unless the authored position contract explicitly says count=3.
  const midColor = authored?.stopCount === 3 && authored.midColor
    ? rgbHex(hexRgb(authored.midColor, '5B9BD5')) : undefined;
  const maxColor = rgbHex(hexRgb(authored?.maxColor, automatic.max));
  const minPosition = stopValue(authored?.minPosition, min, min, max);
  const maxPosition = stopValue(authored?.maxPosition, max, min, max);
  const midPosition = stopValue(authored?.midPosition, (minPosition + maxPosition) / 2, min, max);
  const low = Math.min(minPosition, maxPosition);
  const high = Math.max(minPosition, maxPosition);
  return {
    min,
    max,
    minColor,
    midColor,
    maxColor,
    color(value) {
      if (midColor) {
        return value <= midPosition
          ? interpolateColor(hexRgb(minColor, minColor), hexRgb(midColor, midColor), finiteRatio(value, low, midPosition))
          : interpolateColor(hexRgb(midColor, midColor), hexRgb(maxColor, maxColor), finiteRatio(value, midPosition, high));
      }
      return interpolateColor(hexRgb(minColor, minColor), hexRgb(maxColor, maxColor), finiteRatio(value, low, high));
    },
  };
}

function aggregateRows(map: ChartexRegionMap): Map<RegionMapFeature, number> {
  const result = new Map<RegionMapFeature, number>();
  for (let index = 0; index < map.rows.length; index++) {
    const row = map.rows[index];
    if (!Number.isFinite(row.value)) continue;
    const feature = resolveRegionMapFeature(row.label, row.entityId);
    if (!feature) continue;
    const previous = result.get(feature) ?? 0;
    const value = row.value as number;
    const next = previous + value;
    const bounded = Number.isFinite(next)
      ? next
      : previous > 0 && value > 0
        ? Number.MAX_VALUE
        : previous < 0 && value < 0
          ? -Number.MAX_VALUE
          : 0;
    result.set(feature, bounded);
  }
  return result;
}

function drawTitle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  fontPx: number,
  topPad: number,
): void {
  if (!chart.title && !chart.titlePresent) return;
  const family = chartFontFamily(chart, chart.titleFontFace, 'major');
  ctx.font = `${chart.titleFontBold ?? true ? 'bold ' : ''}${fontPx}px ${family}`;
  ctx.fillStyle = chart.titleFontColor ? `#${chart.titleFontColor}` : '#333333';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillText(chart.title ?? '', rect.x + rect.w / 2, rect.y + topPad);
}

function drawLegend(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  scale: ColorScale,
  x: number,
  y: number,
  w: number,
  h: number,
  ptToPx: number,
): void {
  const name = chart.series[0]?.name ?? '';
  const fontPx = Math.max(8, (chart.legendFontSizeHpt ?? 900) / 100 * ptToPx);
  ctx.font = `${chart.legendFontBold ? 'bold ' : ''}${fontPx}px ${chartFontFamily(chart, chart.legendFontFace, 'minor')}`;
  ctx.fillStyle = chart.legendFontColor ? `#${chart.legendFontColor}` : '#595959';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  if (name) ctx.fillText(name, x + w / 2, y);
  const barW = Math.min(w * .58, 240);
  const barH = Math.max(6, Math.min(12, h * .25));
  const bx = x + (w - barW) / 2;
  const by = y + (name ? fontPx + 2 : 0);
  const gradient = ctx.createLinearGradient(bx, 0, bx + barW, 0);
  gradient.addColorStop(0, scale.minColor);
  if (scale.midColor) gradient.addColorStop(.5, scale.midColor);
  gradient.addColorStop(1, scale.maxColor);
  ctx.fillStyle = gradient;
  ctx.fillRect(bx, by, barW, barH);
  ctx.fillStyle = chart.legendFontColor ? `#${chart.legendFontColor}` : '#595959';
  ctx.textBaseline = 'top';
  ctx.textAlign = 'left';
  const formatCode = chart.series[0]?.valFormatCode ?? null;
  ctx.fillText(formatChartValWithCode(scale.min, formatCode, chart.date1904), bx, by + barH + 2);
  ctx.textAlign = 'right';
  ctx.fillText(formatChartValWithCode(scale.max, formatCode, chart.date1904), bx + barW, by + barH + 2);
}

function drawUnavailableMessage(
  ctx: CanvasRenderingContext2D,
  rect: ChartRect,
  ptToPx: number,
  message: string,
): void {
  ctx.fillStyle = '#666666';
  ctx.font = `${Math.max(9, 9 * ptToPx)}px sans-serif`;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(message, rect.x + rect.w / 2, rect.y + rect.h / 2);
}

/** Paint a deterministic, network-free country-level Region Map. */
export function renderRegionMapChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  rect: ChartRect,
  ptToPx: number,
  shapeRotationDeg = 0,
): boolean {
  if (chart.chartType !== 'regionMap' || !chart.chartexRegionMap) return false;
  const map = chart.chartexRegionMap;
  if (map.rows.length > MAX_REGION_MAP_SOURCE_ROWS) {
    drawUnavailableMessage(ctx, rect, ptToPx, '(chart values exceed rendering limit)');
    return true;
  }
  // Clear/binary geoCache contents can carry the authoritative mapping from
  // localized/ambiguous queries to ISO/entity identities. The parser currently
  // retains only cache provenance, so falling back to English name guesses
  // would color the wrong country. Keep cached maps explicit and fail closed
  // until that bounded identity table is modeled.
  if (map.geography?.cachePresent) {
    drawUnavailableMessage(ctx, rect, ptToPx, '(region map cache is unavailable offline)');
    return true;
  }
  // The fixed Natural Earth asset is country-level. ST_GeoMappingLevel also
  // permits state/county/postal views, but guessing those geometries from a
  // country table would silently color the wrong region. Fail closed until an
  // explicit bounded sub-country asset is supplied by a future renderer module.
  const viewedRegionType = map.geography?.viewedRegionType;
  if (viewedRegionType != null && viewedRegionType !== 'world') {
    // `dataOnly`, country-region lists and sub-country levels each define a
    // distinct authored viewport. The fixed country asset can paint their
    // regions, but fitting all of them to the world would silently discard the
    // view contract. Keep them fail-closed until an Office-observed viewport
    // planner is available instead of inventing a sample-specific crop.
    drawUnavailableMessage(ctx, rect, ptToPx, '(region map detail is unavailable offline)');
    return true;
  }
  const resolved = aggregateRows(map);
  const colorScale = regionMapColorScale(
    [...resolved.values()],
    map.colors,
    chart.chartexAccents?.[0],
  );
  const title = chartTitleBand(chart, rect.h, ptToPx, .02, .015);
  drawTitle(ctx, chart, rect, title.fontPx, title.topPad);
  const legendH = chart.showLegend ? Math.max(32, rect.h * .16) : 0;
  const defaultLegendRect = chart.showLegend ? {
    x: rect.x,
    y: rect.y + title.bandH,
    w: rect.w,
    h: legendH,
  } : null;
  const legendRect = defaultLegendRect && chart.legendManualLayout
    ? resolveManualLayoutRect(chart.legendManualLayout, rect, defaultLegendRect)
      ?? defaultLegendRect
    : defaultLegendRect;
  const reservedLegendH = chart.legendOverlay === true ? 0 : legendH;
  const insetX = rect.w * .03;
  const insetBottom = rect.h * .035;
  const automaticMapRect = {
    x: rect.x + insetX,
    y: rect.y + title.bandH + reservedLegendH,
    w: Math.max(1, rect.w - insetX * 2),
    h: Math.max(1, rect.h - title.bandH - reservedLegendH - insetBottom),
  };
  const mapRect = chart.plotAreaManualLayout
    ? resolveManualLayoutRect(chart.plotAreaManualLayout, rect, automaticMapRect)
      ?? automaticMapRect
    : automaticMapRect;
  paintPlotAreaFrame(
    ctx, chart, mapRect.x, mapRect.y, mapRect.w, mapRect.h, ptToPx, shapeRotationDeg,
  );
  if (legendRect) {
    paintLegendFrame(
      ctx, chart, legendRect, ptToPx, shapeRotationDeg,
    );
    drawLegend(
      ctx, chart, colorScale,
      legendRect.x, legendRect.y, legendRect.w, legendRect.h,
      ptToPx,
    );
  }

  const projection = map.geography?.projectionType;
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (const feature of NATURAL_EARTH_110M) {
    for (const polygon of feature.g) for (const ring of polygon) for (const point of ring) {
      const projected = projectRegionMapPoint(point, projection);
      minX = Math.min(minX, projected.x); maxX = Math.max(maxX, projected.x);
      minY = Math.min(minY, projected.y); maxY = Math.max(maxY, projected.y);
    }
  }
  const scale = Math.min(mapRect.w / Math.max(Number.MIN_VALUE, maxX - minX), mapRect.h / Math.max(Number.MIN_VALUE, maxY - minY));
  const offsetX = mapRect.x + (mapRect.w - (maxX - minX) * scale) / 2 - minX * scale;
  const offsetY = mapRect.y + (mapRect.h - (maxY - minY) * scale) / 2 - minY * scale;
  const canvasPoint = (point: RegionMapPoint) => {
    const projected = projectRegionMapPoint(point, projection);
    return { x: offsetX + projected.x * scale, y: offsetY + projected.y * scale };
  };

  const featureBounds = new Map<RegionMapFeature, { minX: number; minY: number; maxX: number; maxY: number }>();
  ctx.lineWidth = Math.max(.35, .55 * ptToPx);
  ctx.strokeStyle = '#FFFFFF';
  for (const feature of NATURAL_EARTH_110M) {
    ctx.beginPath();
    let bounds = { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity };
    for (const polygon of feature.g) for (const ring of polygon) {
      let previous: RegionMapPoint | undefined;
      for (let index = 0; index < ring.length; index++) {
        const point = ring[index];
        const canvas = canvasPoint(point);
        bounds.minX = Math.min(bounds.minX, canvas.x); bounds.maxX = Math.max(bounds.maxX, canvas.x);
        bounds.minY = Math.min(bounds.minY, canvas.y); bounds.maxY = Math.max(bounds.maxY, canvas.y);
        if (index === 0 || (previous && Math.abs(point[0] - previous[0]) > 180)) ctx.moveTo(canvas.x, canvas.y);
        else ctx.lineTo(canvas.x, canvas.y);
        previous = point;
      }
      ctx.closePath();
    }
    ctx.fillStyle = resolved.has(feature) ? colorScale.color(resolved.get(feature) as number) : '#E0E0E0';
    ctx.fill('evenodd');
    ctx.stroke();
    featureBounds.set(feature, bounds);
  }

  if (map.regionLabelLayout && map.regionLabelLayout !== 'none') {
    ctx.font = `${Math.max(7, (chart.dataLabelFontSizeHpt ?? 800) / 100 * ptToPx)}px ${chartFontFamily(chart, chart.dataLabelFontFace, 'minor')}`;
    ctx.fillStyle = chart.dataLabelFontColor ? `#${chart.dataLabelFontColor}` : '#404040';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    for (const row of map.rows) {
      const feature = resolveRegionMapFeature(row.label, row.entityId);
      const bounds = feature ? featureBounds.get(feature) : undefined;
      if (!feature || !bounds || !row.label) continue;
      const point = canvasPoint(feature.l);
      const fits = ctx.measureText(row.label).width <= Math.max(0, bounds.maxX - bounds.minX - 4);
      if (map.regionLabelLayout === 'showAll' || fits) ctx.fillText(row.label, point.x, point.y);
    }
  }
  if (map.geography?.attribution) {
    ctx.font = `${Math.max(7, 7 * ptToPx)}px sans-serif`;
    ctx.fillStyle = '#777777';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'bottom';
    ctx.fillText(map.geography.attribution, rect.x + rect.w - 4, rect.y + rect.h - 2);
  }
  return true;
}
