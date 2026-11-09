/**
 * OpenType font variations for line metrics: `fvar` axes, `avar` segment
 * maps and the `MVAR` deltas of the global vertical metrics
 * (https://learn.microsoft.com/en-us/typography/opentype/spec/otvaroverview,
 * .../fvar, .../avar, .../mvar). Only the values a line box needs are
 * instanced: hhea ascender/descender/lineGap (`hasc`/`hdsc`/`hlgp`), OS/2
 * usWin (`hcla`/`hcld`) and OS/2 typo (`tasc`/`tdsc`/`tlgp`).
 */

export interface OpenTypeVariationAxis {
  readonly tag: string;
  readonly min: number;
  readonly default: number;
  readonly max: number;
}

interface Region { readonly axes: readonly (readonly [start: number, peak: number, end: number])[] }

/** One MVAR value: per-region deltas of the variation data it selects. */
interface MetricDeltas { readonly regions: readonly number[]; readonly deltas: readonly number[] }

export interface OpenTypeVariation {
  readonly axes: readonly OpenTypeVariationAxis[];
  /** Per axis (fvar order): [fromCoord, toCoord] pairs, or null for identity. */
  readonly avar: readonly (readonly (readonly [number, number])[] | null)[];
  readonly regions: readonly Region[];
  /** MVAR deltas by value tag; empty when the font has no MVAR. */
  readonly metrics: ReadonlyMap<string, MetricDeltas>;
}

export const LINE_METRIC_MVAR_TAGS = ['hasc', 'hdsc', 'hlgp', 'hcla', 'hcld', 'tasc', 'tdsc', 'tlgp'] as const;
export type LineMetricMvarTag = typeof LINE_METRIC_MVAR_TAGS[number];

type Table = { offset: number; length: number } | undefined;

const tagString = (view: DataView, offset: number) => String.fromCharCode(
  view.getUint8(offset), view.getUint8(offset + 1), view.getUint8(offset + 2), view.getUint8(offset + 3));

const fixed = (view: DataView, offset: number) => view.getInt32(offset) / 65536;
const f2dot14 = (view: DataView, offset: number) => view.getInt16(offset) / 16384;

function fits(table: { offset: number; length: number }, at: number, size: number): boolean {
  return at >= 0 && size >= 0 && at + size <= table.length;
}

function parseAxes(view: DataView, fvar: { offset: number; length: number }): OpenTypeVariationAxis[] | null {
  if (!fits(fvar, 0, 16)) return null;
  const axesArrayOffset = view.getUint16(fvar.offset + 4);
  const axisCount = view.getUint16(fvar.offset + 8);
  const axisSize = view.getUint16(fvar.offset + 10);
  if (axisSize < 20 || axisCount === 0 || !fits(fvar, axesArrayOffset, axisCount * axisSize)) return null;
  const axes: OpenTypeVariationAxis[] = [];
  for (let index = 0; index < axisCount; index++) {
    const at = fvar.offset + axesArrayOffset + index * axisSize;
    const axis = { tag: tagString(view, at), min: fixed(view, at + 4), default: fixed(view, at + 8), max: fixed(view, at + 12) };
    if (!(axis.min <= axis.default && axis.default <= axis.max)) return null;
    axes.push(Object.freeze(axis));
  }
  return axes;
}

function parseAvar(view: DataView, avar: Table, axisCount: number): (readonly (readonly [number, number])[] | null)[] | null {
  const identity = Array.from({ length: axisCount }, () => null);
  if (!avar) return identity;
  if (!fits(avar, 0, 8) || view.getUint16(avar.offset) !== 1) return null;
  if (view.getUint16(avar.offset + 6) !== axisCount) return null;
  const maps: (readonly (readonly [number, number])[] | null)[] = [];
  let at = 8;
  for (let axis = 0; axis < axisCount; axis++) {
    if (!fits(avar, at, 2)) return null;
    const count = view.getUint16(avar.offset + at);
    if (!fits(avar, at + 2, count * 4)) return null;
    const pairs: [number, number][] = [];
    for (let index = 0; index < count; index++) {
      const pair = avar.offset + at + 2 + index * 4;
      pairs.push([f2dot14(view, pair), f2dot14(view, pair + 2)]);
    }
    at += 2 + count * 4;
    maps.push(pairs.length >= 3 ? Object.freeze(pairs) : null);
  }
  return maps;
}

function parseRegions(view: DataView, store: number, end: number, axisCount: number): Region[] | null {
  if (store + 4 > end) return null;
  const regionAxisCount = view.getUint16(store);
  const regionCount = view.getUint16(store + 2);
  if (regionAxisCount !== axisCount || store + 4 + regionCount * axisCount * 6 > end) return null;
  const regions: Region[] = [];
  for (let region = 0; region < regionCount; region++) {
    const axes: (readonly [number, number, number])[] = [];
    for (let axis = 0; axis < axisCount; axis++) {
      const at = store + 4 + (region * axisCount + axis) * 6;
      axes.push([f2dot14(view, at), f2dot14(view, at + 2), f2dot14(view, at + 4)]);
    }
    regions.push(Object.freeze({ axes: Object.freeze(axes) }));
  }
  return regions;
}

function itemDeltas(view: DataView, data: number, end: number, inner: number): MetricDeltas | null {
  if (data + 6 > end) return null;
  const itemCount = view.getUint16(data);
  const wordField = view.getUint16(data + 2);
  const longWords = (wordField & 0x8000) !== 0;
  const wordCount = wordField & 0x7fff;
  const regionIndexCount = view.getUint16(data + 4);
  if (wordCount > regionIndexCount || inner >= itemCount) return null;
  const regionIndexes = data + 6;
  const rowSize = longWords
    ? wordCount * 4 + (regionIndexCount - wordCount) * 2
    : wordCount * 2 + (regionIndexCount - wordCount);
  const rows = regionIndexes + regionIndexCount * 2;
  if (rows + itemCount * rowSize > end) return null;
  const regions: number[] = [];
  const deltas: number[] = [];
  let at = rows + inner * rowSize;
  for (let index = 0; index < regionIndexCount; index++) {
    regions.push(view.getUint16(regionIndexes + index * 2));
    if (index < wordCount) {
      deltas.push(longWords ? view.getInt32(at) : view.getInt16(at));
      at += longWords ? 4 : 2;
    } else {
      deltas.push(longWords ? view.getInt16(at) : view.getInt8(at));
      at += longWords ? 2 : 1;
    }
  }
  return { regions: Object.freeze(regions), deltas: Object.freeze(deltas) };
}

/** Parse fvar/avar/MVAR; null when any present table is malformed. */
export function parseOpenTypeVariation(
  view: DataView,
  tables: ReadonlyMap<number, { offset: number; length: number }>,
  tag: (name: string) => number,
): OpenTypeVariation | null {
  const fvar = tables.get(tag('fvar'));
  if (!fvar) return null;
  const axes = parseAxes(view, fvar);
  if (!axes) return null;
  const avar = parseAvar(view, tables.get(tag('avar')), axes.length);
  if (!avar) return null;
  const mvar = tables.get(tag('MVAR'));
  const metrics = new Map<string, MetricDeltas>();
  let regions: Region[] = [];
  if (mvar) {
    if (!fits(mvar, 0, 12) || view.getUint16(mvar.offset) !== 1) return null;
    const recordSize = view.getUint16(mvar.offset + 6);
    const recordCount = view.getUint16(mvar.offset + 8);
    const storeOffset = view.getUint16(mvar.offset + 10);
    if (recordCount > 0) {
      if (recordSize < 8 || !fits(mvar, 12, recordCount * recordSize) || !storeOffset) return null;
      const end = mvar.offset + mvar.length;
      const store = mvar.offset + storeOffset;
      if (store + 8 > end || view.getUint16(store) !== 1) return null;
      const regionList = store + view.getUint32(store + 2);
      const dataCount = view.getUint16(store + 6);
      if (store + 8 + dataCount * 4 > end) return null;
      const parsedRegions = parseRegions(view, regionList, end, axes.length);
      if (!parsedRegions) return null;
      regions = parsedRegions;
      for (let index = 0; index < recordCount; index++) {
        const at = mvar.offset + 12 + index * recordSize;
        const valueTag = tagString(view, at);
        if (!(LINE_METRIC_MVAR_TAGS as readonly string[]).includes(valueTag)) continue;
        const outer = view.getUint16(at + 4);
        const inner = view.getUint16(at + 6);
        // 0xFFFF/0xFFFF is the "no variation data" index (OpenType 1.8.x).
        if (outer === 0xffff && inner === 0xffff) continue;
        if (outer >= dataCount) return null;
        const deltas = itemDeltas(view, store + view.getUint32(store + 8 + outer * 4), end, inner);
        if (!deltas || deltas.regions.some((region) => region >= regions.length)) return null;
        metrics.set(valueTag, deltas);
      }
    }
  }
  return Object.freeze({ axes: Object.freeze(axes), avar: Object.freeze(avar), regions: Object.freeze(regions), metrics });
}

/** Normalized coordinate of a user-space axis value (fvar default
 * normalization followed by the avar segment map). */
function normalize(axis: OpenTypeVariationAxis, map: readonly (readonly [number, number])[] | null, value: number): number {
  const clamped = Math.min(axis.max, Math.max(axis.min, value));
  let normalized = clamped < axis.default
    ? (axis.default === axis.min ? 0 : (clamped - axis.default) / (axis.default - axis.min))
    : clamped > axis.default
      ? (axis.max === axis.default ? 0 : (clamped - axis.default) / (axis.max - axis.default))
      : 0;
  if (map) {
    for (let index = 1; index < map.length; index++) {
      const [fromLow, toLow] = map[index - 1]!;
      const [fromHigh, toHigh] = map[index]!;
      if (normalized <= fromHigh) {
        normalized = fromHigh === fromLow ? toLow
          : toLow + (toHigh - toLow) * (normalized - fromLow) / (fromHigh - fromLow);
        break;
      }
    }
  }
  return normalized;
}

function regionScalar(region: Region, coords: readonly number[]): number {
  let scalar = 1;
  region.axes.forEach(([start, peak, end], axis) => {
    const coord = coords[axis] ?? 0;
    if (start > peak || peak > end || (start < 0 && end > 0 && peak !== 0) || peak === 0 || coord === peak) return;
    if (coord <= start || coord >= end) { scalar = 0; return; }
    scalar *= coord < peak ? (coord - start) / (peak - start) : (end - coord) / (end - peak);
  });
  return scalar;
}

/** Normalized coordinates of an instance given by user-space axis values;
 * an axis without a value takes its default. */
export function variationCoordinates(variation: OpenTypeVariation, values: Readonly<Record<string, number>>): number[] {
  return variation.axes.map((axis, index) => {
    const value = Object.prototype.hasOwnProperty.call(values, axis.tag) ? values[axis.tag]! : axis.default;
    return normalize(axis, variation.avar[index] ?? null, value);
  });
}

/** Whether any line-metric delta varies along `axisTag` (e.g. `opsz`). */
export function lineMetricsVaryAlong(variation: OpenTypeVariation, axisTag: string): boolean {
  const axis = variation.axes.findIndex((candidate) => candidate.tag === axisTag);
  if (axis < 0) return false;
  for (const { regions, deltas } of variation.metrics.values()) {
    if (regions.some((region, index) => deltas[index] !== 0 && variation.regions[region]!.axes[axis]![1] !== 0)) return true;
  }
  return false;
}

/** Summed MVAR delta of one value tag at normalized `coords`. */
export function metricDelta(variation: OpenTypeVariation, valueTag: LineMetricMvarTag, coords: readonly number[]): number {
  const entry = variation.metrics.get(valueTag);
  if (!entry) return 0;
  let delta = 0;
  entry.regions.forEach((region, index) => {
    delta += entry.deltas[index]! * regionScalar(variation.regions[region]!, coords);
  });
  return delta;
}
