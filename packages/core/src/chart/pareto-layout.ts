import type {
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartSeries,
} from '../types/chart';

export interface ParetoPoint {
  sourceIndex: number;
  category: string;
  value: number;
  cumulativeFraction: number;
}

export interface ParetoLayout {
  points: ParetoPoint[];
  /** Source series reordered for the frequency bars. */
  orderedSeries: ChartSeries;
  /** Reordered series whose values are the cumulative 0..1 fractions. */
  series: ChartSeries;
  categories: string[];
}

export interface ParetoLayoutOptions {
  /** Owner-backed Pareto bars sort frequencies; a standalone paretoLine does not. */
  sortDescending?: boolean;
}

function remapIndexed<T extends { idx: number }>(
  values: readonly T[] | null | undefined,
  newIndexBySource: ReadonlyMap<number, number>,
): T[] | null | undefined {
  if (values == null) return values;
  return values.flatMap(value => {
    const idx = newIndexBySource.get(value.idx);
    return idx == null ? [] : [{ ...value, idx }];
  });
}

function reorderNullable<T>(
  values: readonly T[] | null | undefined,
  sourceIndices: readonly number[],
): Array<T | null> | null | undefined {
  if (values == null) return values;
  return sourceIndices.map(index => values[index] ?? null);
}

/**
 * Derive a deterministic Pareto order without mutating the authored model.
 *
 * Finite non-negative values participate. Ties retain source order, missing or
 * invalid values are omitted, and a zero-total series produces finite zero
 * cumulative values. Indexed point/label properties follow their source point
 * through the reorder.
 */
export function planParetoLayout(
  series: ChartSeries,
  chartCategories: readonly string[],
  options: ParetoLayoutOptions = {},
): ParetoLayout {
  const retained = series.values
    .map((value, sourceIndex) => ({ value, sourceIndex }))
    .filter((entry): entry is { value: number; sourceIndex: number } =>
      entry.value != null && Number.isFinite(entry.value) && entry.value >= 0
    );
  if (options.sortDescending !== false) {
    retained.sort((a, b) => b.value - a.value || a.sourceIndex - b.sourceIndex);
  }

  // Normalize before summing so finite values near Number.MAX_VALUE cannot
  // overflow the denominator to Infinity and collapse early fractions to 0.
  const scale = retained[0]?.value ?? 0;
  const normalizedTotal = scale > 0
    ? retained.reduce((sum, entry) => sum + entry.value / scale, 0)
    : 0;
  let running = 0;
  const points = retained.map((entry): ParetoPoint => {
    if (scale > 0) running += entry.value / scale;
    return {
      sourceIndex: entry.sourceIndex,
      category: series.categories?.[entry.sourceIndex]
        ?? chartCategories[entry.sourceIndex]
        ?? String(entry.sourceIndex + 1),
      value: entry.value,
      cumulativeFraction: normalizedTotal > 0
        ? (running >= normalizedTotal ? 1 : running / normalizedTotal)
        : 0,
    };
  });
  const sourceIndices = points.map(point => point.sourceIndex);
  const newIndexBySource = new Map(
    sourceIndices.map((sourceIndex, newIndex) => [sourceIndex, newIndex]),
  );
  const categories = points.map(point => point.category);
  const reordered = {
    ...series,
    categories,
    catFormatCodes: reorderNullable(series.catFormatCodes, sourceIndices),
    dataPointColors: reorderNullable(series.dataPointColors, sourceIndices),
    dataLabelColors: reorderNullable(series.dataLabelColors, sourceIndices),
    dataPointOverrides: remapIndexed<ChartDataPointOverride>(
      series.dataPointOverrides,
      newIndexBySource,
    ),
    dataLabelOverrides: remapIndexed<ChartDataLabelOverride>(
      series.dataLabelOverrides,
      newIndexBySource,
    ),
  };

  return {
    points,
    categories,
    orderedSeries: {
      ...reordered,
      values: points.map(point => point.value),
    },
    series: {
      ...reordered,
      values: points.map(point => point.cumulativeFraction),
    },
  };
}
