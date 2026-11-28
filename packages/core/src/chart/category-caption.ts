import type { ChartModel, ChartSeries } from '../types/chart.js';
import { formatCategoryLabel } from './chart-number-format.js';

/**
 * The one resolver every category-text paint site goes through. It returns ONE
 * resolved record per point, taken from the single source that supplied the
 * text: the text, whether that source is numeric (`c:numRef` / `c:numLit`) or
 * text (`c:strRef` / `c:strLit`), that source's own number-format facts, and
 * who owns it. No paint site looks into caches itself or takes a format from a
 * different series.
 *
 * Owners: `series` (the series' own cache covers the point), `shared` (the
 * chart's shared list, the first parsed series' source, with the format facts
 * stored beside it so they survive series filtering) and `generated` (no
 * source has the point; the text follows the caller's fallback policy).
 *
 * Formatting: text is shown verbatim; a number follows the format code
 * supplied by the site (axis ticks: the authored axis code) or, for point
 * labels and legends, the source's own code, else General with the host
 * decimal separator. Office evidence for a data label or legend inheriting
 * the category AXIS code is lacking, so they do not.
 *
 * Out of scope, deliberately not routed (they paint numbers or text that has
 * no category source record): generated date-axis ticks (`tick.serial`:
 * numeric by construction, formatted with `numeric = true`), multi-level
 * category text (authored strings, painted verbatim), ChartEx charts (their
 * own category model and renderer), and the bar renderer's ChartEx-style
 * waterfall / Pareto helper paths.
 */
export interface CategoryRecord {
  readonly text: string;
  /** `true` numeric source, `false` text source, absent: unknown provenance. */
  readonly numeric: boolean | undefined;
  /** The source's own number-format code for this point (per-point, else cache). */
  readonly formatCode: string | null;
  /** Worksheet built-in number-format id of the source's first cell. */
  readonly builtinFormatId: number | undefined;
  readonly owner: 'series' | 'shared' | 'generated';
}

/** What to paint where no source has the point. `none`: empty text; `chart`:
 * the chart's generated 1..n axis labels (ECMA-376 §21.2.2.24, bounded by the
 * longest series); `index`: always `index + 1`. */
export type CategoryFallback = 'none' | 'chart' | 'index';

const longestSeriesCache = new WeakMap<ChartModel, number>();

function longestSeries(chart: ChartModel): number {
  let longest = longestSeriesCache.get(chart);
  if (longest === undefined) {
    longest = 0;
    for (const series of chart.series) if (series.values.length > longest) longest = series.values.length;
    longestSeriesCache.set(chart, longest);
  }
  return longest;
}

/** The shared category source: the chart-level list with its format facts, or
 * the first series' own list when the chart-level list is empty. */
export interface SharedCategorySource {
  readonly categories: readonly string[];
  readonly numeric: boolean | undefined;
  readonly formatCode: string | undefined;
  readonly formatCodes: readonly (string | null)[] | undefined;
  readonly builtinFormatId: number | undefined;
}

export function sharedCategorySource(chart: ChartModel): SharedCategorySource {
  if (chart.categories.length > 0) {
    // A model from the parsers states its shared source's format beside the
    // categories (it names its provenance). A hand-built or older model does
    // not: its first series' cache carries the source's format.
    const legacy = chart.categoriesNumeric === undefined ? chart.series[0] : undefined;
    return {
      categories: chart.categories,
      numeric: chart.categoriesNumeric,
      formatCode: chart.categoryFormatCode ?? legacy?.catFormatCode ?? undefined,
      formatCodes: chart.categoryFormatCodes ?? legacy?.catFormatCodes ?? undefined,
      builtinFormatId: chart.categoryBuiltinFormatId ?? legacy?.catFormatBuiltinId ?? undefined,
    };
  }
  const first = chart.series[0];
  if (first?.categories && first.categories.length > 0) {
    return {
      categories: first.categories,
      numeric: first.categoriesNumeric,
      formatCode: first.catFormatCode ?? undefined,
      formatCodes: first.catFormatCodes ?? undefined,
      builtinFormatId: first.catFormatBuiltinId ?? undefined,
    };
  }
  return { categories: [], numeric: undefined, formatCode: undefined, formatCodes: undefined, builtinFormatId: undefined };
}

/** `series` carrying the shared category source as its own: the labels with
 * their provenance and format facts, for a series that shares the chart's
 * categories (scatter series filtered by visibility, pie / legend spreads). */
export function seriesWithSharedCategories(chart: ChartModel, series: ChartSeries): ChartSeries {
  const shared = sharedCategorySource(chart);
  return {
    ...series,
    categories: shared.categories as string[],
    categoriesNumeric: shared.numeric,
    catFormatCode: shared.formatCode ?? null,
    catFormatCodes: (shared.formatCodes as (string | null)[] | undefined) ?? null,
    catFormatBuiltinId: shared.builtinFormatId ?? null,
  };
}

/** `series` as displayed: its own categories where it has them, else the
 * shared source carried onto it (legend and label painters read one series). */
export function seriesWithDisplayedCategories(chart: ChartModel, series: ChartSeries): ChartSeries {
  return series.categories && series.categories.length > 0 ? series : seriesWithSharedCategories(chart, series);
}

/** The resolved record of category `index`. `series` `null`: the chart's
 * category axis (shared list). */
export function categorySourceAt(
  chart: ChartModel,
  series: ChartSeries | null | undefined,
  index: number,
  fallback: CategoryFallback = series ? 'none' : 'chart',
): CategoryRecord {
  const own = series?.categories?.[index];
  if (series && own !== undefined) {
    return {
      text: String(own),
      numeric: series.categoriesNumeric,
      formatCode: series.catFormatCodes?.[index] ?? series.catFormatCode ?? null,
      builtinFormatId: series.catFormatBuiltinId ?? undefined,
      owner: 'series',
    };
  }
  const shared = sharedCategorySource(chart);
  const value = shared.categories[index];
  if (value !== undefined) {
    return {
      text: String(value),
      numeric: shared.numeric,
      formatCode: shared.formatCodes?.[index] ?? shared.formatCode ?? null,
      builtinFormatId: shared.builtinFormatId,
      owner: 'shared',
    };
  }
  const text = fallback === 'index'
    ? String(index + 1)
    : fallback === 'chart' && index < longestSeries(chart) ? String(index + 1) : '';
  return { text, numeric: undefined, formatCode: null, builtinFormatId: undefined, owner: 'generated' };
}

export interface CategoryCaptionFormat {
  readonly code?: string | null;
  readonly divisors?: readonly number[] | null;
}

/** Format a resolved record under the site's code (point labels and legends
 * pass the record's own `formatCode`). */
export function formatCategoryRecord(
  chart: ChartModel,
  record: CategoryRecord,
  format: CategoryCaptionFormat = {},
): string {
  return formatCategoryLabel(record.text, format.code, chart.date1904, format.divisors, record.numeric);
}

/** Category-axis tick text of `index`: the shared categories under the axis code. */
export function axisCategoryCaption(chart: ChartModel, index: number, fallback: CategoryFallback = 'chart'): string {
  return formatCategoryRecord(chart, categorySourceAt(chart, null, index, fallback), {
    code: chart.catAxisFormatCode, divisors: chart.catAxisFormatDivisors,
  });
}

/** A point's category text (data labels, legends): the source's own format
 * code, else General. */
export function pointCategoryCaption(
  chart: ChartModel,
  series: ChartSeries,
  index: number,
  fallback: CategoryFallback = 'chart',
): string {
  const record = categorySourceAt(chart, series, index, fallback);
  return formatCategoryRecord(chart, record, { code: record.formatCode });
}

export function pointCategoryCaptionAt(
  chart: ChartModel,
  series: ChartSeries,
  fallback: CategoryFallback = 'chart',
): (index: number) => string {
  return index => pointCategoryCaption(chart, series, index, fallback);
}
