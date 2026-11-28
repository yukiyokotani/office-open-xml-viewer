import type { ChartModel, ChartSeries } from '../types/chart.js';
import { chartCategories } from './category-spacing.js';
import { formatCategoryLabel } from './chart-number-format.js';

/**
 * The one resolver every category-text paint site goes through (axis ticks of
 * every family, secondary and 3-D axes, data-table headers, data labels,
 * legends). A category's text and whether it came from a numeric source
 * (`c:numRef` / `c:numLit`) or a text source (`c:strRef` / `c:strLit`) are
 * resolved together from the one source actually displayed for that point:
 * the series' own categories while its cache covers the point, else the
 * chart's shared categories (itself the first series' when the chart-level
 * list is empty). Text is shown verbatim; a number follows the authored
 * format code (existing §18.8.30 formatter), else General with the host
 * decimal separator.
 */
export interface CategorySource {
  readonly raw: string;
  /** `true` numeric source, `false` text source, absent: unknown provenance. */
  readonly numeric: boolean | null | undefined;
}

function sharedProvenance(chart: ChartModel): boolean | null | undefined {
  if (chart.categories.length > 0) return chart.categoriesNumeric;
  const first = chart.series[0];
  return first?.categories && first.categories.length > 0 ? first.categoriesNumeric : undefined;
}

/** The displayed category text of `index` and its provenance. `series`
 * `null`: the chart's category axis (shared list, generated 1..n when the
 * chart has none). A series: its own cache for the points it covers, else the
 * shared list; `generated` supplies `index + 1` where no source has the
 * point (otherwise empty). */
export function categorySourceAt(
  chart: ChartModel,
  series: ChartSeries | null | undefined,
  index: number,
  generated = false,
): CategorySource {
  const own = series?.categories?.[index];
  if (own !== undefined) return { raw: String(own), numeric: series?.categoriesNumeric };
  if (!series) {
    const raw = chartCategories(chart)[index];
    return { raw: raw === undefined ? '' : String(raw), numeric: sharedProvenance(chart) };
  }
  const shared = chart.categories.length > 0 ? chart.categories : chart.series[0]?.categories ?? [];
  const raw = shared[index];
  if (raw !== undefined) return { raw: String(raw), numeric: sharedProvenance(chart) };
  return { raw: generated ? String(index + 1) : '', numeric: undefined };
}

export interface CategoryCaptionFormat {
  /** The authored axis/series number-format code, if any. */
  readonly code?: string | null;
  /** Per numeric section divisors of `code`. */
  readonly divisors?: readonly number[] | null;
}

/** Format a resolved category source. */
export function formatCategorySource(
  chart: ChartModel,
  source: CategorySource,
  format: CategoryCaptionFormat = {},
): string {
  return formatCategoryLabel(source.raw, format.code, chart.date1904, format.divisors, source.numeric);
}

/** Category-axis tick text of `index`: the shared categories under the axis code. */
export function axisCategoryCaption(chart: ChartModel, index: number): string {
  return formatCategorySource(chart, categorySourceAt(chart, null, index), {
    code: chart.catAxisFormatCode, divisors: chart.catAxisFormatDivisors,
  });
}

/** A point's category text (data labels, legends): the series' own categories
 * where they cover the point, formatted by the series' per-point or cache code
 * and, for a source of known provenance, the category axis code. */
export function pointCategoryCaption(
  chart: ChartModel,
  series: ChartSeries,
  index: number,
  generated = false,
): string {
  const source = categorySourceAt(chart, series, index, generated);
  const own = series.catFormatCodes?.[index] ?? series.catFormatCode ?? null;
  if (own) return formatCategorySource(chart, source, { code: own });
  if (source.numeric === undefined) return formatCategorySource(chart, source);
  return formatCategorySource(chart, source, { code: chart.catAxisFormatCode, divisors: chart.catAxisFormatDivisors });
}

/** `pointCategoryCaption` as a per-index closure, for the painters that take one. */
export function pointCategoryCaptionAt(
  chart: ChartModel,
  series: ChartSeries,
  generated = false,
): (index: number) => string {
  return index => pointCategoryCaption(chart, series, index, generated);
}
