import { formatChartValWithCode } from './chart-number-format.js';

export interface EffectiveDataLabelTextOptions {
  customText?: string | null;
  showCategory?: boolean;
  showSeries?: boolean;
  showValue?: boolean;
  showPercent?: boolean;
  showBubbleSize?: boolean;
  category?: string;
  seriesName?: string;
  sourceValue?: number;
  /** Effective divisor of the value axis associated with this label. Office
   * applies `<c:dispUnits>` to generated `showVal` text while leaving the
   * source value and value-to-pixel geometry unchanged. */
  valueDivisor?: number | null;
  percentRatio?: number;
  bubbleSize?: number;
  formatCode?: string | null;
  /** Section divisors of `formatCode` (an OpenDocument display factor), for
   *  the value and bubble-size text; the percentage has its own format. */
  formatDivisors?: readonly number[] | null;
  percentFormatCode?: string | null;
  date1904?: boolean;
  separator?: string | null;
  defaultSeparator?: string;
}

/** Compose authored data-label content independently from family geometry. */
export function effectiveDataLabelText(options: EffectiveDataLabelTextOptions): string {
  if (options.customText) return options.customText;
  const parts: string[] = [];
  if (options.showCategory && options.category) parts.push(options.category);
  if (options.showSeries && options.seriesName) parts.push(options.seriesName);
  if (options.showValue && options.sourceValue != null) {
    const divisor = options.valueDivisor != null
      && Number.isFinite(options.valueDivisor)
      && options.valueDivisor > 0
      ? options.valueDivisor
      : 1;
    parts.push(formatChartValWithCode(
      options.sourceValue / divisor, options.formatCode ?? null, options.date1904,
      options.formatDivisors,
    ));
  }
  if (options.showPercent && options.percentRatio != null) {
    parts.push(formatChartValWithCode(
      options.percentRatio,
      options.percentFormatCode ?? options.formatCode ?? '0%',
      options.date1904,
    ));
  }
  if (options.showBubbleSize && options.bubbleSize != null) {
    parts.push(formatChartValWithCode(
      options.bubbleSize, options.formatCode ?? null, options.date1904,
      options.formatDivisors,
    ));
  }
  return parts.filter(part => part !== '').join(
    options.separator ?? options.defaultSeparator ?? ' ',
  );
}
