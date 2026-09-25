import type {
  ChartDataLabelOverride,
  ChartDisplayUnits,
  ChartLabelBox,
  ChartModel,
  ChartSeries,
} from '../types/chart.js';
import { effectiveDataLabelText } from './data-label-content.js';
import {
  dataLabelIsDeleted,
  effectiveDataLabelTextStyle,
  type DataLabelTextStyle,
} from './data-label-style.js';
import { mergeChartLabelBoxes } from './label-box.js';

export interface ResolvedChartExLabel {
  text: string;
  showLegendKey: boolean;
  position?: string;
  fontColor?: string;
  fontSizeHpt?: number;
  fontBold?: boolean;
  fontFace?: string;
  manualLayout?: ChartDataLabelOverride['manualLayout'];
  labelBox?: ChartLabelBox;
  richRuns?: ChartDataLabelOverride['richRuns'];
  textStyle: DataLabelTextStyle;
}

function displayUnitDivisor(units: ChartDisplayUnits | null | undefined): number {
  const divisor = units?.divisor;
  return divisor != null && Number.isFinite(divisor) && divisor > 0 ? divisor : 1;
}

/** Shared CT_DataLabels + indexed CT_DataLabel resolution for every ChartEx
 * painter and its resource preflight. Keeping one resolver prevents hierarchy
 * nodes from being counted, prefetched, and painted under different rules. */
export function resolveChartExLabel(
  chart: ChartModel,
  series: ChartSeries | null | undefined,
  index: number,
  category: string,
  value: number,
  defaults: {
    visible: boolean;
    showVal: boolean;
    showCatName: boolean;
    showSerName?: boolean;
    showPercent?: boolean;
  },
  overrideLookup: ReadonlyMap<number, NonNullable<ChartSeries['dataLabelOverrides']>[number]>,
  valueOption?: boolean | number,
  valueDisplayUnits?: ChartDisplayUnits | null,
): ResolvedChartExLabel | null {
  if (!series) return null;
  const definition = series.seriesDataLabels;
  const override = overrideLookup.get(index);
  if (dataLabelIsDeleted(definition, override)) return null;
  if (!definition && !override && !defaults.visible) return null;
  const suppressValue = typeof valueOption === 'boolean' ? valueOption : false;
  const percentRatio = typeof valueOption === 'number' ? valueOption : undefined;
  const showVal = !suppressValue
    && (override?.showVal ?? definition?.showVal ?? defaults.showVal);
  const showCatName = override?.showCatName ?? definition?.showCatName ?? defaults.showCatName;
  const showSerName = override?.showSerName
    ?? definition?.showSerName
    ?? defaults.showSerName
    ?? false;
  const showPercent = override?.showPercent
    ?? definition?.showPercent
    ?? defaults.showPercent
    ?? false;
  const showLegendKey = override?.showLegendKey ?? definition?.showLegendKey ?? false;
  const authoredFormatCode = override?.formatCode
    ?? definition?.formatCode
    ?? chart.dataLabelFormatCode
    ?? null;
  const text = effectiveDataLabelText({
    customText: override?.text,
    showCategory: showCatName,
    showSeries: showSerName,
    showValue: showVal,
    showPercent,
    category,
    seriesName: series.name,
    sourceValue: value,
    valueDivisor: displayUnitDivisor(valueDisplayUnits),
    percentRatio,
    formatCode: authoredFormatCode ?? series.valFormatCode ?? null,
    percentFormatCode: authoredFormatCode ?? '0%',
    date1904: chart.date1904,
    separator: override?.separator ?? definition?.separator,
  });
  if (!text && !showLegendKey) return null;
  return {
    text,
    showLegendKey,
    position: override?.position ?? definition?.position,
    fontColor: override?.fontColor ?? definition?.fontColor,
    fontSizeHpt: override?.fontSizeHpt ?? definition?.fontSizeHpt,
    fontBold: override?.fontBold ?? definition?.fontBold,
    fontFace: override?.fontFace ?? definition?.fontFace,
    manualLayout: override?.manualLayout,
    labelBox: mergeChartLabelBoxes(override?.labelBox, definition?.labelBox),
    richRuns: override?.text ? override.richRuns : undefined,
    textStyle: effectiveDataLabelTextStyle(override, definition),
  };
}
