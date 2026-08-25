import { excelSerialToUtcDate, utcDateToExcelSerial } from '../excel-date.js';
import { MAX_AXIS_TICKS } from './axis-scale.js';

export type ChartDateTimeUnit = 'days' | 'months' | 'years';

export interface DateCategoryAxisPlan {
  positions: number[];
  categoryBandFractions: number[];
  majorTicks: Array<{ serial: number; fraction: number }>;
  minorTicks: Array<{ serial: number; fraction: number }>;
}

export interface DateCategoryAxisOptions {
  categories: readonly string[];
  date1904?: boolean;
  baseTimeUnit?: string | null;
  majorTimeUnit?: string | null;
  majorUnit?: number | null;
  minorTimeUnit?: string | null;
  minorUnit?: number | null;
  explicitMin?: number | null;
  explicitMax?: number | null;
  crossBetween?: boolean;
  reversed?: boolean;
}

function timeUnit(value: string | null | undefined): ChartDateTimeUnit | null {
  return value === 'days' || value === 'months' || value === 'years' ? value : null;
}

function floorToUnit(date: Date, unit: ChartDateTimeUnit): Date {
  switch (unit) {
    case 'years': return new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
    case 'months': return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
    case 'days': return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  }
}

function addUnits(date: Date, count: number, unit: ChartDateTimeUnit): Date {
  if (unit === 'days') return new Date(date.getTime() + count * 86_400_000);
  const result = new Date(date.getTime());
  if (unit === 'months') result.setUTCMonth(result.getUTCMonth() + count);
  else result.setUTCFullYear(result.getUTCFullYear() + count);
  return result;
}

function unitCoordinate(date: Date, unit: ChartDateTimeUnit): number {
  switch (unit) {
    case 'years': return date.getUTCFullYear();
    case 'months': return date.getUTCFullYear() * 12 + date.getUTCMonth();
    case 'days': return Math.floor(date.getTime() / 86_400_000);
  }
}

/**
 * Plan a classic DrawingML `<c:dateAx>` from its authored serial categories.
 * CT_DateAx supplies a `baseTimeUnit`: values within one base calendar unit
 * occupy the same category slot (for example January 1 and January 31 share
 * the January slot when the base unit is months). Major/minor ticks advance on
 * authored UTC calendar boundaries rather than on elapsed-day fractions.
 *
 * OOXML does not define the application's automatic major/minor interval when
 * the corresponding unit is omitted. Category coordinates are nevertheless
 * always calendar coordinates; tick arrays remain empty until a separately
 * validated application-compatible automatic interval is available.
 */
export function planDateCategoryAxis(
  options: DateCategoryAxisOptions,
): DateCategoryAxisPlan | null {
  const serials = options.categories.map(value => Number(value));
  if (serials.length === 0 || serials.some(value => !Number.isFinite(value))) return null;

  const date1904 = options.date1904 === true;
  const baseUnit = timeUnit(options.baseTimeUnit) ?? 'days';
  const majorUnit = timeUnit(options.majorTimeUnit) ?? baseUnit;
  const minorUnit = timeUnit(options.minorTimeUnit) ?? baseUnit;
  const authoredStep = (
    value: number | null | undefined,
    unit: ChartDateTimeUnit,
  ): number | null => {
    if (value == null || !(value > 0) || !Number.isFinite(value)) return null;
    if (unit === 'days') return value;
    // ST_AxisUnit permits positive doubles, but neither ECMA-376 nor MS-OE376
    // defines fractional calendar-month/year arithmetic. Preserve integral
    // authored calendar units and fail closed for fractional ones instead of
    // extrapolating a cadence from a finite set of application observations.
    return value >= 1 && Number.isInteger(value) ? value : null;
  };
  const majorStep = authoredStep(options.majorUnit, majorUnit);
  const minorStep = authoredStep(options.minorUnit, minorUnit);

  const coordinate = (serial: number): number =>
    unitCoordinate(excelSerialToUtcDate(serial, date1904), baseUnit);
  const coordinates = new Array<number>(serials.length);
  let dataMin = Infinity;
  let dataMax = -Infinity;
  let serialMin = Infinity;
  let serialMax = -Infinity;
  for (let index = 0; index < serials.length; index++) {
    const serial = serials[index];
    const value = coordinate(serial);
    coordinates[index] = value;
    dataMin = Math.min(dataMin, value);
    dataMax = Math.max(dataMax, value);
    serialMin = Math.min(serialMin, serial);
    serialMax = Math.max(serialMax, serial);
  }
  const authoredMin = options.explicitMin;
  const authoredMax = options.explicitMax;
  const authoredMinCoordinate = authoredMin != null && Number.isFinite(authoredMin)
    ? coordinate(authoredMin)
    : null;
  const authoredMaxCoordinate = authoredMax != null && Number.isFinite(authoredMax)
    ? coordinate(authoredMax)
    : null;
  const crossBetween = options.crossBetween !== false;
  // A date axis groups values into base-unit buckets. With `crossBetween`, the
  // authored scaling bounds and calendar ticks address bucket boundaries while
  // the category mark is painted at the bucket centre. This distinction is
  // observable when an authored minimum equals the first visible category: a
  // column starts inside the plot instead of being centred on the value axis.
  let domainMin = authoredMinCoordinate ?? dataMin;
  let domainMax = (authoredMaxCoordinate ?? dataMax) + (crossBetween ? 1 : 0);
  if (!(domainMax > domainMin)) {
    domainMin -= 0.5;
    domainMax += 0.5;
  }
  const range = domainMax - domainMin;
  const rawAxisFraction = (serial: number): number => (coordinate(serial) - domainMin) / range;
  const axisFraction = options.reversed
    ? (serial: number): number => 1 - rawAxisFraction(serial)
    : rawAxisFraction;
  const rawCategoryFraction = (serial: number): number =>
    (coordinate(serial) + (crossBetween ? 0.5 : 0) - domainMin) / range;
  const categoryFraction = options.reversed
    ? (serial: number): number => 1 - rawCategoryFraction(serial)
    : rawCategoryFraction;

  const positions = serials.map(categoryFraction);
  const categoryBandFractions = serials.map(() => 1 / range);

  const tickMinSerial = authoredMin != null && Number.isFinite(authoredMin)
    ? authoredMin
    : serialMin;
  const tickMaxSerial = authoredMax != null && Number.isFinite(authoredMax)
    ? authoredMax
    : serialMax;
  const planTicks = (
    unit: ChartDateTimeUnit,
    step: number | null,
  ): DateCategoryAxisPlan['majorTicks'] => {
    if (step == null) return [];
    let tickDate = floorToUnit(excelSerialToUtcDate(tickMinSerial, date1904), unit);
    let tickSerial = utcDateToExcelSerial(tickDate, date1904);
    for (let count = 0; tickSerial < tickMinSerial && count < MAX_AXIS_TICKS; count++) {
      const nextDate = addUnits(tickDate, step, unit);
      const nextSerial = utcDateToExcelSerial(nextDate, date1904);
      if (!(nextSerial > tickSerial)) return [];
      tickDate = nextDate;
      tickSerial = nextSerial;
    }
    if (tickSerial < tickMinSerial) return [];
    const ticks: DateCategoryAxisPlan['majorTicks'] = [];
    while (tickSerial <= tickMaxSerial) {
      // An authored interval remains authored: when it exceeds the shared
      // tick-layer ceiling, omit the whole layer instead of painting a prefix
      // or inventing a coarser interval.
      if (ticks.length === MAX_AXIS_TICKS) return [];
      ticks.push({ serial: tickSerial, fraction: axisFraction(tickSerial) });
      const nextDate = addUnits(tickDate, step, unit);
      const nextSerial = utcDateToExcelSerial(nextDate, date1904);
      if (!(nextSerial > tickSerial)) break;
      tickDate = nextDate;
      tickSerial = nextSerial;
    }
    return ticks;
  };
  const majorTicks = planTicks(majorUnit, majorStep);
  const majorCoordinates = new Set(majorTicks.map(tick => coordinate(tick.serial)));
  const minorTicks = minorStep == null
    ? []
    : planTicks(minorUnit, minorStep)
      .filter(tick => !majorCoordinates.has(coordinate(tick.serial)));

  return { positions, categoryBandFractions, majorTicks, minorTicks };
}
