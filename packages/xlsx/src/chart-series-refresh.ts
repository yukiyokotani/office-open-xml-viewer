/**
 * Re-resolve cache-free chart series from live (post-recalculation) cell
 * values.
 *
 * Classic DrawingML series (`<c:val>` / `<c:cat>`, scatter `<c:yVal>` /
 * `<c:xVal>`) whose `<c:f>` reference has no authored cache or literal are
 * resolved once at parse time from cached `<v>` values (see
 * `XlsxChartReferenceResolver`, `collect_number_source` /
 * `collect_string_source`). A workbook whose formula cells carry no cached
 * values therefore parses charts with empty series — and the opt-in
 * recalculation (`recalc.ts`) that later fills the grid never reaches them,
 * because the parser drops the source formula after resolving.
 *
 * The parser now retains that link instead (`ChartSeries.valueFormula` /
 * `categoryFormula`, `ChartModel.categoryFormula`; present only when the
 * source was genuinely host-resolved). After any recalculation, the workbook
 * calls {@link refreshWorksheetCharts} for affected sheets: every formula-
 * backed series whose source range lives on a sheet with new values is read
 * again from the live model and the series (or shared categories) replaced
 * wholesale. The returned patches ride the existing
 * `applyRecalculatedValues` worker message, so the render worker — whose own
 * copy was admitted from the archive with the same empty caches — converges
 * to the same series.
 *
 * Mapping mirrors the parse-time Rust resolver exactly, so a sheet that
 * *does* carry cached values refreshes to identical data:
 * - numbers: finite numbers only, everything else is a gap (`null`);
 * - categories: text verbatim, numbers via `String(number)`, gaps as `""`;
 * - ranges read row-major; missing rows/cells count as empty (never an
 *   error — blank cells are gaps, matching parse-time behavior);
 * - defined names expand first, last match wins case-insensitively (same
 *   rule as the parser's `expand_defined_name`).
 *
 * Out of scope (kept as parse-time snapshots): bubble sizes, per-point
 * visibility provenance, and ChartEx series — none of which the refresh
 * touches, so their values never change here.
 */

import type { CellValue, DefinedName, Worksheet } from './types.js';
import type { RecalcChartPatch } from './worker-protocol.js';

/** 1-based worksheet bounds (ECMA-376 §18.3.1.73 grid limits). */
const MAX_ROW = 1048576;
const MAX_COL = 16384;

interface A1 {
  col: number;
  row: number;
}

interface RangeRef {
  /** Lower-cased sheet name as authored (quotes unescaped). */
  sheet: string;
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/**
 * Expand a defined name the way the parser does: the last
 * case-insensitive match wins, so a sheet-local name shadows its workbook
 * twin. Returns the trimmed body, or the trimmed input when no name matches.
 */
function expandDefinedName(formula: string, names: readonly DefinedName[]): string {
  const trimmed = formula.trim();
  for (let i = names.length - 1; i >= 0; i--) {
    const name = names[i];
    if (name && name.name.toLowerCase() === trimmed.toLowerCase()) {
      return name.formula.trim();
    }
  }
  return trimmed;
}

function parseA1Cell(text: string): A1 | undefined {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(text.trim());
  if (!match) return undefined;
  const letters = (match[1] as string).toUpperCase();
  let col = 0;
  for (const ch of letters) col = col * 26 + (ch.charCodeAt(0) - 64);
  const row = Number(match[2]);
  if (!Number.isSafeInteger(row) || row < 1 || row > MAX_ROW) return undefined;
  if (col < 1 || col > MAX_COL) return undefined;
  return { col, row };
}

/**
 * Split `Sheet!A1:B2` / `'My Sheet'!$A$1` into sheet + rectangle. The
 * `!` inside brackets (`[0]!A1`, external links) is not a resolvable local
 * reference and is rejected, matching the parser which resolves only local
 * worksheet ranges.
 */
function parseRangeFormula(formula: string): RangeRef | undefined {
  const text = formula.trim();
  if (text.includes('[')) return undefined;
  let sheet: string;
  let rangeText: string;
  if (text.startsWith("'")) {
    const end = text.indexOf("'", 1);
    if (end < 0) return undefined;
    sheet = text.slice(1, end).replace(/''/g, "'");
    const rest = text.slice(end + 1);
    if (!rest.startsWith('!')) return undefined;
    rangeText = rest.slice(1);
  } else {
    const bang = text.indexOf('!');
    if (bang < 0) return undefined;
    sheet = text.slice(0, bang);
    rangeText = text.slice(bang + 1);
  }
  if (sheet.length === 0 || rangeText.length === 0) return undefined;
  const parts = rangeText.split(':');
  if (parts.length > 2) return undefined;
  const a = parseA1Cell(parts[0] as string);
  if (!a) return undefined;
  const b = parts.length === 2 ? parseA1Cell(parts[1] as string) : a;
  if (!b) return undefined;
  return {
    sheet: sheet.toLowerCase(),
    top: Math.min(a.row, b.row),
    left: Math.min(a.col, b.col),
    bottom: Math.max(a.row, b.row),
    right: Math.max(a.col, b.col),
  };
}

type CellLookup = (row: number, col: number) => CellValue | undefined;

function makeCellLookup(ws: Worksheet): CellLookup {
  const byRow = new Map<number, Map<number, CellValue>>();
  for (const row of ws.rows) {
    let cols = byRow.get(row.index);
    if (!cols) {
      cols = new Map<number, CellValue>();
      byRow.set(row.index, cols);
    }
    for (const cell of row.cells) cols.set(cell.col, cell.value);
  }
  return (row, col) => byRow.get(row)?.get(col);
}

function readNumber(value: CellValue | undefined): number | null {
  if (!value || value.type !== 'number') return null;
  return Number.isFinite(value.number) ? value.number : null;
}

function readCategory(value: CellValue | undefined): string {
  if (!value) return '';
  if (value.type === 'text') return value.text;
  // Mirrors the parse-time resolver: numbers stringify, everything else
  // (bool / error / empty / unresolved shared refs) is a blank slot.
  if (value.type === 'number') return String(value.number);
  return '';
}

function sameValues(a: ReadonlyArray<number | null>, b: ReadonlyArray<number | null>): boolean {
  if (a.length !== b.length) return false;
  return a.every((v, i) => {
    const w = b[i];
    return v === w || (v !== null && w !== null && Object.is(v, w));
  });
}

function sameStrings(a: ReadonlyArray<string>, b: ReadonlyArray<string>): boolean {
  return a.length === b.length && a.every((v, i) => v === b[i]);
}

export interface ChartRefreshScope {
  /** Live post-recalc worksheet models. */
  sheets: ReadonlyArray<Worksheet>;
}

/**
 * Re-resolve every formula-backed chart series on `ws` whose source range
 * lives on one of `recalculatedSheets` (lower-cased names). Defined names
 * expand against the chart's own sheet, mirroring the parser (whose chart
 * resolver receives the drawing sheet's names). Returns one patch per chart
 * that actually changed; callers apply the patches to the same objects they
 * were read from.
 */
export function refreshWorksheetCharts(
  ws: Worksheet,
  scope: ChartRefreshScope,
  recalculatedSheets: ReadonlySet<string>,
): RecalcChartPatch[] {
  if (!ws.charts || ws.charts.length === 0) return [];
  const patches: RecalcChartPatch[] = [];
  const lookups = new Map<string, CellLookup>();
  const lookupOf = (sheetName: string): CellLookup | undefined => {
    const key = sheetName.toLowerCase();
    const hit = lookups.get(key);
    if (hit) return hit;
    const target = scope.sheets.find((s) => s.name.toLowerCase() === key);
    if (!target) return undefined;
    const lookup = makeCellLookup(target);
    lookups.set(key, lookup);
    return lookup;
  };
  ws.charts.forEach((anchor, chartIndex) => {
    const chart = anchor.chart;
    if (!chart || !chart.series) return;
    const names = ws.definedNames ?? [];
    const readNumbers = (formula: string): (number | null)[] | undefined => {
      const ref = parseRangeFormula(expandDefinedName(formula, names));
      if (!ref || !recalculatedSheets.has(ref.sheet)) return undefined;
      const lookup = lookupOf(ref.sheet);
      if (!lookup) return undefined;
      const out: (number | null)[] = [];
      for (let row = ref.top; row <= ref.bottom; row++) {
        for (let col = ref.left; col <= ref.right; col++) {
          out.push(readNumber(lookup(row, col)));
        }
      }
      return out;
    };
    const readStrings = (formula: string): string[] | undefined => {
      const ref = parseRangeFormula(expandDefinedName(formula, names));
      if (!ref || !recalculatedSheets.has(ref.sheet)) return undefined;
      const lookup = lookupOf(ref.sheet);
      if (!lookup) return undefined;
      const out: string[] = [];
      for (let row = ref.top; row <= ref.bottom; row++) {
        for (let col = ref.left; col <= ref.right; col++) {
          out.push(readCategory(lookup(row, col)));
        }
      }
      return out;
    };
    const seriesPatches: RecalcChartPatch['series'] = [];
    chart.series.forEach((series, seriesIndex) => {
      const patch: { seriesIndex: number; values?: (number | null)[]; categories?: string[] } = {
        seriesIndex,
      };
      let changed = false;
      if (series.valueFormula) {
        const values = readNumbers(series.valueFormula);
        if (values && !sameValues(values, series.values)) {
          series.values = values;
          patch.values = values;
          changed = true;
        }
      }
      if (series.categoryFormula) {
        const categories = readStrings(series.categoryFormula);
        if (categories && !(series.categories && sameStrings(categories, series.categories))) {
          series.categories = categories;
          patch.categories = categories;
          changed = true;
        }
      }
      // Series without their own formula fall back to the chart-level
      // categories, refreshed once below — nothing per-series to do here.
      if (changed) seriesPatches.push(patch);
    });
    let chartCategories: string[] | undefined;
    if (
      chart.categoryFormula &&
      chart.series.some((s) => s.categoryFormula == null && s.categories == null)
    ) {
      const resolved = readStrings(chart.categoryFormula);
      if (resolved && !sameStrings(resolved, chart.categories)) {
        chart.categories = resolved;
        chartCategories = resolved;
      }
    }
    if (seriesPatches.length > 0 || chartCategories) {
      patches.push({
        chartIndex,
        ...(chartCategories ? { categories: chartCategories } : {}),
        series: seriesPatches,
      });
    }
  });
  return patches;
}
