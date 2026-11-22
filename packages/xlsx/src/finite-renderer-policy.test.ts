import {
  DEFAULT_XLSX_WORKSHEET_POLICY,
  normalizeXlsxWorksheetPolicy,
  type NormalizedXlsxWorksheetPolicy,
} from './internal/worksheet-policy.js';
import { describe, expect, it } from 'vitest';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { getSheetRenderCache, inheritSheetRenderCache } from './renderer.js';
import { bindWorksheetPolicy, getWorksheetPolicy } from './worksheet-policy-context.js';
import { createSheetViewModel } from './internal/sheet-viewer-runtime.js';
import { createSizeOverriddenWorksheet } from './worker-protocol.js';
import type { Cell, Styles, TableInfo, Worksheet } from './types.js';

const DEFAULT_CAP = 250_000;
const RAISED_CAP = 300_000;
const ROWS = 12_550;
const COLS = 20;
const TOTAL_CELLS = ROWS * COLS; // 251000

type Row = Worksheet['rows'][number];
type MergeRange = Worksheet['mergeCells'][number];

function raisedPolicy(): NormalizedXlsxWorksheetPolicy {
  return normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxCells: RAISED_CAP } });
}

function baseWorksheet(overrides: Partial<Worksheet> = {}): Worksheet {
  return {
    name: 'FinitePolicy',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    ...overrides,
  };
}

function numericCell(row: number, col: number): Cell {
  return {
    row,
    col,
    value: { type: 'number', number: row * 100 + col },
    styleIndex: 0,
  };
}

function buildLargeWorksheet(): Worksheet {
  const rows: Row[] = new Array<Row>(ROWS);
  for (let r = 1; r <= ROWS; r++) {
    const cells: Cell[] = new Array<Cell>(COLS);
    for (let c = 1; c <= COLS; c++) cells[c - 1] = numericCell(r, c);
    rows[r - 1] = { index: r, height: null, cells };
  }
  return baseWorksheet({ rows });
}

let largeFixture: Worksheet | null = null;
/** One 251000-cell fixture per file; never rebuilt or deep-copied. */
function largeWorksheet(): Worksheet {
  largeFixture ??= buildLargeWorksheet();
  return largeFixture;
}

function lastSourceCell(ws: Worksheet): Cell {
  const lastRow = ws.rows[ws.rows.length - 1] as unknown as { cells: Cell[] };
  return lastRow.cells[lastRow.cells.length - 1];
}

function cachedCell(ws: Worksheet, row: number, col: number): unknown {
  return getSheetRenderCache(ws).cellMap.get(`${row}:${col}`);
}

function expectConfigurableLimit(run: () => unknown, limit: number, observed?: number): void {
  let caught: unknown = null;
  try {
    run();
  } catch (error) {
    caught = error;
  }
  expect(caught).toBeInstanceOf(OoxmlResourceLimitError);
  const typed = caught as OoxmlResourceLimitError;
  expect(typed.code).toBe('ooxml-resource-limit');
  expect(typed.details).toMatchObject({
    stage: 'rendering',
    violation: { format: 'xlsx', metric: 'entry-count', limit, configurable: true },
  });
  const violation = (typed.details as { violation: { observed: number } }).violation;
  if (observed === undefined) {
    expect(violation.observed).toBeGreaterThan(limit);
  } else {
    expect(violation.observed).toBe(observed);
  }
}

describe('renderer coordinate index honours the bound worksheet policy', () => {
  it('admits 251000 cells only under a raised bound policy, across projections, and rejects stale cache after rebinding', () => {
    const source = largeWorksheet();
    const rowsRef = source.rows;
    const rowCount = source.rows.length;
    const keysBefore = Object.keys(source).sort();
    const lastCellBefore = lastSourceCell(source);
    const lastValueBefore = JSON.stringify(lastCellBefore);
    const rowHeightKeysBefore = Object.keys(source.rowHeights).length;

    // Default policy: 251000 entries exceed the default 250000 cap.
    expect(getWorksheetPolicy(source)).toBe(DEFAULT_XLSX_WORKSHEET_POLICY);
    expectConfigurableLimit(() => getSheetRenderCache(source), DEFAULT_CAP);

    // Raised policy: construction succeeds and the last cell matches input.
    const raised = raisedPolicy();
    expect(raised.maxCoordinateIndexEntries).toBe(RAISED_CAP);
    bindWorksheetPolicy(source, raised);
    const cache = getSheetRenderCache(source);
    expect(cache.cellMap.size).toBe(TOTAL_CELLS);
    expect(cachedCell(source, ROWS, COLS)).toBe(lastCellBefore);

    // Main-thread viewer projection inherits the policy.
    const view = createSheetViewModel(source);
    expect(getWorksheetPolicy(view)).toBe(raised);
    expect(getSheetRenderCache(view).cellMap.size).toBe(TOTAL_CELLS);
    expect(cachedCell(view, ROWS, COLS)).toBe(lastCellBefore);

    // Worker size-override projection inherits the policy.
    const overridden = createSizeOverriddenWorksheet(source, { rows: { 1: 30 } });
    expect(overridden).not.toBe(source);
    expect(getWorksheetPolicy(overridden)).toBe(raised);
    expect(getSheetRenderCache(overridden).cellMap.size).toBe(TOTAL_CELLS);
    expect(cachedCell(overridden, ROWS, COLS)).toBe(lastCellBefore);

    // Inherited (auto-height) cache reuse preserves the policy.
    const reused = createSizeOverriddenWorksheet(source, { rows: { 2: 40 } });
    inheritSheetRenderCache(source, reused, { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles);
    expect(getWorksheetPolicy(reused)).toBe(raised);
    expect(getSheetRenderCache(reused).cellMap.size).toBe(TOTAL_CELLS);
    expect(cachedCell(reused, ROWS, COLS)).toBe(lastCellBefore);

    // Rebinding to default must invalidate the higher-allowance cache.
    bindWorksheetPolicy(source, DEFAULT_XLSX_WORKSHEET_POLICY);
    expectConfigurableLimit(() => getSheetRenderCache(source), DEFAULT_CAP);

    // Source unmodified; no public policy stamp on the model.
    expect(source.rows).toBe(rowsRef);
    expect(source.rows.length).toBe(rowCount);
    expect(Object.keys(source).sort()).toEqual(keysBefore);
    expect(lastSourceCell(source)).toBe(lastCellBefore);
    expect(JSON.stringify(lastSourceCell(source))).toBe(lastValueBefore);
    expect(Object.keys(source.rowHeights).length).toBe(rowHeightKeysBefore);
    expect(Object.keys(view).sort()).toEqual(keysBefore);
  });

  it('applies the bound cap to merge skip and table style indexes on a tiny graph', () => {
    const twoCellRows = (): Row[] => [
      { index: 1, cells: [numericCell(1, 1), numericCell(1, 2)] } as unknown as Row,
    ];
    // 2 columns x 125501 rows = 251002 covered entries.
    const merge = { top: 1, left: 1, bottom: 125_501, right: 2 } as unknown as MergeRange;

    const mergedDefault = baseWorksheet({ rows: twoCellRows(), mergeCells: [merge] });
    const mergeKeysBefore = Object.keys(mergedDefault).sort();
    expectConfigurableLimit(() => getSheetRenderCache(mergedDefault), DEFAULT_CAP);

    const mergedRaised = baseWorksheet({ rows: twoCellRows(), mergeCells: [merge] });
    bindWorksheetPolicy(mergedRaised, raisedPolicy());
    const mergedCache = getSheetRenderCache(mergedRaised);
    expect(mergedCache.cellMap.size).toBe(2);
    expect(Object.keys(mergedDefault).sort()).toEqual(mergeKeysBefore);
    expect(mergedDefault.mergeCells[0]).toBe(merge);

    // Styled table range of derived cap + 1 refuses even with only 2 cells.
    const range = { top: 1, left: 1, bottom: RAISED_CAP + 1, right: 1 } as unknown as TableInfo['range'];
    const tableInfo: TableInfo = {
      range,
      styleName: 'TableStyleLight1',
      headerRowCount: 1,
      totalsRowCount: 0,
      showRowStripes: true,
      showColumnStripes: false,
      showFirstColumn: false,
      showLastColumn: false,
      accentColor: '#4472C4',
      columns: [],
      styleElements: [{ kind: 'wholeTable', size: 1, dxf: { font: null, fill: null, border: null } }],
    };
    const tabled = baseWorksheet({ rows: twoCellRows(), tables: [tableInfo] } as Partial<Worksheet>);
    bindWorksheetPolicy(tabled, raisedPolicy());
    expectConfigurableLimit(() => getSheetRenderCache(tabled), RAISED_CAP, RAISED_CAP + 1);
  });
});
