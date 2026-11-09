import { describe, expect, it } from 'vitest';
import { refreshWorksheetCharts } from './chart-series-refresh.js';
import type { Cell, Row, Worksheet } from './types.js';
import type { ChartAnchor } from './types.js';

function cell(row: number, col: number, value: Cell['value']): Cell {
  return { row, col, value };
}

function sheetWithChart(
  name: string,
  cells: Cell[],
  charts: ChartAnchor[],
): Worksheet {
  const byRow = new Map<number, Cell[]>();
  for (const c of cells) {
    const list = byRow.get(c.row) ?? [];
    list.push(c);
    byRow.set(c.row, list);
  }
  const rows: Row[] = [...byRow.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([index, rowCells]) => ({ index, height: null, cells: rowCells }));
  return {
    name,
    rows,
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts,
  } as Worksheet;
}

function barChart(opts: {
  values: (number | null)[];
  categories?: string[] | null;
  valueFormula?: string | null;
  categoryFormula?: string | null;
  chartCategories?: string[];
  chartCategoryFormula?: string | null;
}): ChartAnchor {
  return {
    fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
    toCol: 1, toColOff: 0, toRow: 1, toRowOff: 0,
    chart: {
      chartType: 'clusteredBar',
      title: null,
      categories: opts.chartCategories ?? [],
      categoryFormula: opts.chartCategoryFormula ?? null,
      series: [
        {
          name: 'Revenue',
          color: null,
          values: opts.values,
          valueFormula: opts.valueFormula ?? null,
          categoryFormula: opts.categoryFormula ?? null,
          categories: opts.categories === undefined ? null : opts.categories,
        },
      ],
    },
  } as unknown as ChartAnchor;
}

/** The reported shape: a cache-free bar chart over recalculated cells —
 *  inline categories resolve at parse, values arrive only via recalc. */
function summarySheet(): Worksheet {
  const cells: Cell[] = [
    cell(4, 1, { type: 'text', text: 'Tunis' }),
    cell(5, 1, { type: 'text', text: 'Sfax' }),
    cell(4, 3, { type: 'number', number: 65153.59 }),
    cell(5, 3, { type: 'number', number: 75892.10 }),
  ];
  return sheetWithChart('Summary', cells, [
    barChart({
      values: [null, null],
      categories: null,
      valueFormula: "'Summary'!$C$4:$C$5",
      chartCategories: [],
      chartCategoryFormula: "'Summary'!$A$4:$A$5",
    }),
  ]);
}

describe('refreshWorksheetCharts', () => {
  it('fills cache-free series values and shared categories from recalculated cells', () => {
    const ws = summarySheet();
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [ws] },
      new Set(['summary']),
    );
    expect(patches).toHaveLength(1);
    expect(patches[0]?.chartIndex).toBe(0);
    expect(patches[0]?.categories).toEqual(['Tunis', 'Sfax']);
    expect(patches[0]?.series).toEqual([
      { seriesIndex: 0, values: [65153.59, 75892.1] },
    ]);
    // Live model updated in place for the painter.
    expect(ws.charts[0]?.chart.series[0]?.values).toEqual([65153.59, 75892.1]);
    expect(ws.charts[0]?.chart.categories).toEqual(['Tunis', 'Sfax']);
  });

  it('is a no-op when series already match the live model', () => {
    const ws = summarySheet();
    ws.charts[0]!.chart.series[0]!.values = [65153.59, 75892.1];
    ws.charts[0]!.chart.categories = ['Tunis', 'Sfax'];
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [ws] },
      new Set(['summary']),
    );
    expect(patches).toEqual([]);
  });

  it('leaves authored (formula-free) series untouched', () => {
    const ws = sheetWithChart('S1', [cell(1, 1, { type: 'number', number: 7 })], [
      barChart({ values: [1], categories: ['a'] }),
    ]);
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [ws] },
      new Set(['s1']),
    );
    expect(patches).toEqual([]);
    expect(ws.charts[0]?.chart.series[0]?.values).toEqual([1]);
  });

  it('skips series whose source sheet is missing', () => {
    const ws = sheetWithChart('S1', [], [
      barChart({ values: [null], valueFormula: "'Gone'!$A$1" }),
    ]);
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [ws] },
      new Set(['gone']),
    );
    expect(patches).toEqual([]);
    expect(ws.charts[0]?.chart.series[0]?.values).toEqual([null]);
  });

  it('ignores series on sheets without new values', () => {
    const ws = summarySheet();
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [ws] },
      new Set(['other']),
    );
    expect(patches).toEqual([]);
  });

  it('resolves cross-sheet references against the recalculated sheet', () => {
    const data = sheetWithChart('Data', [cell(1, 1, { type: 'number', number: 42 })], []);
    const ws = sheetWithChart('Chart', [], [
      barChart({ values: [null], valueFormula: 'Data!$A$1' }),
    ]);
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [data, ws] },
      new Set(['data']),
    );
    expect(patches[0]?.series).toEqual([{ seriesIndex: 0, values: [42] }]);
  });

  it('maps gaps to null and numbers to category text', () => {
    const ws = sheetWithChart('S1', [
      cell(1, 1, { type: 'number', number: 5 }),
      cell(3, 1, { type: 'error', error: '#DIV/0!' }),
    ], [
      barChart({
        values: [null, null, null],
        categories: null,
        valueFormula: 'S1!$A$1:$A$3',
        chartCategories: [],
        chartCategoryFormula: 'S1!$A$1:$A$3',
      }),
    ]);
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [ws] },
      new Set(['s1']),
    );
    expect(patches[0]?.series).toEqual([
      { seriesIndex: 0, values: [5, null, null] },
    ]);
    expect(patches[0]?.categories).toEqual(['5', '', '']);
  });

  it('handles quoted sheet names, absolutes, and single cells', () => {
    const ws = sheetWithChart('My Sheet', [cell(2, 2, { type: 'number', number: 9 })], [
      barChart({ values: [null], valueFormula: "'My Sheet'!$B$2" }),
    ]);
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [ws] },
      new Set(['my sheet']),
    );
    expect(patches[0]?.series).toEqual([{ seriesIndex: 0, values: [9] }]);
  });

  it('expands defined names last-match-wins like the parser', () => {
    const ws = sheetWithChart('S1', [cell(1, 1, { type: 'number', number: 3 })], [
      barChart({ values: [null], valueFormula: 'Total' }),
    ]);
    ws.definedNames = [
      { name: 'Total', formula: 'S1!$Z$9' },
      { name: 'total', formula: 'S1!$A$1' },
    ];
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [ws] },
      new Set(['s1']),
    );
    expect(patches[0]?.series).toEqual([{ seriesIndex: 0, values: [3] }]);
  });

  it('rejects out-of-grid and external references', () => {
    const ws = sheetWithChart('S1', [cell(1, 1, { type: 'number', number: 1 })], [
      barChart({ values: [null], valueFormula: 'S1!$A$1048577' }),
      barChart({ values: [null], valueFormula: '[0]Sheet1!$A$1' }),
    ]);
    const patches = refreshWorksheetCharts(
      ws,
      { sheets: [ws] },
      new Set(['s1']),
    );
    expect(patches).toEqual([]);
  });
});
