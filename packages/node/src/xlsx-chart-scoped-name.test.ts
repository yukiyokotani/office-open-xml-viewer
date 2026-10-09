import { describe, expect, it } from 'vitest';
import { materializeXlsxWorkbook } from './xlsx.ts';
import { scopedNameChartXlsx } from './test-xlsx-scoped-name-chart.ts';

describe('Node XLSX chart ingress of sheet-scoped defined names (#1547)', () => {
  it('feeds each sheet chart the unqualified name visible from that sheet', async () => {
    const { workbookIndex, worksheets } = await materializeXlsxWorkbook(scopedNameChartXlsx());
    expect(workbookIndex.workbook.sheets.map(({ name }) => name)).toEqual(['Data', 'Other', 'Plain']);
    const [data, other, plain] = worksheets;

    // Data's local `Revenue` (authored after the global) shadows the global twin; the
    // Other-sheet local is out of scope even though it is authored last.
    expect(data.charts.map(({ chart }) => chart.series.map(({ values }) => values)))
      .toEqual([[[11, 12, 13]]]);
    // Global-only control: Plain has no local twin.
    expect(plain.charts.map(({ chart }) => chart.series.map(({ values }) => values)))
      .toEqual([[[21, 22, 23]]]);

    // The shadowed global stays in the sheet inventory, workbook scope first,
    // so last-match consumers (CF formulas, internal hyperlinks) agree.
    expect(data.definedNames).toEqual([
      { name: 'Revenue', formula: 'Data!$C$2:$C$4' },
      { name: 'Revenue', formula: 'Data!$B$2:$B$4' },
    ]);
    expect(other.definedNames).toEqual([
      { name: 'Revenue', formula: 'Data!$C$2:$C$4' },
      { name: 'Revenue', formula: 'Data!$D$2:$D$4' },
    ]);
    expect(plain.definedNames).toEqual([{ name: 'Revenue', formula: 'Data!$C$2:$C$4' }]);
  });
});
