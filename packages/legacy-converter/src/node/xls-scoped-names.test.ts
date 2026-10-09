// [MS-XLS] 2.4.150 Lbl scopes through the Node session to the XLSX
// conditional-formatting consumer: an unqualified name resolves to the
// worksheet-local definition before the workbook one (Microsoft Support,
// "Names in formulas"), whatever the Lbl record order.
import { expect, it } from 'vitest';
import { testXlsSource } from '../test-sources.js';
import { buildXlsScopedNamesFixture } from '../xls-scoped-names-fixture.js';
import { compileCf, evaluateCf, materializeXlsxWorkbook } from './node-facade.js';

it('keeps both same-name scopes and resolves the current-sheet local first in CF formulas', async () => {
  const { workbookIndex, worksheets } = await materializeXlsxWorkbook(buildXlsScopedNamesFixture(), { modelSources: [testXlsSource()] });
  expect(worksheets.map(sheet => sheet.name)).toEqual(['A', 'B']);
  // Sheet A sees the global and its own local; B sees only the global.
  expect(worksheets.map(sheet => sheet.definedNames)).toEqual([
    [{ name: 'Rate', formula: '5' }, { name: 'Rate', formula: '1' }],
    [{ name: 'Rate', formula: '5' }],
  ]);
  const filled = worksheets.map((sheet) => {
    expect(sheet.conditionalFormats).toEqual([expect.objectContaining({
      rules: [expect.objectContaining({ type: 'expression', formula: 'A1=Rate' })],
    })]);
    const context = compileCf(sheet);
    return sheet.rows.flatMap(row => row.cells).map(cell => [
      cell.col, evaluateCf(cell, cell.row, cell.col, context, workbookIndex.styles.dxfs).fill !== undefined,
    ]);
  });
  // A1 = 1 matches Rate only on A (local 1); B1 = 5 only on B (global 5).
  expect(filled).toEqual([[[1, true], [2, false]], [[1, false], [2, true]]]);
});
