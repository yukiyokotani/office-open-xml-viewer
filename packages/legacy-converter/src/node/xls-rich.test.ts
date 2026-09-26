// BIFF8 rich shared strings (MS-XLS 2.5.293) through the direct XLS reader into
// XLSX rich-text cell values. Run validation and the end-sentinel rule are
// unit-tested in xls/rich.rs.
import { expect, it } from 'vitest';
import { testXlsSource } from '../test-sources.js';
import { openXlsxWorkbook } from './node-facade.js';
import { buildXlsRichFixture } from './xls-rich-fixture.js';

it('keeps separate BIFF run fonts as XLSX rich-text runs', async () => {
  const workbook = await openXlsxWorkbook(buildXlsRichFixture(), { modelSources: [testXlsSource()] });
  try {
    const cells = [];
    for await (const chunk of workbook.worksheetRows(0)) {
      if (chunk.kind === 'rows') for (const row of chunk.rows) cells.push(...row.cells);
    }
    expect(cells.find(cell => cell.row === 2 && cell.col === 2)?.value).toMatchObject({
      type: 'text', text: 'base RED normal', runs: [
        { text: 'base ' },
        { text: 'RED ', font: { name: 'Arial', size: 24, color: '#FF0000', bold: false } },
        { text: 'normal', font: { name: 'Arial', size: 11, bold: false } },
      ],
    });
  } finally { await workbook.close(); }
});
