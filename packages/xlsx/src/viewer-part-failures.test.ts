import { describe, expect, it, vi } from 'vitest';
import { XlsxSheetViewer } from './viewer.js';
import type { XlsxWorkbook } from './workbook.js';
import type { Worksheet } from './types.js';
import { xlsxPartFailureLog } from './part-failures.js';
import { installDom, makeContainer, makeEl } from './viewer-destroy-test-dom.js';

/** #1974 — viewer delivery of XLSX part failures: workbook-level replay at
 *  commit, lazy sheet failures as they are admitted, nothing after destroy. */
function worksheet(name: string): Worksheet {
  return {
    name, rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
  };
}

describe('XlsxSheetViewer part failures (#1974)', () => {
  it('replays the workbook failure and delivers sheet failures as they are admitted', async () => {
    installDom();
    const source = worksheet('Sheet1');
    const workbook = {
      mode: 'main',
      sheetCount: 2,
      sheetNames: ['Sheet1', 'Sheet2'],
      tabColors: {} as Record<number, string>,
      getWorksheet: vi.fn().mockResolvedValue(source),
      acquireWorksheetLease: async () => ({ worksheet: source, release: () => undefined }),
      isHidden: () => false,
      destroy: vi.fn(),
      get partFailures() { return xlsxPartFailureLog(this).list; },
    } as unknown as XlsxWorkbook;
    // The viewer reads the engine's log through the identity registry.
    const log = xlsxPartFailureLog(workbook);
    log.add({ kind: 'workbook', message: 'xl/sharedStrings.xml: bad' });
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const seen: string[] = [];
    const viewer = XlsxSheetViewer.fromWorkbook(canvas as unknown as HTMLCanvasElement, workbook, {
      onPartFailure: (f) => seen.push(f.kind === 'workbook' ? 'workbook' : `sheet:${f.sheetIndex}`),
    });
    await vi.waitFor(() => expect(seen).toEqual(['workbook']));
    log.add({ kind: 'sheet', sheetIndex: 1, sheetName: 'Sheet2', message: 'xl/worksheets/sheet2.xml: bad' });
    expect(seen).toEqual(['workbook', 'sheet:1']);
    expect(viewer.partFailures.map((f) => f.kind)).toEqual(['workbook', 'sheet']);
    viewer.destroy();
    log.add({ kind: 'sheet', sheetIndex: 0, sheetName: 'Sheet1', message: 'x' });
    expect(seen).toEqual(['workbook', 'sheet:1']);
    expect(viewer.partFailures).toEqual([]);
  });
});
