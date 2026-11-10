import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';
import { xlsxCellDisplayTexts, type XlsxWorkbook } from './workbook.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX2 — destroy() must drop the find state (twin of the DocxViewer/PptxViewer
 * tests). findNext()/findPrev() on a torn-down viewer must return null, not a
 * stale match (which would even try to sheet-switch/scroll a dead viewer).
 */
function makeFakeWorkbook(): XlsxWorkbook {
  const ws = {
    rows: [{ index: 1, cells: [{ row: 1, col: 1, value: { type: 'text', text: 'hello world' } }] }],
  };
  return {
    sheetCount: 1,
    sheetNames: ['Sheet1'],
    getWorksheet: () => Promise.resolve(ws),
    acquireWorksheetLease: async () => ({ worksheet: ws, release: () => undefined }),
    cellText: () => 'hello world',
    [xlsxCellDisplayTexts]: () => () => ['hello world'],
    destroy: vi.fn(),
    isHidden: () => false,
  } as unknown as XlsxWorkbook;
}

describe('XlsxViewer.destroy() — find state invalidation', () => {
  it('findNext()/findPrev() return null after destroy() (no stale matches)', async () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement);
    (v as unknown as { wb: XlsxWorkbook }).wb = makeFakeWorkbook();

    // A live find with a real active match…
    const matches = await v.findText('hello');
    expect(matches).toHaveLength(1);
    expect(matches[0].location.ref).toBe('A1');
    const active = await v.findNext();
    expect(active).not.toBeNull();
    expect(active?.matchIndex).toBe(0);

    // …must not survive teardown.
    v.destroy();
    expect(await v.findNext()).toBeNull();
    expect(await v.findPrev()).toBeNull();
  });
});
