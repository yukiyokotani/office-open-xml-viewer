import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxSheetViewer, XlsxViewer } from './viewer.js';
import { installDom, makeContainer, makeEl } from './viewer-destroy-test-dom.js';
import { xlsxCellDisplayTexts, type XlsxWorkbook } from './workbook.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * Text highlights on the workbook viewers: `highlightText` and its navigation
 * are separate from the find on both `XlsxViewer` and `XlsxSheetViewer`.
 */
function makeFakeWorkbook(): XlsxWorkbook {
  const ws = {
    rows: [{
      index: 1,
      cells: [
        { row: 1, col: 1, value: { type: 'text', text: 'Revenue and returns' } },
        { row: 1, col: 2, value: { type: 'text', text: 'returns' } },
      ],
    }],
  };
  return {
    sheetCount: 1,
    sheetNames: ['Sheet1'],
    getWorksheet: () => Promise.resolve(ws),
    acquireWorksheetLease: async () => ({ worksheet: ws, release: () => undefined }),
    [xlsxCellDisplayTexts]: () => (cell: { value: { text: string } }) => [cell.value.text],
    destroy: vi.fn(),
    isHidden: () => false,
  } as unknown as XlsxWorkbook;
}

const GROUPS = [
  { id: 'system', terms: ['revenue'], color: 'rgba(0, 0, 255, 0.3)' },
  { id: 'user', terms: ['returns'], wholeWord: true },
];

async function exercise(viewer: XlsxViewer | XlsxSheetViewer): Promise<void> {
  const matches = await viewer.highlightText(GROUPS);
  expect(matches.map((match) => [match.matchIndex, match.group, match.term, match.text, match.location.ref])).toEqual([
    [0, 'system', 'revenue', 'Revenue', 'A1'],
    [1, 'user', 'returns', 'returns', 'A1'],
    [2, 'user', 'returns', 'returns', 'B1'],
  ]);
  expect(await viewer.findText('returns')).toHaveLength(2);
  expect((await viewer.findNext())?.matchIndex).toBe(0);
  expect(await viewer.highlightPrev()).toMatchObject({ matchIndex: 2, group: 'user', location: { ref: 'B1' } });
  expect(await viewer.highlightNext()).toMatchObject({ matchIndex: 0, group: 'system' });
  expect((await viewer.findNext())?.matchIndex).toBe(1);

  viewer.clearFind();
  expect(await viewer.findNext()).toBeNull();
  expect(await viewer.highlightNext()).toMatchObject({ matchIndex: 1 });

  await viewer.findText('returns');
  viewer.clearHighlights();
  expect(await viewer.highlightNext()).toBeNull();
  expect((await viewer.findNext())?.matchIndex).toBe(0);

  expect(await viewer.highlightText(GROUPS)).toHaveLength(3);
  // A repeated group id rejects, even with no terms, and clears nothing.
  await expect(viewer.highlightText([{ id: 'x', terms: [] }, { id: 'x', terms: [] }])).rejects.toThrow('group ids must be unique');
  expect(await viewer.highlightNext()).toMatchObject({ matchIndex: 0 });
  expect(await viewer.highlightText([])).toEqual([]);
  expect(await viewer.highlightNext()).toBeNull();
}

describe('XLSX viewers text highlights', () => {
  it('XlsxViewer highlights, navigates and clears independently of the find', async () => {
    installDom();
    const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    (viewer as unknown as { wb: XlsxWorkbook }).wb = makeFakeWorkbook();
    await exercise(viewer);
    await viewer.highlightText(GROUPS);
    viewer.destroy();
    expect(await viewer.highlightNext()).toBeNull();
  });

  it('XlsxSheetViewer exposes the same methods', async () => {
    installDom();
    const viewer = new XlsxSheetViewer(makeEl('canvas') as unknown as HTMLCanvasElement);
    (viewer as unknown as { engine: { wb: XlsxWorkbook } }).engine.wb = makeFakeWorkbook();
    await exercise(viewer);
    viewer.destroy();
  });
});
