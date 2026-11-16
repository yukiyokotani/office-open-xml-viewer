import { describe, expect, it } from 'vitest';
import { FindAdapter, type FindAdapterHost } from './find-adapter.js';
import { xlsxCellDisplayTexts, type XlsxWorkbook } from '../../workbook.js';
import type { Worksheet } from '../../types.js';

/**
 * The find overlay of the workbook viewer: text highlights (one box per group
 * per cell, nested, the active one ringed) under the find boxes.
 */
interface Box { cssText: string; }

function makeHost(cells: Record<string, string>[]) {
  const sheets = cells.map((sheet) => ({
    freezeRows: 0,
    freezeCols: 0,
    colWidths: {},
    rowHeights: {},
    rows: [{
      index: 1,
      cells: Object.entries(sheet).map(([ref, text]) => ({
        row: Number(ref.slice(1)), col: ref.charCodeAt(0) - 64, value: { type: 'text', text },
      })),
    }],
  })) as unknown as Worksheet[];
  const workbook = {
    sheetNames: cells.map((_, index) => `S${index + 1}`),
    acquireWorksheetLease: async (sheet: number) => ({ worksheet: sheets[sheet], release: () => undefined }),
    [xlsxCellDisplayTexts]: () => (cell: { value: { text: string } }) => [cell.value.text],
  } as unknown as XlsxWorkbook;
  const boxes: Box[] = [];
  const state = { sheet: 0, scrolled: [] as string[] };
  const host: FindAdapterHost = {
    ownerDocument: {
      createElement: () => {
        const box: Box = { cssText: '' };
        return { style: box };
      },
    } as unknown as Document,
    overlayHost: {
      clearFind: () => { boxes.length = 0; },
      appendFind: (element: { style: Box }) => { boxes.push(element.style); },
    } as unknown as FindAdapterHost['overlayHost'],
    workbook: () => workbook,
    sheetCount: () => sheets.length,
    worksheet: () => sheets[state.sheet],
    currentSheet: () => state.sheet,
    scale: () => 1,
    highlightColors: () => ({ match: 'rgba(1, 2, 3, 0.4)', active: 'rgba(4, 5, 6, 0.7)' }),
    cellRect: (row, col) => ({ x: 100 * col, y: 30 * row, w: 100, h: 30 }),
    screenX: (x) => x,
    goToSheet: async (index) => { state.sheet = index; },
    scrollCellIntoView: (row, col) => { state.scrolled.push(`${row},${col}`); },
  };
  return { adapter: new FindAdapter(host), boxes, state };
}

const geometry = (box: Box) => box.cssText.match(/left:[^;]+;top:[^;]+;width:[^;]+;height:[^;]+/)?.[0];
const paint = (box: Box) => [
  box.cssText.match(/background:([^;]+)/)?.[1], box.cssText.match(/border:([^;]+)/)?.[1],
];

describe('FindAdapter text highlights', () => {
  const groups = [
    { id: 'system', terms: ['revenue'], color: 'rgba(0, 0, 255, 0.3)' },
    { id: 'user', terms: ['returns'], color: 'rgba(255, 0, 0, 0.3)' },
  ];

  it('boxes each matched cell once per group, nested, under the find boxes', async () => {
    const { adapter, boxes } = makeHost([{ A1: 'revenue and returns, revenue', B1: 'returns' }]);
    const matches = await adapter.highlight(groups);
    expect(matches.map((match) => [match.group, match.location.ref])).toEqual([
      ['system', 'A1'], ['user', 'A1'], ['system', 'A1'], ['user', 'B1'],
    ]);
    await adapter.find('and');
    expect(boxes.map(paint)).toEqual([
      ['rgba(0, 0, 255, 0.3)', '2px solid rgba(0, 0, 255, 0.3)'],
      ['rgba(255, 0, 0, 0.3)', '2px solid rgba(255, 0, 0, 0.3)'],
      ['rgba(255, 0, 0, 0.3)', '2px solid rgba(255, 0, 0, 0.3)'],
      ['rgba(1, 2, 3, 0.4)', '2px solid rgba(1, 2, 3, 0.4)'],
    ]);
    expect(boxes.slice(0, 3).map(geometry)).toEqual([
      'left:100px;top:30px;width:100px;height:30px',
      'left:103px;top:33px;width:94px;height:24px',
      'left:200px;top:30px;width:100px;height:30px',
    ]);
    // A find box is never inset.
    expect(geometry(boxes[3])).toBe('left:100px;top:30px;width:100px;height:30px');
  });

  it('keeps a visible core when many groups share a small cell', async () => {
    const { adapter, boxes } = makeHost([{ A1: 'a b c d e f g h' }]);
    await adapter.highlight('abcdefgh'.split('').map((term) => ({ id: term, terms: [term] })));
    expect(boxes).toHaveLength(8);
    expect(geometry(boxes[7])).toBe('left:113px;top:43px;width:74px;height:4px');
  });

  it('navigates across sheets, rings the active cell, and leaves the find cursor alone', async () => {
    const { adapter, boxes, state } = makeHost([{ A1: 'revenue' }, { C2: 'returns' }]);
    await adapter.highlight(groups);
    await adapter.find('revenue');
    expect(await adapter.next()).toMatchObject({ matchIndex: 0 });
    expect(await adapter.highlightPrev()).toMatchObject({ group: 'user', location: { sheet: 1, ref: 'C2' } });
    expect(state.sheet).toBe(1);
    expect(state.scrolled.at(-1)).toBe('2,3');
    expect(boxes.map((box) => box.cssText.includes('outline:2px solid'))).toEqual([true]);
    expect(await adapter.highlightNext()).toMatchObject({ group: 'system', location: { sheet: 0, ref: 'A1' } });
    expect(boxes.map((box) => box.cssText.includes('outline:2px solid'))).toEqual([true, false]);
    expect(boxes[1].cssText).toContain('background:rgba(4, 5, 6, 0.7)');
  });

  it('clears each kind on its own', async () => {
    const { adapter, boxes } = makeHost([{ A1: 'revenue' }]);
    await adapter.highlight(groups);
    await adapter.find('revenue');
    adapter.clear();
    expect(boxes).toHaveLength(1);
    expect(await adapter.next()).toBeNull();
    await adapter.find('revenue');
    adapter.clearHighlights();
    expect(boxes.map((box) => paint(box)[0])).toEqual(['rgba(1, 2, 3, 0.4)']);
    expect(await adapter.highlightNext()).toBeNull();
    await adapter.highlight(groups);
    adapter.invalidate();
    adapter.updateOverlay();
    expect(boxes).toEqual([]);
  });
});
