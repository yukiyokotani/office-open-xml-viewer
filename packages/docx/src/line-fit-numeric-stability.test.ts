import { describe, expect, it } from 'vitest';
import type { DocTable } from './types.js';
import { layoutLines, type LayoutTextSeg } from './line-layout.js';
import { acquireRetainedTable } from './layout/table-acquisition.js';
import type { LayoutServices, ParagraphLayout, TableFormatInput } from './layout/types.js';

// Synthetic Canvas advance at an AutoFit minimum. Production cell acquisition
// must return a width that can still fit the very advance used by that minimum.
const advance = Math.fround(42.4);
const margin = 5.4;
const indent = 21.6;
const borders = { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };

function renderCell(column: number, spacing = 0, firstIndent = indent) {
  const table = {
    rows: [{ cells: [{ content: [{ type: 'paragraph', runs: [] }], colSpan: 1,
      vMerge: null, borders, background: null, vAlign: 'top' }], gridBefore: 0,
      gridAfter: 0, isHeader: false, cantSplit: false }],
    colWidths: [column], borders, cellMarginTop: 0, cellMarginRight: margin,
    cellMarginBottom: 0, cellMarginLeft: margin, jc: 'left', bidiVisual: false,
  } as unknown as DocTable;
  const format: TableFormatInput = {
    effectiveStyleId: null, ordinaryFlow: true, positioning: null, firstRowException: null,
    rows: [{ height: null, cantSplit: false, repeatedHeader: false,
      cellSpacingPt: spacing, justification: null, exception: null,
      cells: [{ marginsPt: { top: 0, right: margin, bottom: 0, left: margin } }] }],
  };
  const ctx = {
    font: '12px serif', letterSpacing: '0px',
    measureText: (text: string) => ({
      width: text === 'ABCDEF' ? advance : [...text].length * advance / 6,
      fontBoundingBoxAscent: 11, fontBoundingBoxDescent: 3,
      actualBoundingBoxAscent: 11, actualBoundingBoxDescent: 3,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
  const text: LayoutTextSeg = {
    text: 'ABCDEF', bold: true, italic: false, underline: false,
    strikethrough: false, fontSize: 12, color: null, fontFamily: 'Synthetic Face',
    vertAlign: null, measuredWidth: 0,
  };
  let result: string[] = [];
  const retained = acquireRetainedTable(table, [column], column, { yPt: 0 },
    { story: 'body', storyInstance: 'body', path: [0] }, {
      layoutServices: () => ({}) as LayoutServices,
      tableFormat: () => format,
      resolveColumns: () => [],
      createCellState: (state) => state,
      acquireParagraph: (_state, _paragraph, width) => {
        result = layoutLines(ctx, [text], width, firstIndent, 1)
          .map((line) => line.segments.map((part) => 'text' in part ? part.text : '').join(''));
        const bounds = { xPt: 0, yPt: 0, widthPt: width, heightPt: 8 };
        return { kind: 'paragraph', id: 'synthetic',
          source: { story: 'body', storyInstance: 'body', path: [0, 0, 0] },
          flowDomainId: 'table-cell', ordinaryFlow: true,
          flowBounds: bounds, inkBounds: bounds, advancePt: 8,
          alignment: 'left', bidi: false, spacing: { beforePt: 0, afterPt: 0 },
          borders: [], lines: [],
        } as unknown as ParagraphLayout;
      },
      registerFloatingTable: () => null,
      advanceState: () => {},
    });
  return { lines: result, retainedWidth: retained.layout.rows[0]?.cells[0]?.contentBounds.widthPt };
}

describe('table intrinsic line-fit boundary', () => {
  it('keeps an equal measured minimum after margin acquisition but rejects one twip less', () => {
    const column = advance + indent + (margin + margin);
    const equal = renderCell(column);
    expect(equal.lines).toEqual(['ABCDEF']);
    expect(equal.retainedWidth).toBe(advance + indent);
    expect(renderCell(column - 0.05).lines).toHaveLength(2);
  });

  it('keeps spacing and zero-indent equality distinct from a real deficit', () => {
    const spacing = 1;
    const column = advance + (margin + margin) + 2 * spacing;
    const equal = renderCell(column, spacing, 0);
    expect(equal.lines).toEqual(['ABCDEF']);
    expect(equal.retainedWidth).toBe(advance);
    expect(renderCell(column - 0.05, spacing, 0).lines).toHaveLength(2);
  });
});
