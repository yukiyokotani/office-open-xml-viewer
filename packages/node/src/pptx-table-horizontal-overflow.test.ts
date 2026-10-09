import { describe, expect, it } from 'vitest';
import { renderTable, type PptxTextRunInfo } from '../../pptx/src/renderer.ts';
import type { TableCell, TableElement, TextBody } from '../../pptx/src/types.ts';
import { loadSkiaForTests } from './test-imports.ts';

const skia = await loadSkiaForTests();
const { Canvas } = (skia ?? {}) as typeof import('skia-canvas');
const EMU = 12700;

function body(text: string, secondParagraph = false): TextBody {
  const paragraph = { alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: { type: 'pct', val: 100000 },
    bullet: { type: 'none' }, eaLnBrk: true,
    runs: [{ type: 'text', text, fontSize: 20, fontFamily: 'Arial' }],
  };
  return { verticalAnchor: 't', lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    vert: 'horz', wrap: 'none', autoFit: 'none',
    paragraphs: [paragraph, ...(secondParagraph ? [paragraph] : [])],
    defaultFontSize: null, defaultBold: null, defaultItalic: null,
  } as unknown as TextBody;
}

function table(overflow?: 'clip' | 'overflow', verticalSpill = false): TableElement {
  const empty: TableCell = { textBody: null, fill: null, borderL: null, borderR: null,
    borderT: null, borderB: null, gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false };
  const first = { ...empty, textBody: body(verticalSpill ? 'A' : 'AAAAAAAA', verticalSpill),
    ...(overflow === undefined ? {} : { horzOverflow: overflow }) };
  const height = (verticalSpill ? 12 : 80) * EMU;
  return { type: 'table', x: 40 * EMU, y: 40 * EMU, width: 120 * EMU, height,
    rotation: 0, flipH: false, flipV: false, cols: [60 * EMU, 60 * EMU],
    rows: [{ height, cells: [first, empty] }] };
}

function paint(input: TableElement) {
  const canvas = new Canvas(240, 240);
  const runs: PptxTextRunInfo[] = [];
  renderTable(canvas.getContext('2d') as unknown as CanvasRenderingContext2D,
    input, 1 / EMU, undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 }, run => runs.push(run));
  const rgba = canvas.getContext('2d').getImageData(0, 0, 240, 240).data;
  let left = 240, top = 240, right = -1, bottom = -1;
  for (let y = 0; y < 240; y++) for (let x = 0; x < 240; x++) {
    if (rgba[(y * 240 + x) * 4 + 3] === 0) continue;
    left = Math.min(left, x); right = Math.max(right, x);
    top = Math.min(top, y); bottom = Math.max(bottom, y);
  }
  return { bounds: { left, top, right, bottom }, runs };
}

describe.skipIf(!skia)('PPTX cell horizontal overflow on actual Canvas', () => {
  it('clips at the cell edge by default and keeps explicitly overflowing text visible', () => {
    // ECMA-376 §21.1.3.17 and CT_TableCellProperties: omitted horzOverflow is
    // clip. Empty neighbouring cells have no fill to conceal overflow defects.
    const clipped = paint(table());
    const overflowing = paint(table('overflow'));
    expect(clipped.bounds.right).toBeLessThan(100);
    expect(overflowing.bounds.right).toBeGreaterThan(100);
    expect(clipped.runs.map(run => run.text).join('')).toBe('AAAAAAAA');
    expect(clipped.runs).not.toHaveLength(0);
    for (const run of clipped.runs) {
      expect(run.cellHorzOverflow).toBe('clip');
      expect([run.shapeX, run.shapeY, run.shapeW, run.shapeH]).toEqual([40, 40, 60, 80]);
    }
    expect(overflowing.runs.every(run => run.cellHorzOverflow === 'overflow')).toBe(true);
  });

  it('preserves existing vertical spill while restricting the physical horizontal axis', () => {
    // The change must not introduce a new vertical overflow/row growth rule.
    // A second paragraph exceeds this authored row; only x clipping is new.
    const result = paint(table(undefined, true));
    expect(result.bounds.right).toBeLessThan(100);
    expect(result.bounds.bottom).toBeGreaterThan(52);
  });

  it('leaves rotated and stacked vertical cell bodies on their unclipped path', () => {
    // The physical-x clip is admitted only for horizontal bodies. A 140px glyph
    // line is wider than the 60px cell in both modes, so a clip would hold ink
    // inside x=[40,100]; the omitted (clip) default must not apply here.
    for (const vert of ['vert', 'wordArtVert']) {
      const input = table();
      const textBody = input.rows[0].cells[0].textBody!;
      textBody.vert = vert;
      const textRun = textBody.paragraphs[0].runs[0];
      if (textRun.type === 'text') { textRun.text = 'W'; textRun.fontSize = 140; }
      const result = paint(input);
      expect(result.bounds.left < 40 || result.bounds.right > 100, vert).toBe(true);
      expect(result.runs.map(run => run.text).join(''), vert).toBe('W');
      for (const run of result.runs) {
        expect(run.tableCell, vert).toEqual({ row: 0, column: 0 });
        expect(run.cellHorzOverflow, vert).toBeUndefined();
        expect(run.textBodyRotation, vert).toBe(vert === 'vert' ? 90 : undefined);
      }
    }
  });

  it('applies the cell clip before graphic-frame rotation and reflection', () => {
    const clippedTable = table();
    clippedTable.rotation = 90; clippedTable.flipH = true;
    const overflowingTable = table('overflow');
    overflowingTable.rotation = 90; overflowingTable.flipH = true;
    // Cell x=[40,100], reflected about frame x=100 then rotated about y=80,
    // becomes device y=[80,140]. This independent rectangle is not a glyph box.
    expect(paint(clippedTable).bounds.top).toBeGreaterThanOrEqual(80);
    expect(paint(overflowingTable).bounds.top).toBeLessThan(80);
  });
});
