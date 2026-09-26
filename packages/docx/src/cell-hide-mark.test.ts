import { beforeAll, describe, expect, it } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

// ECMA-376 §17.4.21 `<w:tcPr><w:hideMark>`: a cell's end-of-cell mark does not
// count toward the row height. The final paragraph that holds only that mark
// (the whole content of an empty cell) owns no height, per cell. Observed in
// Word's PDF output: such a row ends at the cell's last text paragraph even
// when other rows' cells have content, and an all-empty hideMark row
// collapses to its 1pt minimum.

interface Matrix { a: number; b: number; c: number; d: number; e: number; f: number }
interface FillTextCall { text: string; x: number; y: number; matrix: Matrix }

const IDENTITY: Matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

function multiply(m: Matrix, n: Matrix): Matrix {
  return {
    a: m.a * n.a + m.c * n.b,
    b: m.b * n.a + m.d * n.b,
    c: m.a * n.c + m.c * n.d,
    d: m.b * n.c + m.d * n.d,
    e: m.a * n.e + m.c * n.f + m.e,
    f: m.b * n.e + m.d * n.f + m.f,
  };
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: FillTextCall[] } {
  let font = '10px serif';
  let matrix = IDENTITY;
  const stack: Matrix[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
        actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: [...s].length * p,
      } as TextMetrics;
    },
    save() { stack.push(matrix); },
    restore() { matrix = stack.pop() ?? matrix; },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) {
      matrix = multiply(matrix, { a, b, c, d, e, f });
    },
    setTransform(a: number, b: number, c: number, d: number, e: number, f: number) {
      matrix = { a, b, c, d, e, f };
    },
    getTransform() { return { ...matrix }; },
    translate(x: number, y: number) { matrix = multiply(matrix, { ...IDENTITY, e: x, f: y }); },
    scale(x: number, y: number) { matrix = multiply(matrix, { ...IDENTITY, a: x, d: y }); },
    rotate(angle: number) {
      const cos = Math.cos(angle);
      const sin = Math.sin(angle);
      matrix = multiply(matrix, { a: cos, b: sin, c: -sin, d: cos, e: 0, f: 0 });
    },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {}, drawImage() {},
    clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) {
      calls.push({
        text,
        x: matrix.a * x + matrix.c * y + matrix.e,
        y: matrix.b * x + matrix.d * y + matrix.f,
        matrix: { ...matrix },
      });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

const TEST_FONT = 'Synthetic Untabled Serif';

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: TEST_FONT, fontFamilyEastAsia: TEST_FONT,
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  };
}

function paraOf(text: string, alignment = 'left'): CellElement {
  return {
    type: 'paragraph',
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT,
    widowControl: false,
  } as unknown as CellElement;
}

function cell(text: string, options: Partial<DocTableCell> & { alignment?: string } = {}): DocTableCell {
  const { alignment, ...rest } = options;
  return {
    content: [paraOf(text, alignment)],
    colSpan: 1,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null,
    vAlign: 'top',
    widthPt: 100,
    ...rest,
  } as DocTableCell;
}

function tableOf(rows: DocTableRow[]): DocTable {
  return {
    colWidths: [100, 100],
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    layout: 'fixed',
  } as DocTable;
}

function row(cells: DocTableCell[], rowHeight: number | null = null, rule = 'auto'): DocTableRow {
  return { cells, rowHeight, rowHeightRule: rule, isHeader: false } as DocTableRow;
}

function modelOf(table: DocTable, before: BodyElement[] = []): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [...before, { type: 'table', ...table } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(table: DocTable, before: BodyElement[] = []): Promise<FillTextCall[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(modelOf(table, before), canvas, 0, { dpr: 1, width: 400 });
  return calls;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeRecordingCanvas().canvas.getContext('2d'); }
  };
});

function emptyPara(): CellElement {
  return { ...(paraOf('') as unknown as DocParagraph), runs: [] } as unknown as CellElement;
}

async function nextRowTop(first: DocTableRow): Promise<number> {
  const calls = await render(tableOf([first, row([cell('next'), cell('')])]));
  const next = calls.find((call) => call.text === 'next');
  expect(next).toBeDefined();
  return next!.y - 8; // 10pt line, 8pt ascent above the baseline
}

describe('ECMA-376 §17.4.21 hideMark', () => {
  it('drops a hideMark cell\'s final empty paragraph from the row height', async () => {
    const content = [paraOf('a'), paraOf('b'), emptyPara()];
    expect(await nextRowTop(row([cell('', { content }), cell('x')]))).toBeCloseTo(30, 5);
    expect(await nextRowTop(row([cell('', { content, hideMark: true }), cell('x')])))
      .toBeCloseTo(20, 5);
  });

  it('collapses a row of empty hideMark cells to its minimum height', async () => {
    const empty = (hideMark: boolean) => cell('', { content: [emptyPara()], hideMark });
    expect(await nextRowTop(row([empty(false), empty(false)], 1, 'atLeast'))).toBeCloseTo(10, 5);
    expect(await nextRowTop(row([empty(true), empty(true)], 1, 'atLeast'))).toBeCloseTo(1, 5);
    // Per cell: an empty cell without hideMark keeps its mark's line.
    expect(await nextRowTop(row([empty(true), empty(false)], 1, 'atLeast'))).toBeCloseTo(10, 5);
  });

  it('keeps a final paragraph with text and leaves other cells unchanged', async () => {
    expect(await nextRowTop(row([
      cell('', { content: [paraOf('a'), paraOf('b')], hideMark: true }),
      cell('', { content: [emptyPara()], hideMark: true }),
    ]))).toBeCloseTo(20, 5);
  });
});
