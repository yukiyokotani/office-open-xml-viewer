import { beforeAll, describe, expect, it } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import { layoutDocument } from './document-layout.js';
import { textRunGeometryForPage } from './layout/text-index.js';
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

// ECMA-376 §17.4.72 `<w:tcPr><w:textDirection>`: a rotated cell lays its lines
// along the cell height. tbRl turns the text frame a quarter clockwise (first
// line at the right edge, reading top to bottom), btLr a quarter
// counter-clockwise (first line at the left edge, reading bottom to top), and
// tbRlV is tbRl with East Asian glyphs upright (§17.18.93).

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

function firstRowHeight(table: DocTable): number {
  const retained = layoutDocument(modelOf(table)).pages[0]?.layers.body
    .find((node) => node.kind === 'table');
  if (!retained || retained.kind !== 'table') throw new Error('Table was not laid out');
  return retained.rows[0]!.heightPt;
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

function applyMatrix(m: { a: number; b: number; c: number; d: number; e: number; f: number }, x: number, y: number) {
  return { x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f };
}

describe('ECMA-376 §17.4.72 cell text direction', () => {
  it('tbRl turns the text a quarter clockwise', async () => {
    const calls = await render(tableOf([
      row([cell('ABCDE', { textDirection: 'tbRl' }), cell('x')]),
    ]));
    const text = calls.find((call) => call.text === 'A');
    expect(text).toBeDefined();
    // Clockwise quarter turn: each wrapped column advances from right to left.
    expect(text!.matrix.a).toBeCloseTo(0);
    expect(text!.matrix.b).toBeCloseTo(1);
    expect(calls.filter((call) => /[A-E]/.test(call.text)).map((call) => call.text).join('')).toBe('ABCDE');
    // First line at the right edge of the 100pt cell: its baseline sits left of x=100.
    expect(text!.x).toBeLessThan(100);
    expect(text!.x).toBeGreaterThan(80);
    expect(text!.y).toBeCloseTo(0, 0);
    const neighbour = calls.find((call) => call.text === 'x');
    expect(neighbour!.matrix.b).toBeCloseTo(0);
  });

  it('auto rotated rows wrap into columns instead of growing with text length', () => {
    for (const direction of ['tbRl', 'btLr']) {
      const height = (text: string) => firstRowHeight({
        ...tableOf([row([cell(text, { textDirection: direction }), cell('N')])]),
        colWidths: [220, 140],
      });
      expect(height('ABCDEFGH')).toBeCloseTo(height('A'), 1);
    }
  });

  it('grows a rotated row once its wrapped columns no longer fit the cell width', () => {
    const height = (text: string) => firstRowHeight({
      ...tableOf([row([cell(text, { textDirection: 'tbRl' }), cell('N')])]),
      colWidths: [220, 140],
    });
    expect(height('A'.repeat(30))).toBeGreaterThan(height('A'.repeat(8)));
  });

  it('rotated auto and atLeast rows honor margins, authored floors, and a taller neighbor', () => {
    const height = (rotated: DocTableCell, neighbor: DocTableCell, floor: number | null = null, rule = 'auto') =>
      firstRowHeight({ ...tableOf([row([rotated, neighbor], floor, rule)]), colWidths: [220, 140] });
    for (const direction of ['tbRl', 'btLr']) {
      const rotated = cell('ABCD', { textDirection: direction });
      const shortNeighbor = cell('N');
      const natural = height(rotated, shortNeighbor);
      expect(height(rotated, shortNeighbor, 5, 'atLeast')).toBeCloseTo(natural, 1);
      expect(height(rotated, shortNeighbor, 60, 'atLeast')).toBeCloseTo(60, 1);
      expect(height(rotated, shortNeighbor, 20, 'exact')).toBeCloseTo(20, 1);
      expect(height(cell('ABCD', { textDirection: direction, marginTop: 12, marginBottom: 8 }), shortNeighbor))
        .toBeCloseTo(natural + 20, 1);
      const narrowGlyphs = {
        ...paraOf('ABCD'),
        runs: [{ type: 'text', ...textRun('ABCD'), charScale: 0.7 }],
      } as CellElement;
      expect(height(cell('ABCD', {
        textDirection: direction, content: [narrowGlyphs], marginTop: 12, marginBottom: 8,
      }), shortNeighbor)).toBeCloseTo(natural + 20, 1);
      const tallNeighbor = cell('N', { content: [paraOf('1'), paraOf('2'), paraOf('3'), paraOf('4'), paraOf('5')] });
      expect(height(rotated, tallNeighbor)).toBeCloseTo(height(cell('ABCD'), tallNeighbor), 1);
      expect(height(cell('AB', { textDirection: direction, content: [paraOf('AB'), paraOf('CDE')] }), shortNeighbor))
        .toBeCloseTo(natural, 1);
    }
  });

  it('btLr turns the text a quarter counter-clockwise starting at the bottom-left', async () => {
    const calls = await render(tableOf([
      row([cell('ABCDE', { textDirection: 'btLr' }), cell('x')]),
    ]));
    const text = calls.find((call) => call.text === 'A')!;
    expect(text.matrix.b).toBeCloseTo(-1);
    expect(text.x).toBeGreaterThan(0);
    expect(text.x).toBeLessThan(20);
    // The 10pt auto row wraps five glyphs into adjacent columns.
    expect(text.y).toBeCloseTo(10, 0);
  });

  it('keeps an exact row height and aligns lines along it', async () => {
    const calls = await render(tableOf([
      row([cell('AB', { textDirection: 'tbRl', alignment: 'right' }), cell('x')], 80, 'exact'),
    ]));
    const text = calls.find((call) => call.text.includes('AB'))!;
    // Right-aligned along the 80pt line: the two 10pt glyphs start at y=60.
    expect(text.y).toBeCloseTo(60, 0);
  });

  it('lets a taller neighbour lengthen the line and centres text along it', async () => {
    const calls = await render(tableOf([
      row([
        cell('AB', { textDirection: 'tbRl', alignment: 'center' }),
        cell('tall', { content: [paraOf('1'), paraOf('2'), paraOf('3'), paraOf('4'), paraOf('5')] }),
      ]),
    ]));
    const text = calls.find((call) => call.text.includes('AB'))!;
    const lastLine = calls.find((call) => call.text === '5')!;
    const rowHeight = lastLine.y + 2; // descent of the last 10pt line
    expect(text.y).toBeCloseTo((rowHeight - 20) / 2, 0);
  });

  it('tbRlV keeps East Asian glyphs upright in a vertically merged cell', async () => {
    const calls = await render(tableOf([
      row([cell('連絡先', { textDirection: 'tbRlV', vMerge: true, vAlign: 'center', alignment: 'center' }), cell('a')]),
      row([cell('', { vMerge: false }), cell('b')]),
      row([cell('', { vMerge: false }), cell('c')]),
    ]));
    const glyphs = calls.filter((call) => /[連絡先]/.test(call.text));
    expect(glyphs.length).toBeGreaterThan(0);
    // Upright glyphs are painted with the page orientation restored.
    for (const glyph of glyphs) expect(glyph.matrix.b).toBeCloseTo(0);
    expect(glyphs.map((g) => g.text).join('')).toBe('連絡先');
  });

  it('moves the rotated text frame with a table placed below other content', async () => {
    const lead = paraOf('lead') as unknown as BodyElement;
    const calls = await render(tableOf([
      row([cell('ABCDE', { textDirection: 'tbRl' }), cell('x')]),
    ]), [lead, lead, lead]);
    const text = calls.find((call) => call.text === 'A')!;
    const neighbour = calls.find((call) => call.text === 'x')!;
    expect(text.matrix.b).toBeCloseTo(1);
    expect(text.x).toBeGreaterThan(80);
    expect(text.x).toBeLessThan(100);
    // The line starts at the row top, which is where the neighbour's line box starts.
    expect(text.y).toBeGreaterThan(20);
    expect(text.y).toBeCloseTo(neighbour.y - 8, 0);
  });

  it('projects rotated text runs to the page for hit-testing where they are painted', async () => {
    const lead = paraOf('lead') as unknown as BodyElement;
    for (const [textDirection, b] of [['tbRl', 1], ['btLr', -1]] as const) {
      const table = tableOf([row([cell('ABCDE', { textDirection }), cell('x')])]);
      const runs = textRunGeometryForPage(layoutDocument(modelOf(table, [lead])), 0)
        .filter((geometry) => /^[A-E]$/.test(geometry.placement.text));
      expect(runs.map((run) => run.placement.text).join('')).toBe('ABCDE');
      expect(runs.every((run) => Math.abs(run.pointToPage.b - b) < 0.01)).toBe(true);
      const points = runs.flatMap(({ pointToPage, placement: { bounds } }) => [
        applyMatrix(pointToPage, bounds.xPt, bounds.yPt),
        applyMatrix(pointToPage, bounds.xPt + bounds.widthPt, bounds.yPt + bounds.heightPt),
      ]);
      // Five one-glyph columns remain inside the 10pt row below the lead.
      const top = Math.min(...points.map((point) => point.y));
      const bottom = Math.max(...points.map((point) => point.y));
      const left = Math.min(...points.map((point) => point.x));
      const right = Math.max(...points.map((point) => point.x));
      expect(bottom - top).toBeCloseTo(10, 0);
      expect(top).toBeGreaterThanOrEqual(9);
      expect(left).toBeGreaterThanOrEqual(0);
      expect(right).toBeLessThanOrEqual(100);
      const painted = (await render(table, [lead])).find((call) => call.text === 'A')!;
      expect(painted.matrix.b).toBeCloseTo(b);
      expect(painted.x).toBeGreaterThanOrEqual(left - 0.01);
      expect(painted.x).toBeLessThanOrEqual(right + 0.01);
      expect(painted.y).toBeGreaterThanOrEqual(top - 0.01);
      expect(painted.y).toBeLessThanOrEqual(bottom + 0.01);
    }
  });

  it('leaves lrTbV and tbLrV cells horizontal (not yet rendered rotated)', async () => {
    for (const textDirection of ['lrTbV', 'tbLrV']) {
      const calls = await render(tableOf([row([cell('ABC', { textDirection }), cell('x')])]));
      const text = calls.find((call) => call.text.includes('ABC'))!;
      expect(text.matrix.b).toBeCloseTo(0);
    }
  });
});
