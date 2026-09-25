import { beforeAll, describe, expect, it } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import type {
  BodyElement,
  BorderSpec,
  CellElement,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  SectionProps,
} from './types';
import type { ResolvedBorderSegment } from './layout/types.js';

// ECMA-376 §17.4.73 tl2br / §17.4.79 tr2bl: diagonal cell borders run between
// the physical corners of the cell box and take no part in edge conflicts.

const TEST_FONT = 'Synthetic Untabled Serif';

interface Stroke { lineWidth: number; points: { x: number; y: number }[]; color: string }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; strokes: Stroke[] } {
  let font = '10px serif';
  let path: { x: number; y: number }[] = [];
  const strokes: Stroke[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
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
    save() {}, restore() {}, transform() {}, setTransform() {}, translate() {}, scale() {}, rotate() {},
    getTransform() { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }; },
    beginPath() { path = []; },
    moveTo(x: number, y: number) { path.push({ x, y }); },
    lineTo(x: number, y: number) { path.push({ x, y }); },
    stroke() { strokes.push({ lineWidth: ctx.lineWidth, points: [...path], color: String(ctx.strokeStyle) }); },
    closePath() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000' as string, lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, strokes };
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeRecordingCanvas().canvas.getContext('2d'); }
  };
});

function paraOf(text: string): CellElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: TEST_FONT, fontFamilyEastAsia: TEST_FONT,
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT,
    widowControl: false,
  } as unknown as CellElement;
}

const border = (style: string, width = 0.5, color = 'ff0000'): BorderSpec => ({ style, width, color });

function cell(text: string, extra: Partial<DocTableCell> = {}): DocTableCell {
  return {
    content: [paraOf(text)],
    colSpan: 1,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null,
    vAlign: 'top',
    widthPt: 100,
    ...extra,
  } as DocTableCell;
}

function modelOf(rows: DocTableRow[]): DocxDocumentModel {
  const table = {
    colWidths: [100, 100],
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    layout: 'fixed',
  } as DocTable;
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...table } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

const row = (cells: DocTableCell[], rowHeight: number | null = 30, rule = 'exact'): DocTableRow =>
  ({ cells, rowHeight, rowHeightRule: rule, isHeader: false }) as DocTableRow;

function tableBorders(model: DocxDocumentModel): readonly ResolvedBorderSegment[] {
  const table = layoutDocument(model).pages[0]!.layers.body.find((node) => node.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  return table.borders;
}

const diagonal = (segment: ResolvedBorderSegment) =>
  segment.from.xPt !== segment.to.xPt && segment.from.yPt !== segment.to.yPt;

describe('ECMA-376 §17.4.73/§17.4.79 cell diagonal borders', () => {
  it('runs tl2br and tr2bl between the physical corners of the cell box', () => {
    const borders = tableBorders(modelOf([row([
      cell('a', { borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null, tl2br: border('single') } }),
      cell('b', { borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null, tr2bl: border('dashed', 1, '00ff00') } }),
    ])]));
    const diagonals = borders.filter(diagonal);
    expect(diagonals).toHaveLength(2);
    const [tl2br, tr2bl] = diagonals;
    expect(tl2br!.from).toEqual({ xPt: 0, yPt: 0 });
    expect(tl2br!.to).toEqual({ xPt: 100, yPt: 30 });
    expect(tl2br!.color).toBe('#ff0000');
    expect(tl2br!.style).toBe('solid');
    expect(tr2bl!.from).toEqual({ xPt: 200, yPt: 0 });
    expect(tr2bl!.to).toEqual({ xPt: 100, yPt: 30 });
    expect(tr2bl!.style).toBe('dashed');
  });

  it('spans a vertical merge and ignores nil diagonals', () => {
    const none = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
    const borders = tableBorders(modelOf([
      row([cell('a', { vMerge: true, borders: { ...none, tl2br: border('single'), tr2bl: border('nil') } }), cell('b')]),
      row([cell('', { vMerge: false }), cell('c')]),
    ]));
    const diagonals = borders.filter(diagonal);
    expect(diagonals).toHaveLength(1);
    expect(diagonals[0]!.from).toEqual({ xPt: 0, yPt: 0 });
    expect(diagonals[0]!.to).toEqual({ xPt: 100, yPt: 60 });
  });

  it('paints a double diagonal as two parallel rails', async () => {
    const none = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
    const model = modelOf([row([cell('a', { borders: { ...none, tl2br: border('double', 3) } }), cell('b')])]);
    const layout = layoutDocument(model);
    const { canvas, strokes } = makeRecordingCanvas();
    await paintLayoutPage(layout, 0, canvas, { dpr: 1, scale: 1 });
    const rails = strokes.filter((stroke) => stroke.points.length === 2
      && stroke.points[0]!.x !== stroke.points[1]!.x
      && stroke.points[0]!.y !== stroke.points[1]!.y);
    expect(rails).toHaveLength(2);
    // Each rail is a third of the 3pt width, offset symmetrically about the
    // corner-to-corner line along its normal.
    for (const rail of rails) expect(rail.lineWidth).toBeCloseTo(1);
    const [a, b] = rails;
    const mid = (stroke: Stroke) => ({
      x: (stroke.points[0]!.x + stroke.points[1]!.x) / 2,
      y: (stroke.points[0]!.y + stroke.points[1]!.y) / 2,
    });
    const centre = { x: (mid(a!).x + mid(b!).x) / 2, y: (mid(a!).y + mid(b!).y) / 2 };
    expect(centre.x).toBeCloseTo(50);
    expect(centre.y).toBeCloseTo(15);
    expect(Math.hypot(mid(a!).x - mid(b!).x, mid(a!).y - mid(b!).y)).toBeCloseTo(2);
  });
});
