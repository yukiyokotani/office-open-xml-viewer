import { afterEach, describe, expect, it, vi } from 'vitest';
import { renderSlide, renderTable, type PptxTextRunInfo } from './renderer.js';
import type { Slide, TableCell, TableElement, TableRow, TextBody } from './types.js';

const EMU = 12_700;
const SCALE = 1 / EMU;
const COL = 60 * EMU;
const renderContext = { themeMajorFont: null, themeMinorFont: null, dpr: 1 } as const;

afterEach(() => vi.restoreAllMocks());

function recordingContext(): CanvasRenderingContext2D {
  const ctx = {
    canvas: { width: 1000, height: 1000 },
    lineWidth: 1,
    strokeStyle: '#000000',
    fillStyle: '#000000',
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
    globalAlpha: 1,
    save() {}, restore() {},
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {},
    fill() {}, fillRect() {}, strokeRect() {}, clearRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setTransform() {}, transform() {}, resetTransform() {},
    setLineDash() {}, getLineDash() { return []; },
    drawImage() {}, arc() {}, arcTo() {}, ellipse() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    createPattern() { return null; },
    measureText: () => ({
      width: 0,
      fontBoundingBoxAscent: 8,
      fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8,
      actualBoundingBoxDescent: 2,
    } as TextMetrics),
    fillText() {}, strokeText() {},
    font: '10px sans-serif',
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    letterSpacing: '0px',
    globalCompositeOperation: 'source-over' as GlobalCompositeOperation,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function cell(overrides: Partial<TableCell> = {}): TableCell {
  return {
    textBody: null, fill: null,
    borderL: null, borderR: null, borderT: null, borderB: null,
    diagonalTL: null, diagonalTR: null,
    gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false,
    ...overrides,
  } as TableCell;
}

function tableOf(rows: TableCell[][], cols: number[]): TableElement {
  const rowHeight = 20 * EMU;
  const tableRows: TableRow[] = rows.map((cells) => ({ height: rowHeight, cells }));
  return {
    type: 'table', x: 0, y: 0,
    width: cols.reduce((sum, width) => sum + width, 0),
    height: rowHeight * rows.length,
    rotation: 0, flipH: false, flipV: false,
    cols, rows: tableRows,
  };
}

function textBody(text: string): TextBody {
  return ({
    verticalAnchor: 'ctr',
    paragraphs: [{
      alignment: 'l', marL: 0, marR: 0, indent: 0,
      spaceBefore: null, spaceAfter: null, spaceLine: null,
      runs: [{ type: 'text', text, fontSize: 12, fontFamily: 'Arial' }],
      bullet: { type: 'none' }, eaLnBrk: true,
    }],
    defaultFontSize: null, defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  }) as unknown as TextBody;
}

function collectRuns(table: TableElement): PptxTextRunInfo[] {
  const runs: PptxTextRunInfo[] = [];
  renderTable(
    recordingContext(),
    table,
    SCALE,
    undefined,
    renderContext,
    (run) => runs.push(run),
  );
  return runs;
}

describe('DrawingML table text selection runs', () => {
  it('emits table cell identity and geometry in row-major order', () => {
    const table = tableOf([[
      cell({ textBody: textBody('A') }),
      cell({ textBody: textBody('B') }),
    ]], [COL, COL]);
    table.id = '27';
    table.rows[0].height = 30 * EMU;
    table.height = 30 * EMU;

    const runs = collectRuns(table);

    expect(runs.map((run) => ({
      text: run.text,
      shapeId: run.shapeId,
      tableCell: run.tableCell,
      shapeX: run.shapeX,
      shapeY: run.shapeY,
      shapeW: run.shapeW,
      shapeH: run.shapeH,
      rotation: run.rotation,
    }))).toEqual([
      {
        text: 'A', shapeId: '27', tableCell: { row: 0, column: 0 },
        shapeX: 0, shapeY: 0, shapeW: 60, shapeH: 30, rotation: 0,
      },
      {
        text: 'B', shapeId: '27', tableCell: { row: 0, column: 1 },
        shapeX: 60, shapeY: 0, shapeW: 60, shapeH: 30, rotation: 0,
      },
    ]);
  });

  it('keeps empty cells as semantic gaps instead of fabricating selectable text', () => {
    const table = tableOf([[
      cell({ textBody: textBody('A') }),
      cell({ textBody: null }),
      cell({ textBody: textBody('C') }),
    ]], [COL, COL, COL]);

    expect(collectRuns(table).map((run) => ({
      text: run.text,
      tableCell: run.tableCell,
    }))).toEqual([
      { text: 'A', tableCell: { row: 0, column: 0 } },
      { text: 'C', tableCell: { row: 0, column: 2 } },
    ]);
  });

  it('skips selection transform math when no text-run callback is requested', () => {
    const table = tableOf([[cell({ textBody: textBody('A') })]], [COL]);
    const cos = vi.spyOn(Math, 'cos');
    const sin = vi.spyOn(Math, 'sin');

    renderTable(recordingContext(), table, SCALE, undefined, renderContext);

    expect(cos).not.toHaveBeenCalled();
    expect(sin).not.toHaveBeenCalled();
  });

  it('maps cell frames through table rotation and horizontal flip', () => {
    const table = tableOf([[
      cell({ textBody: textBody('A') }),
      cell({ textBody: textBody('B') }),
    ]], [COL, COL]);
    table.rotation = 90;
    table.flipH = true;
    table.rows[0].height = 30 * EMU;
    table.height = 30 * EMU;

    expect(collectRuns(table).map((run) => ({
      text: run.text,
      shapeX: run.shapeX,
      shapeY: run.shapeY,
      rotation: run.rotation,
      shapeFlipH: run.shapeFlipH,
    }))).toEqual([
      { text: 'A', shapeX: 30, shapeY: 30, rotation: 90, shapeFlipH: true },
      { text: 'B', shapeX: 30, shapeY: -30, rotation: 90, shapeFlipH: true },
    ]);
  });

  it('maps cell frames through a vertical table flip', () => {
    const table = tableOf([
      [cell({ textBody: textBody('Top') })],
      [cell({ textBody: textBody('Bottom') })],
    ], [COL]);
    table.flipV = true;

    expect(collectRuns(table).map((run) => ({
      text: run.text,
      shapeY: run.shapeY,
      shapeFlipV: run.shapeFlipV,
      tableCell: run.tableCell,
    }))).toEqual([
      { text: 'Top', shapeY: 20, shapeFlipV: true, tableCell: { row: 0, column: 0 } },
      { text: 'Bottom', shapeY: 0, shapeFlipV: true, tableCell: { row: 1, column: 0 } },
    ]);
  });

  it('maps cell frames through an arbitrary authored rotation', () => {
    const table = tableOf([[
      cell({ textBody: textBody('A') }),
      cell({ textBody: textBody('B') }),
    ]], [COL, COL]);
    table.rotation = 30;
    table.rows[0].height = 30 * EMU;
    table.height = 30 * EMU;

    const runs = collectRuns(table);

    // Cell centres are ±30px from the table centre. Applying the authored 30°
    // frame rotation gives offsets ±(30*cos(30°), 30*sin(30°)).
    expect(runs[0].shapeX).toBeCloseTo(4.0192, 4);
    expect(runs[0].shapeY).toBeCloseTo(-15, 4);
    expect(runs[1].shapeX).toBeCloseTo(55.9808, 4);
    expect(runs[1].shapeY).toBeCloseTo(15, 4);
    expect(runs.map((run) => run.rotation)).toEqual([30, 30]);
  });

  it('carries table element identity through the slide render pipeline', async () => {
    const table = tableOf([[cell({ textBody: textBody('Selectable') })]], [COL]);
    table.id = '31';
    const context = recordingContext();
    const canvas = {
      width: 0,
      height: 0,
      getContext: () => context,
    } as unknown as OffscreenCanvas;
    const slide: Slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [table],
      elementSources: [{ origin: 'layout' }],
    };
    const runs: PptxTextRunInfo[] = [];

    await renderSlide(canvas, slide, table.width, table.height, { width: 60, dpr: 1 }, (run) => {
      runs.push(run);
    });

    expect(runs).not.toHaveLength(0);
    expect(runs.every((run) =>
      run.shapeId === '31' && run.elementIndex === 0 && run.origin === 'layout')).toBe(true);
  });
});
