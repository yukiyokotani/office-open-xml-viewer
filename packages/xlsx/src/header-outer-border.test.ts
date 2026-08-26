import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

const STYLES: Styles = {
  fonts: [{ bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null }],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

function worksheet(rightToLeft: boolean): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
    rightToLeft,
  } as Worksheet;
}

interface Segment { x1: number; y1: number; x2: number; y2: number; stroke: string }
interface Fill { x: number; y: number; w: number; h: number; fill: string }
interface TextPaint { text: string; fill: string }

function recordingCtx(width = 300, height = 120): {
  ctx: CanvasRenderingContext2D;
  segments: Segment[];
  fills: Fill[];
  texts: TextPaint[];
} {
  const segments: Segment[] = [];
  const fills: Fill[] = [];
  const texts: TextPaint[] = [];
  let strokeStyle = '#000';
  let fillStyle = '#000';
  let cursor: [number, number] | null = null;
  const ctx: Record<string, unknown> = {
    canvas: { width, height },
    font: '11px sans-serif',
    get fillStyle() { return fillStyle; },
    set fillStyle(value: string) { fillStyle = value; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(value: string) { strokeStyle = value; },
    lineWidth: 1,
    textBaseline: 'alphabetic',
    textAlign: 'left',
    letterSpacing: '0px',
    direction: 'ltr',
    globalAlpha: 1,
    measureText: (text: string) => ({ width: text.length * 8 }),
    fillText: (text: string) => { texts.push({ text, fill: fillStyle }); },
    strokeText: () => {},
    fillRect: (x: number, y: number, w: number, h: number) => {
      fills.push({ x, y, w, h, fill: fillStyle });
    },
    strokeRect: () => {}, clearRect: () => {},
    beginPath: () => { cursor = null; }, closePath: () => {},
    moveTo: (x: number, y: number) => { cursor = [x, y]; },
    lineTo: (x: number, y: number) => {
      if (cursor) segments.push({ x1: cursor[0], y1: cursor[1], x2: x, y2: y, stroke: strokeStyle });
      cursor = [x, y];
    },
    rect: () => {}, arc: () => {}, fill: () => {}, stroke: () => {}, clip: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, setLineDash: () => {}, setTransform: () => {},
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, segments, fills, texts };
}

describe('XLSX header frame ownership', () => {
  it('does not paint sheet gridlines over a cell background fill', () => {
    const ws = worksheet(false);
    ws.rows = [{
      index: 1,
      height: null,
      cells: [{ col: 1, row: 1, styleIndex: 1, value: { type: 'empty' } }],
    }];
    const styles: Styles = {
      ...STYLES,
      fills: [
        { patternType: 'none', fgColor: null, bgColor: null },
        { patternType: 'solid', fgColor: 'FFFFFF', bgColor: null },
      ],
      cellXfs: [
        STYLES.cellXfs[0],
        { ...STYLES.cellXfs[0], fillId: 1 },
      ],
    };
    const { ctx, segments } = recordingCtx();

    renderViewport(ctx, ws, styles, { row: 1, col: 1, rows: 1, cols: 1 });

    expect(segments.filter(segment => segment.stroke === '#d0d0d0')).toEqual([]);
  });

  it('inherits a column background style for otherwise empty cells', () => {
    const ws = worksheet(false);
    ws.colStyleRanges = [{ min: 1, max: 16_384, styleIndex: 1 }];
    const styles: Styles = {
      ...STYLES,
      fills: [
        { patternType: 'none', fgColor: null, bgColor: null },
        { patternType: 'solid', fgColor: 'FFFFFF', bgColor: null },
      ],
      cellXfs: [
        STYLES.cellXfs[0],
        { ...STYLES.cellXfs[0], fillId: 1 },
      ],
    };
    const { ctx, segments } = recordingCtx();

    renderViewport(ctx, ws, styles, { row: 1, col: 1, rows: 1, cols: 1 });

    expect(segments.filter(segment => segment.stroke === '#d0d0d0')).toEqual([]);
  });

  it('bounds frozen-band materialization to the visible canvas', () => {
    const ws = worksheet(false);
    let rowReads = 0;
    let colReads = 0;
    ws.rowHeights = new Proxy({}, {
      get: (target, key, receiver) => {
        if (typeof key === 'string' && /^\d+$/.test(key)) rowReads++;
        return Reflect.get(target, key, receiver);
      },
    });
    ws.colWidths = new Proxy({}, {
      get: (target, key, receiver) => {
        if (typeof key === 'string' && /^\d+$/.test(key)) colReads++;
        return Reflect.get(target, key, receiver);
      },
    });
    const { ctx } = recordingCtx();

    renderViewport(ctx, ws, { ...STYLES }, { row: 1, col: 1, rows: 2, cols: 2 }, {
      freezeRows: 4_294_967_295,
      freezeCols: 4_294_967_295,
    });

    expect(rowReads).toBeLessThan(100);
    expect(colReads).toBeLessThan(100);
  });

  it.each([
    { direction: 'LTR', rtl: false, outerX: 0.5, dividerX: 49.5 },
    { direction: 'RTL', rtl: true, outerX: 299.5, dividerX: 250.5 },
  ])('leaves the $direction outer frame to the host container', ({ rtl, outerX, dividerX }) => {
    const { ctx, segments } = recordingCtx();
    renderViewport(ctx, worksheet(rtl), STYLES, { row: 1, col: 1, rows: 2, cols: 2 });

    const headerSegments = segments.filter(({ stroke }) => stroke === '#c8ccd0');
    expect(headerSegments.some(({ x1, x2 }) => x1 === outerX && x2 === outerX)).toBe(false);
    expect(headerSegments.some(({ y1, y2 }) => y1 === 0.5 && y2 === 0.5)).toBe(false);
    // The row-header/data divider remains part of the spreadsheet grid.
    expect(headerSegments.some(({ x1, x2 }) => x1 === dividerX && x2 === dividerX)).toBe(true);
  });

  it.each([
    { direction: 'LTR', rtl: false },
    { direction: 'RTL', rtl: true },
  ])('draws the $direction frozen-row separator through the row-number header', ({ rtl }) => {
    const ws = worksheet(rtl);
    ws.freezeRows = 1;
    const { ctx, segments } = recordingCtx();
    renderViewport(ctx, ws, STYLES, { row: 2, col: 1, rows: 2, cols: 2 }, { freezeRows: 1 });

    expect(segments.some(segment =>
      segment.stroke === '#7a7a7a' && segment.x1 === 0 && segment.x2 === 300,
    )).toBe(true);
  });

  it('themes only Viewer-owned row and column header chrome', () => {
    const { ctx, segments, fills, texts } = recordingCtx();
    renderViewport(ctx, worksheet(false), STYLES, { row: 1, col: 1, rows: 2, cols: 2 }, {
      chromeColors: {
        surface: '#101820',
        mutedSurface: '#182430',
        text: '#f5f7fa',
        border: '#52606d',
        selectedSurface: '#203a56',
        accent: '#62a8e5',
      },
    });

    expect(fills.some(({ fill }) => fill === '#101820')).toBe(true);
    expect(segments.some(({ stroke }) => stroke === '#52606d')).toBe(true);
    expect(texts.some(({ fill }) => fill === '#f5f7fa')).toBe(true);
    // Authored worksheet/grid paint keeps its own palette.
    expect(segments.some(({ stroke }) => stroke === '#d0d0d0')).toBe(true);
  });
});
