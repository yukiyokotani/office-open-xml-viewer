import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Cell, Styles, Worksheet } from './types.js';

const OVERFLOW_TEXT = 'Country Adaptation and Resilience (A&R) Readiness Assessment';

const STYLES: Styles = {
  fonts: [{
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    size: 11,
    color: null,
    name: 'Calibri',
  }],
  fills: [],
  borders: [],
  cellXfs: [{
    fontId: 0,
    fillId: 0,
    borderId: 0,
    numFmtId: 0,
    alignH: 'left',
    wrapText: false,
  } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

function textCell(col: number, text: string): Cell {
  return {
    row: 1,
    col,
    styleIndex: 0,
    value: { type: 'text', text },
  } as Cell;
}

function worksheet(
  cells: Cell[],
  mergeCells: Worksheet['mergeCells'] = [],
): Worksheet {
  return {
    name: 'Intro',
    rows: [{ index: 1, height: null, cells }],
    colWidths: { 1: 3, 2: 8.43, 3: 8.43, 4: 8.43 },
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells,
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
  } as Worksheet;
}

function recordingContext(): {
  ctx: CanvasRenderingContext2D;
  texts: string[];
} {
  const texts: string[] = [];
  const state: Record<string, unknown> = {
    canvas: { width: 240, height: 90 },
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    font: '11px Calibri',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    globalAlpha: 1,
    letterSpacing: '0px',
    measureText: (text: string) => ({ width: [...text].length * 7 }),
    fillText: (text: string) => { texts.push(text); },
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  const noOp = () => {};
  const ctx = new Proxy(state, {
    get(target, property) {
      return property in target ? target[property as string] : noOp;
    },
    set(target, property, value) {
      target[property as string] = value;
      return true;
    },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

describe('virtualized cell text overflow', () => {
  it('keeps an off-screen anchor visible while its text spills into a visible empty cell', () => {
    const recording = recordingContext();

    // B1 itself is outside the viewport, but its measured text extends through
    // the empty visible C1 cell. Excel keeps the text painted until that spill
    // range leaves the viewport as well.
    renderViewport(
      recording.ctx,
      worksheet([textCell(2, OVERFLOW_TEXT)]),
      STYLES,
      { row: 1, col: 3, rows: 1, cols: 2 },
    );

    expect(recording.texts).toContain(OVERFLOW_TEXT);
  });

  it('does not paint the off-screen value through a non-empty visible cell', () => {
    const recording = recordingContext();

    renderViewport(
      recording.ctx,
      worksheet([textCell(2, OVERFLOW_TEXT), textCell(3, 'blocker')]),
      STYLES,
      { row: 1, col: 3, rows: 1, cols: 2 },
    );

    expect(recording.texts).not.toContain(OVERFLOW_TEXT);
  });

  it('keeps an off-screen merged anchor visible while the merged range intersects the viewport', () => {
    const recording = recordingContext();

    // The value is anchored in B, while the authored merge extends through D.
    // Culling against B alone must not remove the merged cell when only D is
    // visible (the shape used by the long introductory paragraph).
    renderViewport(
      recording.ctx,
      worksheet(
        [textCell(2, OVERFLOW_TEXT)],
        [{ top: 1, left: 2, bottom: 1, right: 4 }],
      ),
      STYLES,
      // Keep B in the virtualized band list while scrolling its own rectangle
      // fully behind the row header. This is the boundary where the merged
      // pre-pass used to skip B (because its band was present) and the main
      // pass then culled it using only B's width.
      { row: 1, col: 2, rows: 1, cols: 3 },
      { scrollOffsetX: 70 },
    );

    expect(recording.texts).toContain(OVERFLOW_TEXT);
  });

  it('uses the full merged height when its anchor row is clipped', () => {
    const recording = recordingContext();

    renderViewport(
      recording.ctx,
      worksheet(
        [textCell(2, OVERFLOW_TEXT)],
        [{ top: 1, left: 2, bottom: 3, right: 2 }],
      ),
      STYLES,
      { row: 1, col: 2, rows: 3, cols: 1 },
      { scrollOffsetY: 20 },
    );

    expect(recording.texts).toContain(OVERFLOW_TEXT);
  });

  it('spills a General boolean both ways, as paint centres it', () => {
    const recording = recordingContext();
    // ECMA-376 §18.18.40 General centres booleans. "TRUE" overflows its
    // narrow C column by the same amount on each side, so the half that
    // spills left into the visible B keeps the off-screen anchor painted;
    // a left-anchored scan would have culled it.
    const general: Styles = {
      ...STYLES,
      cellXfs: [{ ...STYLES.cellXfs[0], alignH: null }],
    };
    const ws = worksheet([
      { row: 1, col: 3, styleIndex: 0, value: { type: 'bool', bool: true } } as Cell,
    ]);
    ws.colWidths = { 1: 8.43, 2: 8.43, 3: 1 };
    renderViewport(recording.ctx, ws, general, { row: 1, col: 1, rows: 1, cols: 2 });

    expect(recording.texts).toContain('TRUE');
  });
});
