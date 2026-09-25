import { describe, expect, it } from 'vitest';
import { renderWorksheetViewport } from './render-orchestrator.js';
import type { ParsedWorkbook, Styles, Worksheet } from './types.js';

const STYLES: Styles = {
  fonts: [{
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    size: 11,
    color: null,
    name: null,
  }],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

function context(): { ctx: CanvasRenderingContext2D; texts: string[] } {
  const texts: string[] = [];
  const state: Record<string, unknown> = {
    canvas: { width: 320, height: 180 },
    fillStyle: '#000000',
    font: '11px sans-serif',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    measureText: (text: string) => ({ width: [...text].length * 7 }),
    fillText: (text: string) => texts.push(text),
  };
  const noop = () => {};
  const ctx = new Proxy(state, {
    get(target, property) {
      return property in target ? target[property as string] : noop;
    },
    set(target, property, value) {
      target[property as string] = value;
      return true;
    },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts };
}

function dialogSheet(): Worksheet {
  return {
    name: 'Dialog',
    isDialogSheet: true,
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 0,
    defaultRowHeight: 0,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
  } as Worksheet;
}

describe('XLSX dialog-sheet rendering', () => {
  it('shows a concise non-error notice without exposing parser diagnostics', async () => {
    const recording = context();
    const target = {
      width: 320,
      height: 180,
      getContext: () => recording.ctx,
    } as unknown as OffscreenCanvas;

    await renderWorksheetViewport(
      { ws: dialogSheet(), styles: STYLES as ParsedWorkbook['styles'] },
      target,
      { row: 1, col: 1, rows: 1, cols: 1 },
    );

    expect(recording.texts).toEqual(['Legacy dialog sheets are not displayed']);
    expect(recording.texts.join(' ')).not.toMatch(/xl\/dialogsheets|MCE|could not be displayed/i);
  });
});
