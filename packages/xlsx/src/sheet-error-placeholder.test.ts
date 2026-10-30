import { describe, expect, it } from 'vitest';
import { renderWorksheetViewport } from './render-orchestrator.js';
import type { ParsedWorkbook, Styles, Worksheet } from './types.js';
import type { WireErrorPlaceholder } from '@silurus/ooxml-core/internal/part-failures';

/** #1974 — the broken-sheet placeholder honours resolved caller text. */
const STYLES = {
  fonts: [{ bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null }],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 }],
  numFmts: [],
  dxfs: [],
} as unknown as Styles;

function recordingTarget(): { target: OffscreenCanvas; texts: string[] } {
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
    get: (target, property) => property in target ? target[property as string] : noop,
    set: (target, property, value) => { target[property as string] = value; return true; },
  });
  const target = { width: 320, height: 180, getContext: () => ctx } as unknown as OffscreenCanvas;
  return { target, texts };
}

function brokenSheet(): Worksheet {
  return {
    name: 'Budget',
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
    parseError: 'xl/worksheets/sheet2.xml: unexpected end of stream',
  } as unknown as Worksheet;
}

async function paint(errorPlaceholderText?: WireErrorPlaceholder): Promise<string[]> {
  const { target, texts } = recordingTarget();
  await renderWorksheetViewport(
    { ws: brokenSheet(), styles: STYLES as ParsedWorkbook['styles'] },
    target,
    { row: 1, col: 1, rows: 1, cols: 1 },
    { width: 320, height: 180, dpr: 1, ...(errorPlaceholderText ? { errorPlaceholderText } : {}) },
  );
  return texts;
}

describe('XLSX sheet placeholder text (#1974)', () => {
  it('keeps the built-in text by default', async () => {
    const texts = await paint();
    expect(texts).toContain('Sheet "Budget" could not be displayed');
    expect(texts.join(' ')).toContain('sheet2.xml');
    expect(await paint({ state: 'builtin' })).toEqual(texts);
  });

  it('paints a custom title and hides the detail for null', async () => {
    const texts = await paint({ state: 'resolved', title: 'Not available', detail: null });
    expect(texts).toContain('Not available');
    expect(texts.join(' ')).not.toContain('sheet2.xml');
  });

  it('paints nothing when the text still has to be resolved on main', async () => {
    await expect(paint({ state: 'resolve-on-main' })).rejects.toThrow(/main thread/);
  });
});
