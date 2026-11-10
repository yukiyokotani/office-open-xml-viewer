import { afterEach, describe, expect, it, vi } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

type FaceStub = { family: string; weight: string; style: string; status: string };

/** Paints one General number (123456789) in a merged A1:C1 whose width leaves
 *  room for two characters, so no General form fits (issue #2179). Returns the
 *  texts drawn. The anchor is on screen, or scrolled out of view with the
 *  merge still visible. */
function paint(opts: { icon: boolean; anchorOffscreen: boolean; faces: FaceStub[]; uncovered?: boolean; wrap?: boolean }): string[] {
  vi.stubGlobal('document', { fonts: opts.faces });
  const texts: string[] = [];
  const noop = () => {};
  const ctx = new Proxy({
    canvas: { width: 600, height: 300 },
    // `uncovered`: some glyph falls back, so the serif fallback measures wider.
    measureText(this: { font?: string }, text: string) {
      return { width: text.length * 7 + (opts.uncovered && /serif$/.test(this.font ?? '') ? 1 : 0) };
    },
    fillText: (text: string) => texts.push(text),
  } as Record<string | symbol, unknown>, {
    get: (target, key) => key in target ? target[key] : noop,
    set: (target, key, value) => { target[key] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const styles: Styles = {
    fonts: [{ name: 'Arial', size: 11, bold: false, italic: false, underline: false, strike: false, color: null }],
    fills: [], borders: [], dxfs: [],
    cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, wrapText: opts.wrap === true }] as Styles['cellXfs'],
    numFmts: [],
  };
  const ws = {
    name: 'General',
    rows: [{ index: 1, height: null, cells: [
      { row: 1, col: 1, styleIndex: 0, value: { type: 'number', number: 123456789 } },
    ] }],
    colWidths: { 1: 1, 2: 1, 3: 1 }, rowHeights: {},
    defaultColWidth: 1, defaultRowHeight: 15,
    mergeCells: [{ top: 1, left: 1, bottom: 1, right: 3 }],
    freezeRows: 0, freezeCols: 0, images: [], charts: [],
    conditionalFormats: opts.icon ? [{
      sqref: [{ top: 0, left: 0, bottom: 10, right: 10 }],
      rules: [{
        type: 'iconSet', iconSet: '3Arrows', reverse: false, priority: 1,
        cfvos: [{ kind: 'percent', value: '0' }, { kind: 'percent', value: '33' }, { kind: 'percent', value: '67' }],
      }],
    }] : [],
  } as unknown as Worksheet;
  renderViewport(ctx, ws, styles, opts.anchorOffscreen
    ? { row: 1, col: 2, rows: 1, cols: 2 }
    : { row: 1, col: 1, rows: 1, cols: 3 });
  return texts;
}

const LOADED_ARIAL: FaceStub = { family: 'Arial', weight: '400', style: 'normal', status: 'loaded' };
const hashes = (texts: string[]) => texts.filter((text) => /^#+$/.test(text));

afterEach(() => { vi.unstubAllGlobals(); });

describe('General numbers too wide for a merged cell (issue #2179)', () => {
  it.each([false, true])('fills "#" when the authored face is loaded (anchor offscreen: %s)', (anchorOffscreen) => {
    const texts = paint({ icon: false, anchorOffscreen, faces: [LOADED_ARIAL] });
    expect(hashes(texts)).toHaveLength(1);
    expect(texts).not.toContain('1E+08');
  });

  it.each([false, true])('keeps the shortest form beside an icon-set icon (anchor offscreen: %s)', (anchorOffscreen) => {
    const texts = paint({ icon: true, anchorOffscreen, faces: [LOADED_ARIAL] });
    expect(hashes(texts)).toHaveLength(0);
    expect(texts).toContain('1E+08');
  });

  it('leaves a wrapped cell to the wrap painter, as its automatic row height was measured', () => {
    for (const anchorOffscreen of [false, true]) {
      const texts = paint({ icon: false, anchorOffscreen, faces: [LOADED_ARIAL], wrap: true });
      expect(hashes(texts)).toHaveLength(0);
      expect(texts).not.toContain('1E+08');
      expect(texts.join('')).toContain('123456789');
    }
  });

  it('keeps the shortest form when the loaded face does not cover the text', () => {
    for (const anchorOffscreen of [false, true]) {
      const texts = paint({ icon: false, anchorOffscreen, faces: [LOADED_ARIAL], uncovered: true });
      expect(hashes(texts)).toHaveLength(0);
      expect(texts).toContain('1E+08');
    }
  });

  it.each([
    ['unloaded', { ...LOADED_ARIAL, status: 'unloaded' }],
    ['failed', { ...LOADED_ARIAL, status: 'error' }],
    ['bold-only', { ...LOADED_ARIAL, weight: '700' }],
    ['italic-only', { ...LOADED_ARIAL, style: 'italic' }],
    ['another family', { ...LOADED_ARIAL, family: 'Calibri' }],
  ] as const)('keeps the shortest form when the face is %s', (_, face) => {
    for (const anchorOffscreen of [false, true]) {
      const texts = paint({ icon: false, anchorOffscreen, faces: [face] });
      expect(hashes(texts)).toHaveLength(0);
      expect(texts).toContain('1E+08');
    }
  });
});
