import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

/** Paints one number in a merged A1:C1 of `charWidth`-character columns and
 *  returns the texts drawn. The anchor is on screen, or scrolled out of view
 *  with the merge still visible (issue #2179). */
function paint(opts: {
  value: number; anchorOffscreen: boolean; charWidth: number;
  icon?: boolean; wrap?: boolean; formatCode?: string;
}): string[] {
  const texts: string[] = [];
  const noop = () => {};
  const ctx = new Proxy({
    canvas: { width: 600, height: 300 },
    measureText: (text: string) => ({ width: text.length * 7 }),
    fillText: (text: string) => texts.push(text),
  } as Record<string | symbol, unknown>, {
    get: (target, key) => key in target ? target[key] : noop,
    set: (target, key, value) => { target[key] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const styles: Styles = {
    fonts: [{ name: 'Arial', size: 11, bold: false, italic: false, underline: false, strike: false, color: null }],
    fills: [], borders: [], dxfs: [],
    cellXfs: [{
      fontId: 0, fillId: 0, borderId: 0, numFmtId: opts.formatCode ? 164 : 0, wrapText: opts.wrap === true,
    }] as Styles['cellXfs'],
    numFmts: opts.formatCode ? [{ numFmtId: 164, formatCode: opts.formatCode }] : [],
  };
  const w = opts.charWidth;
  const ws = {
    name: 'General',
    rows: [{ index: 1, height: null, cells: [
      { row: 1, col: 1, styleIndex: 0, value: { type: 'number', number: opts.value } },
    ] }],
    colWidths: { 1: w, 2: w, 3: w }, rowHeights: {},
    defaultColWidth: w, defaultRowHeight: 15,
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

const hashes = (texts: string[]) => texts.filter((text) => /^#+$/.test(text));
const both = [false, true];

describe('General numbers in a merged cell (issue #2179)', () => {
  it.each(both)('shortens to the merged width (anchor offscreen: %s)', (anchorOffscreen) => {
    // Three 3-character columns leave room for nine 7 px characters.
    const narrow = paint({ value: Math.PI, anchorOffscreen, charWidth: 3 });
    expect(narrow).toContain('3.1415927');
    expect(narrow).not.toContain('3.141592654');
  });

  it('shortens identically whether the anchor is on screen or not, icon included', () => {
    for (const icon of [false, true]) {
      const shown = (anchorOffscreen: boolean) =>
        paint({ value: Math.PI, anchorOffscreen, charWidth: 3, icon }).filter((t) => t.startsWith('3'));
      expect(shown(true)).toEqual(shown(false));
    }
    // The icon reserves width, so the icon cell shows a shorter form.
    const plain = paint({ value: Math.PI, anchorOffscreen: true, charWidth: 3 }).find((t) => t.startsWith('3'));
    const iconed = paint({ value: Math.PI, anchorOffscreen: true, charWidth: 3, icon: true }).find((t) => t.startsWith('3'));
    expect((iconed ?? '').length).toBeLessThan((plain ?? '').length);
  });

  it.each(both)('keeps the shortest form instead of a "#" fill (anchor offscreen: %s)', (anchorOffscreen) => {
    const texts = paint({ value: 123456789, anchorOffscreen, charWidth: 1 });
    expect(hashes(texts)).toHaveLength(0);
    expect(texts).toContain('1E+08');
  });

  it.each(both)('leaves an explicit format unchanged when too wide (anchor offscreen: %s)', (anchorOffscreen) => {
    const texts = paint({ value: 123456.789, anchorOffscreen, charWidth: 1, formatCode: '0.000' });
    expect(hashes(texts)).toHaveLength(0);
    expect(texts).toContain('123456.789');
  });

  it.each(both)('leaves a wrapped cell to the wrap painter (anchor offscreen: %s)', (anchorOffscreen) => {
    const texts = paint({ value: 123456789, anchorOffscreen, charWidth: 1, wrap: true });
    expect(hashes(texts)).toHaveLength(0);
    expect(texts).not.toContain('1E+08');
    expect(texts.join('')).toContain('123456789');
  });
});
