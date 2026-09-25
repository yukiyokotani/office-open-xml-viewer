import { describe, expect, it } from 'vitest';
import { drawWrappedPlainText } from './renderer.js';
import type { CellFont } from './types.js';

const font: CellFont = { name: 'Calibri', size: 12, bold: true, italic: false,
  underline: false, strike: false, color: null };

describe('plain-cell wrapping with a wider substitute face', () => {
  it('uses the painted Canvas face for both line breaks and glyph widths', () => {
    const text = 'Alpha Beta Gamma';
    const calls: { text: string; y: number; maxWidth?: number }[] = [];
    const ctx = {
      textBaseline: 'alphabetic',
      measureText: (value: string) => ({ width: value.length * 12 }),
      fillText: (value: string, _x: number, y: number, maxWidth?: number) => {
        calls.push({ text: value, y, maxWidth });
      },
    } as unknown as CanvasRenderingContext2D;
    const geom = { alignH: 'right', alignV: 'bottom', cx: 0, cy: 22,
      cellW: 96, cellH: 45, leftPad: 3, paddingX: 3, paddingY: 2 };

    drawWrappedPlainText(ctx, text, geom.cellW - 3, font, geom, 1);

    // Every word is laid out even if the fixed row is too short for the
    // substitute. The caller's cell clip, not a narrower unrelated font
    // profile, determines what remains visible inside the fixed row.
    expect(calls.map(({ text: line }) => line.trim())).toEqual(['Alpha', 'Beta', 'Gamma']);
    expect(calls.map(({ maxWidth }) => maxWidth)).toEqual([undefined, undefined, undefined]);
    expect(calls[0].y).toBeLessThan(geom.cy);
  });
});
