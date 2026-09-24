import { describe, expect, it } from 'vitest';
import { drawWrappedPlainText } from './renderer.js';
import { calibriCompatibleBasicLatinWidth, shouldUseCalibriCompatibleWrap } from './fonts/carlito-basic-latin.js';
import type { CellFont } from './types.js';

const font: CellFont = { name: 'Calibri', size: 12, bold: true, italic: false,
  underline: false, strike: false, color: null };
const route = { family: 'Calibri', bold: true, italic: false,
  checkedOfficeTuples: new Set(['calibri:700:normal']), hasExactRoute: false,
  hasDeclaredFace: false, googleSubstitutes: false };

describe('missing Calibri Bold plain-cell wrapping', () => {
  it('keeps a fixed-height heading visible when the authored face is absent', () => {
    const text = 'Alpha Beta Gamma';
    const sizePx = 16;
    const reference = (value: string) => calibriCompatibleBasicLatinWidth(value, sizePx) ?? 0;
    const available = reference('Alpha Beta') + 0.5;
    const calls: { text: string; y: number; maxWidth?: number }[] = [];
    const ctx = {
      textBaseline: 'alphabetic',
      measureText: (value: string) => ({ width: reference(value) * 1.35 }),
      fillText: (value: string, _x: number, y: number, maxWidth?: number) => {
        calls.push({ text: value, y, maxWidth });
      },
    } as unknown as CanvasRenderingContext2D;
    const geom = { alignH: 'right', alignV: 'bottom', cx: 0, cy: 22,
      cellW: available + 6, cellH: 45, leftPad: 3, paddingX: 3, paddingY: 2 };

    drawWrappedPlainText(ctx, text, geom.cellW - 3, font, geom, 1);
    const fallback = [...calls];
    calls.length = 0;
    expect(shouldUseCalibriCompatibleWrap(route, text)).toBe(true);
    drawWrappedPlainText(ctx, text, geom.cellW - 3, font, geom, 1, reference);

    expect(fallback).toHaveLength(3);
    expect(fallback[0].y).toBeLessThan(geom.cy);
    expect(calls.map(({ text }) => text.trim())).toEqual(['Alpha Beta', 'Gamma']);
    expect(calls[0].y).toBeGreaterThan(geom.cy);
    expect(calls.every(({ text, maxWidth }) => maxWidth === reference(text))).toBe(true);

    calls.length = 0;
    drawWrappedPlainText(ctx, text, geom.cellW - 3, font, { ...geom, cellH: 100 }, 1, reference);
    expect(calls).toHaveLength(3);
    expect(calls.every(({ maxWidth }) => maxWidth === undefined)).toBe(true);
  });

  it('does not borrow another tuple’s completed preflight or override usable faces', () => {
    expect(shouldUseCalibriCompatibleWrap({ ...route,
      checkedOfficeTuples: new Set(['calibri']) }, 'Alpha Beta')).toBe(false);
    expect(shouldUseCalibriCompatibleWrap({ ...route, hasExactRoute: true }, 'Alpha Beta')).toBe(false);
    expect(shouldUseCalibriCompatibleWrap({ ...route, hasDeclaredFace: true }, 'Alpha Beta')).toBe(false);
    expect(shouldUseCalibriCompatibleWrap({ ...route, googleSubstitutes: true }, 'Alpha Beta')).toBe(false);
    expect(shouldUseCalibriCompatibleWrap(route, '日本語')).toBe(false);
  });
});
