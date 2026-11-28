import { describe, expect, it } from 'vitest';
import { wrapMeasuredText } from './axis.js';

// One unit per code point keeps widths deterministic.
const ctx = {
  measureText: (text: string) => ({ width: Array.from(text).length * 10 }),
} as unknown as CanvasRenderingContext2D;

describe('wrapMeasuredText', () => {
  it('starts a new line at each authored break even when the text would fit', () => {
    expect(wrapMeasuredText(ctx, 'FY2024\n1Q営業利益', 1000)).toEqual(['FY2024', '1Q営業利益']);
    expect(wrapMeasuredText(ctx, 'A\r\nB\rC', 1000)).toEqual(['A', 'B', 'C']);
  });

  it('keeps an interior blank line and trims outer breaks like outer spaces', () => {
    expect(wrapMeasuredText(ctx, '\nA\n\nB\n', 1000)).toEqual(['A', '', 'B']);
  });

  it('still wraps each authored line by width', () => {
    expect(wrapMeasuredText(ctx, 'aa bb\ncc', 30)).toEqual(['aa', 'bb', 'cc']);
  });

  // A grouped number's separator can be a no-break space (fr-FR groups with
  // U+202F; U+00A0 and U+2007 behave alike): never a line-break opportunity.
  it('does not break at, or rewrite, a no-break space', () => {
    for (const space of ['\u202f', '\u00a0', '\u2007']) {
      expect(wrapMeasuredText(ctx, `1${space}234,5`, 1000)).toEqual([`1${space}234,5`]);
      expect(wrapMeasuredText(ctx, `ab 1${space}234`, 50)).toEqual(['ab', `1${space}234`]);
    }
  });

  it('never splits beside a no-break space when a token is split by code point', () => {
    for (const space of ['\u202f', '\u00a0', '\u2007']) {
      const lines = wrapMeasuredText(ctx, `1${space}234,5`, 20);
      expect(lines.join(''), space).toBe(`1${space}234,5`);
      for (const line of lines) {
        expect(line.startsWith(space) || line.endsWith(space), JSON.stringify(lines)).toBe(false);
      }
    }
  });

  it('keeps unbroken text unchanged', () => {
    expect(wrapMeasuredText(ctx, '  North  region ', 1000)).toEqual(['North region']);
    expect(wrapMeasuredText(ctx, '', 1000)).toEqual(['']);
  });
});
