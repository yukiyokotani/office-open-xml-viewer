import { describe, expect, it } from 'vitest';
import { breakDrawingMlText, type DrawingMlInputRun } from './index.js';

// Synthetic advances keep the test independent of the host Canvas fonts. The
// assertions are the line choices in the matched Office wrap controls (C00–C11).
const measure = (value: string): number =>
  [...value].reduce((width, ch) => width + (/\p{Script=Han}/u.test(ch) || ch === '、' ? 20 : 10), 0);

function lines(parts: readonly string[], width: number, defaultTabSize = 72): string[] {
  const runs: DrawingMlInputRun<string>[] = parts.map((text) => ({ type: 'text', text, style: 'same' }));
  return breakDrawingMlText(runs, {
    maxWidth: width,
    measureText: measure,
    sameStyle: (a, b) => a === b,
    defaultTabSize,
  }).map((line) => line.segments.map((segment) => segment.type === 'text'
    ? segment.text : segment.type === 'tab' ? '\t' : '').join('').replace(/ +$/u, ''));
}

describe('matched PowerPoint and Excel DrawingML wrap controls', () => {
  it('wraps a Latin word at the same grapheme with or without a run seam (C00/C01)', () => {
    expect(lines(['abcdef'], 45)).toEqual(['abcd', 'ef']);
    expect(lines(['abc', 'def'], 45)).toEqual(['abcd', 'ef']);
  });

  it('uses the space, hyphen, and CJK/Latin opportunities (C02–C04/C07)', () => {
    expect(lines(['abc def'], 50)).toEqual(['abc', 'def']);
    expect(lines(['日本語Power'], 70)).toEqual(['日本語', 'Power']);
    expect(lines(['non-managed'], 70)).toEqual(['non-', 'managed']);
    expect(lines(['abc ', '$', '100'], 55)).toEqual(['abc', '$100']);
  });

  it('does not reserve a visual continuation for terminal ordinary spaces (C05/C06)', () => {
    expect(lines(['abc  '], 35)).toEqual(['abc']);
    expect(lines(['abc', '  '], 35)).toEqual(['abc']);
  });

  it('retains the advance of spaces before an authored line break for alignment', () => {
    const result = breakDrawingMlText<string>([
      { type: 'text', text: 'Analyze & ', style: 'same' },
      { type: 'break' },
      { type: 'text', text: 'Control', style: 'same' },
    ], { maxWidth: 200, measureText: measure });
    expect(result.map((line) => ({
      text: line.segments.map((segment) => segment.type === 'text' ? segment.text : '').join(''),
      width: line.width,
    }))).toEqual([
      { text: 'Analyze & ', width: 100 },
      { text: 'Control', width: 70 },
    ]);
  });

  it('keeps the observed authored punctuation seam and NBSP behavior (C08/C10)', () => {
    expect(lines(['日本語', '、次'], 60)).toEqual(['日本語', '、次']);
    expect(lines(['abc\u00a0def'], 50)).toEqual(['abc\u00a0d', 'ef']);
  });

  it('breaks an overwide word at grapheme boundaries (C09)', () => {
    expect(lines(['supercalifragilistic'], 80)).toEqual(['supercal', 'ifragili', 'stic']);
  });

  it('carries an overflowing tab to the next line and seats one glyph after its stop (C11)', () => {
    expect(lines(['abc\tdef'], 85)).toEqual(['abc', '\td', 'ef']);
  });

  it('puts a display equation on its own line between text runs', () => {
    const result = breakDrawingMlText<string>([
      { type: 'text', text: 'before', style: 'same' },
      { type: 'object', width: 30, style: 'same', payload: 'equation', display: true },
      { type: 'text', text: 'after', style: 'same' },
    ], { maxWidth: 200, measureText: measure });
    expect(result.map((line) => line.segments.map((segment) =>
      segment.type === 'text' ? segment.text : segment.type === 'object' ? '[equation]' : '\t').join('')))
      .toEqual(['before', '[equation]', 'after']);
  });

  it('bounds work for a long unbreakable word in a one-character box', () => {
    let measuredCharacters = 0;
    const result = breakDrawingMlText([{ type: 'text', text: 'a'.repeat(3200), style: 'same' }], {
      maxWidth: 1,
      measureText(value) { measuredCharacters += value.length; return value.length; },
    });
    expect(result).toHaveLength(3200);
    // Deterministic work bound instead of wall-clock time (#2039): a linear
    // breaker measures ~4n characters here; re-measuring every prefix is
    // ~n²/2 ≈ 5.1M.
    expect(measuredCharacters).toBeLessThan(30_000);
  });

  it('bounds measurement for negative tracking in a one-character box', () => {
    let measuredCharacters = 0;
    const result = breakDrawingMlText([{ type: 'text', text: 'a'.repeat(3200), style: 'same' }], {
      maxWidth: 1,
      nonMonotoneMeasure: true,
      measureText(value) { measuredCharacters += value.length; return value.length; },
    });
    expect(result).toHaveLength(3200);
    // The bounded shaping window measures ~63n characters; re-measuring every
    // prefix is ~n²/2 ≈ 5.1M.
    expect(measuredCharacters).toBeLessThan(1_000_000);
  });

  it('bounds tab resolution for a densely tabbed negative-tracking paragraph', () => {
    let measuredCharacters = 0;
    let tabResolutions = 0;
    const result = breakDrawingMlText([{ type: 'text', text: 'ab\t'.repeat(1067), style: 'same' }], {
      maxWidth: 9,
      defaultTabSize: 72,
      nonMonotoneMeasure: true,
      tabStartPen() { tabResolutions++; return 0; },
      measureText(value) {
        measuredCharacters += value.length;
        const glyphs = [...value].length;
        return glyphs * 9 - 1.5 * Math.max(0, glyphs - 1);
      },
    });
    expect(result.length).toBeGreaterThan(1000);
    // One pen per line for the fit and one for the closed line's paint width;
    // re-resolving every candidate prefix would call this per candidate.
    expect(tabResolutions).toBeLessThanOrEqual(2 * result.length);
    // ~2n characters for n = 3201; re-measuring every prefix is ~n²/2 ≈ 5.1M.
    expect(measuredCharacters).toBeLessThan(100_000);
  });

  it('keeps the last fitting prefix past a shaping window under negative tracking', () => {
    // 'a' advances 9px and 'b' 30px with -10px tracking: each 'a' after the
    // first narrows the line, so the 20th-glyph prefix fits after the first
    // glyph alone overflows, and the final 'b' does not.
    const text = `${'a'.repeat(20)}b`;
    const result = breakDrawingMlText([{ type: 'text', text, style: 'same' }], {
      maxWidth: 8,
      nonMonotoneMeasure: true,
      measureText(value) {
        let width = 0;
        for (const ch of value) width += ch === 'b' ? 30 : 9;
        return width - 10 * Math.max(0, value.length - 1);
      },
    });
    expect(result.map((line) => line.segments.map((segment) => segment.type === 'text' ? segment.text : '').join('')))
      .toEqual(['a'.repeat(20), 'b']);
  });

  it('keeps the last fitting prefix when negative tracking makes widths non-monotone', () => {
    const result = breakDrawingMlText([{ type: 'text', text: 'abcd', style: 'same' }], {
      maxWidth: 1,
      nonMonotoneMeasure: true,
      measureText(value) { return [0, 1, 3, 0.5, 5][value.length]; },
    });
    expect(result.map((line) => line.segments.map((segment) => segment.type === 'text' ? segment.text : '').join('')))
      .toEqual(['abc', 'd']);
  });

});
