import { describe, expect, it } from 'vitest';
import { breakDrawingMlText, drawingMlSegmentSourceRanges, type DrawingMlInputRun } from './index.js';

// Synthetic advances keep the test independent of the host Canvas fonts. The
// assertions are the line choices in the matched Office wrap controls (C00–C11).
const measure = (value: string): number =>
  [...value].reduce((width, ch) => width + (/\p{Script=Han}/u.test(ch) || ch === '、' ? 20 : 10), 0);

it('keeps nonpaint source metadata line-local without changing contextual measurement', () => {
  const runs = ['A', 'B'].map((text, metric) => ({ type: 'text' as const, text, style: { font: 'same', metric } }));
  const result = breakDrawingMlText(runs, { maxWidth: 12,
    measureText: text => text === 'AB' ? 16 : 5, sameStyle: (a, b) => a.font === b.font });
  expect(result.map(line => line.segments.map(seg => seg.type === 'text' ? seg.text : ''))).toEqual([['A'], ['B']]);
  expect(result.map(line => line.segments[0].style.metric)).toEqual([0, 1]);
  for (const line of result) for (const segment of line.segments) {
    if (segment.type !== 'text') continue;
    expect(drawingMlSegmentSourceRanges(segment).map(range => runs[range.run].text.slice(range.start, range.end)).join('')).toBe(segment.text);
  }
  const wide = breakDrawingMlText(runs, { maxWidth: 20, measureText: text => text === 'AB' ? 16 : 5,
    sameStyle: (a, b) => a.font === b.font });
  expect(wide[0].width).toBe(16);
  expect(drawingMlSegmentSourceRanges(wide[0].segments[0])).toEqual([{run: 0, start: 0, end: 1}, {run: 1, start: 0, end: 1}]);
});

it('retains exact display slices through grapheme seams, hard breaks and terminal trimming', () => {
  const runs: DrawingMlInputRun<number>[] = [{ type: 'text', text: 'A', style: 0 },
    { type: 'text', text: '\u0301B', style: 1 }, { type: 'break' }, { type: 'text', text: 'C ', style: 2 }];
  const result = breakDrawingMlText(runs, { maxWidth: 10, measureText: text => [...text.replace('\u0301', '')].length * 10,
    sameStyle: () => true });
  expect(result.map(line => line.segments.map(seg => seg.type === 'text' ? seg.text : '').join(''))).toEqual(['Á', 'B', 'C']);
  for (const line of result) for (const segment of line.segments) {
    if (segment.type !== 'text') continue;
    const slices = drawingMlSegmentSourceRanges(segment).map(range => {
      const source = runs[range.run]; return source.type === 'text' ? source.text.slice(range.start, range.end) : '';
    });
    expect(slices.join('')).toBe(segment.text);
  }
});

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

  it('keeps grapheme boundaries across paint styles during emergency wrapping', () => {
    for (const nonMonotoneMeasure of [false, true]) {
      for (const [mark, expected] of [['\ua9e5', ['\u1000\ua9e5']], ['\uaa7c', ['\u1000\uaa7c']],
        ['\uaa7b', ['\u1000', '\uaa7b']], ['\uaa7d', ['\u1000', '\uaa7d']]] as const) {
        const result = breakDrawingMlText([
          { type: 'text', text: '\u1000', style: 'base' },
          { type: 'text', text: mark, style: 'mark' },
        ], { maxWidth: 10, nonMonotoneMeasure,
          measureText: (text) => [...text].reduce((sum, ch) => sum + (ch === '\u1000' ? 20 : 0), 0) });
        expect(result.map((line) => line.segments.map((segment) => segment.type === 'text' ? segment.text : '').join('')))
          .toEqual(expected);
        expect(result.flatMap((line) => line.segments).map((segment) => segment.style)).toEqual(['base', 'mark']);
      }
    }
  });

  it('fits complete cross-style graphemes and keeps a space-attached mark when wrapping', () => {
    const wrapped = (base: string, width: number) => breakDrawingMlText([
      { type: 'text', text: `A${base}`, style: 'base' },
      { type: 'text', text: '\u0301B', style: 'mark' },
    ], { maxWidth: width, measureText: (text) => [...text].reduce((sum, ch) => sum + (ch === '\u0301' ? 0 : 20), 0) })
      .map((line) => line.segments.map((segment) => segment.type === 'text' ? segment.text : '').join(''));
    expect(wrapped('', 20)).toEqual(['A\u0301', 'B']);
    expect(wrapped(' ', 20)).toEqual(['A', ' \u0301', 'B']);
  });

  it('recovers grapheme boundaries shifted inside a styled run by regional-indicator pairing', () => {
    const result = breakDrawingMlText([
      { type: 'text', text: '🇦', style: 'first' },
      { type: 'text', text: '🇧🇨🇩', style: 'rest' },
    ], { maxWidth: 40, measureText: (text) => [...text].length * 20 });
    expect(result.map((line) => line.segments.map((segment) => segment.type === 'text' ? segment.text : '').join('')))
      .toEqual(['🇦🇧', '🇨🇩']);
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
    const start = performance.now();
    const result = breakDrawingMlText([{ type: 'text', text: 'a'.repeat(3200), style: 'same' }], {
      maxWidth: 1,
      measureText(value) { measuredCharacters += value.length; return value.length; },
    });
    const elapsedMs = performance.now() - start;
    expect(result).toHaveLength(3200);
    expect(measuredCharacters).toBeLessThan(30_000);
    expect(elapsedMs).toBeLessThan(100);
  });

  it('bounds measurement for negative tracking in a one-character box', () => {
    let measuredCharacters = 0;
    const start = performance.now();
    const result = breakDrawingMlText([{ type: 'text', text: 'a'.repeat(3200), style: 'same' }], {
      maxWidth: 1,
      nonMonotoneMeasure: true,
      measureText(value) { measuredCharacters += value.length; return value.length; },
    });
    const elapsedMs = performance.now() - start;
    expect(result).toHaveLength(3200);
    expect(measuredCharacters).toBeLessThan(1_000_000);
    expect(elapsedMs).toBeLessThan(100);
  });

  it('bounds tab resolution for a densely tabbed negative-tracking paragraph', () => {
    let measuredCharacters = 0;
    let tabResolutions = 0;
    const start = performance.now();
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
    const elapsedMs = performance.now() - start;
    expect(result.length).toBeGreaterThan(1000);
    // One pen per line for the fit and one for the closed line's paint width;
    // re-resolving every candidate prefix would call this per candidate.
    expect(tabResolutions).toBeLessThanOrEqual(2 * result.length);
    expect(measuredCharacters).toBeLessThan(100_000);
    expect(elapsedMs).toBeLessThan(100);
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
