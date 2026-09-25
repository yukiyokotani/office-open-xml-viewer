import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types.js';

function mockCtx(): CanvasRenderingContext2D {
  let font = '';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: [...text].length * 10 }) as TextMetrics,
    fillRect() {},
    fillText() {},
    fillStyle: '',
    strokeStyle: '',
  } as unknown as CanvasRenderingContext2D;
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', fontFamily: 'Arial',
  };
}

function paragraph(text: string, runs: TextRunData[] = [run(text)]): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [],
    eaLnBrk: true, runs,
  } as Paragraph;
}

function lineTexts(lines: ReturnType<typeof layoutParagraph>): string[] {
  return lines.map((line) => line.segments.map((segment) => segment.text).join(''));
}

describe('PPTX Latin compound hyphen wrapping', () => {
  it('uses an authored hyphen as a soft break opportunity', () => {
    // UAX #14 class HY provides a break opportunity AFTER the hyphen. With a
    // 140px line, the first compound fragment fits the preceding words exactly:
    // "aaaa bbbb non-" = 14 cells. Keeping the compound indivisible creates a
    // third line even though PowerPoint fits the same text in two.
    expect(lineTexts(layoutParagraph(
      mockCtx(), paragraph('aaaa bbbb non-managed bbbb'),
      140, 20, '000000', 1, 0,
    ))).toEqual(['aaaa bbbb non-', 'managed bbbb']);
  });

  it('uses U+2010 HYPHEN as the same explicit compound boundary', () => {
    expect(lineTexts(layoutParagraph(
      mockCtx(), paragraph('aaaa bbbb non‐managed bbbb'),
      140, 20, '000000', 1, 0,
    ))).toEqual(['aaaa bbbb non‐', 'managed bbbb']);
  });

  it('keeps option prefixes intact while using an internal compound hyphen', () => {
    expect(lineTexts(layoutParagraph(
      mockCtx(), paragraph('aaaa --list-masters bbbb'),
      120, 20, '000000', 1, 0,
    ))).toEqual(['aaaa --list-', 'masters bbbb']);
  });

  it('keeps the break opportunity across a formatting-run boundary', () => {
    expect(lineTexts(layoutParagraph(
      mockCtx(), paragraph('', [run('aaaaaaaa-'), run('bbbb')]),
      100, 20, '000000', 1, 0,
    ))).toEqual(['aaaaaaaa-', 'bbbb']);
  });

  it('does not split a numeric range, word-initial hyphen, or Hebrew compound', () => {
    const numeric = lineTexts(layoutParagraph(
      mockCtx(), paragraph('aaaa bbbb 2024-2025 bbbb'),
      140, 20, '000000', 1, 0,
    ));
    expect(numeric).not.toContain('aaaa bbbb 2024-');

    const wordInitial = lineTexts(layoutParagraph(
      mockCtx(), paragraph('aaaa bbbb -managed bbbb'),
      140, 20, '000000', 1, 0,
    ));
    expect(wordInitial).not.toContain('aaaa bbbb -');

    // Non-Latin scripts remain outside this deliberately bounded path (and in
    // particular must not override UAX #14 LB21a for Hebrew text).
    const hebrew = lineTexts(layoutParagraph(
      mockCtx(), paragraph('aaaa bbbb אב-גד bbbb'),
      140, 20, '000000', 1, 0,
    ));
    expect(hebrew).not.toContain('aaaa bbbb אב-');
  });
});
