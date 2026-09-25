import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph } from './renderer.js';
import type { Paragraph } from './types.js';

function mockCtx(): CanvasRenderingContext2D {
  let font = '';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: [...text].length * 10 }),
    fillRect() {},
    fillText() {},
    fillStyle: '',
    strokeStyle: '',
  } as unknown as CanvasRenderingContext2D;
}

function run(text: string, fontFamily = 'Arial'): TextRunData {
  return {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: '000000',
    fontFamily,
    fontFamilyEa: 'Meiryo',
  };
}

function paragraph(runs: TextRunData[]): Paragraph {
  return {
    alignment: 'l',
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' },
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs,
  } as Paragraph;
}

function lines(runs: TextRunData[], width = 70): string[] {
  return layoutParagraph(
    mockCtx(), paragraph(runs), width, 20, '000000', 1, 0,
  ).map((line) => line.segments.map((segment) => segment.text).join(''));
}

describe('pptx CJK/Latin word wrapping (#1474)', () => {
  it('moves a Latin word intact after CJK text in the same DrawingML run', () => {
    expect(lines([run('日本語Power')])).toEqual(['日本語', 'Power']);
  });

  it('moves a Latin word intact after CJK text across a formatting-run boundary', () => {
    expect(lines([run('日本語'), run('Power', 'Courier New')]))
      .toEqual(['日本語', 'Power']);
  });

  it('does not invent a break across a formatting seam inside a Latin word', () => {
    expect(lines([run('YoY+'), run('11.9%', 'Courier New')]))
      .toEqual(['YoY+11.9%']);
  });
});
