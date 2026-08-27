import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph, naturalWidthExceedsBbox } from './renderer.js';
import type { Paragraph, TextBody } from './types';

function mockCtx(): CanvasRenderingContext2D {
  let font = '20px Arial';
  const fontPx = () => Number.parseFloat(/([0-9.]+)px/.exec(font)?.[1] ?? '20');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: [...text].length * fontPx() }),
  } as unknown as CanvasRenderingContext2D;
}

function run(text: string, overrides: Partial<TextRunData> = {}): TextRunData {
  return {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: null,
    color: '000000',
    fontFamily: 'Arial',
    ...overrides,
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

function lineText(line: { segments: { text: string }[] }): string {
  return line.segments.map((segment) => segment.text).join('');
}

function body(runs: TextRunData[]): TextBody {
  return {
    verticalAnchor: 't',
    paragraphs: [paragraph(runs)],
    defaultFontSize: 15,
    defaultBold: null,
    defaultItalic: null,
    lIns: 0,
    rIns: 0,
    tIns: 0,
    bIns: 0,
    wrap: 'square',
    vert: 'horz',
    autoFit: 'sp',
  };
}

describe('pptx baseline run sizing', () => {
  it('measures a superscript citation at 65% while retaining the authored size', () => {
    const lines = layoutParagraph(
      mockCtx(),
      paragraph([
        run('aaaa bbbb citation'),
        run(' 1', { baseline: 30000 }),
      ]),
      390,
      20,
      '000000',
      1,
      0,
    );

    expect(lines.map(lineText)).toEqual(['aaaa bbbb citation 1']);
    const citation = lines[0].segments.at(-1);
    expect(citation?.font).toContain('13px');
    expect(citation?.sizePx).toBe(20);
  });

  it('keeps the full-size control greedy, including PowerPoint-style orphaning', () => {
    const lines = layoutParagraph(
      mockCtx(),
      paragraph([
        run('aaaa bbbb citation'),
        run(' 1', { baseline: 0 }),
      ]),
      390,
      20,
      '000000',
      1,
      0,
    );

    expect(lines.map(lineText)).toEqual(['aaaa bbbb citation ', '1']);
    expect(lines[0].segments.at(-1)?.font).toContain('20px');
  });

  it('uses the reduced width in the shape-autofit preflight too', () => {
    const scale = 96 / 914400;
    const context = mockCtx();
    const prefix = run('aaaa bbbb citation');

    expect(naturalWidthExceedsBbox(
      context,
      body([prefix, run(' 1', { baseline: 0 })]),
      390,
      0,
      0,
      scale,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    )).toBe(true);

    expect(naturalWidthExceedsBbox(
      context,
      body([prefix, run(' 1', { baseline: 30000 })]),
      390,
      0,
      0,
      scale,
      { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    )).toBe(false);
  });
});
