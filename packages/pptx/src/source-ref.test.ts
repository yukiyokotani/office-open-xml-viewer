import { describe, expect, it } from 'vitest';
import type { TextRunData, TextSourceRef } from '@silurus/ooxml-core';
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

function run(text: string, sourceRefs: TextSourceRef[]): TextRunData {
  return {
    type: 'text',
    text,
    sourceRefs,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 20,
    color: '000000',
    fontFamily: 'Arial',
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

describe('PPTX text source references', () => {
  it('rebases source intervals onto automatically wrapped visual segments', () => {
    const sourceRefs: TextSourceRef[] = [{
      path: [{ namespaceUri: 'urn:a', localName: 't', index: 0 }],
      textStart: 0,
      textEnd: 10,
      sourceStart: 0,
      sourceEnd: 10,
    }];
    const lines = layoutParagraph(
      mockCtx(),
      paragraph([run('Alpha Beta', sourceRefs)]),
      60,
      20,
      '000000',
      1,
      0,
    );
    const segments = lines.flatMap((line) => line.segments.filter((segment) => segment.text));

    expect(segments.map((segment) => segment.text)).toEqual(['Alpha ', 'Beta']);
    expect(segments.map((segment) => segment.sourceRefs)).toEqual([
      [expect.objectContaining({ textStart: 0, textEnd: 6, sourceStart: 0, sourceEnd: 6 })],
      [expect.objectContaining({ textStart: 0, textEnd: 4, sourceStart: 6, sourceEnd: 10 })],
    ]);
  });

  it('keeps separate XML nodes when same-style runs merge visually', () => {
    const lines = layoutParagraph(
      mockCtx(),
      paragraph([
        run('Hello', [{
          path: [{ localName: 't', index: 0 }],
          textStart: 0,
          textEnd: 5,
          sourceStart: 0,
          sourceEnd: 5,
        }]),
        run('世界', [{
          path: [{ localName: 't', index: 1 }],
          textStart: 0,
          textEnd: 2,
          sourceStart: 0,
          sourceEnd: 2,
        }]),
      ]),
      200,
      20,
      '000000',
      1,
      0,
    );
    const refs = lines.flatMap((line) => line.segments).flatMap((segment) => segment.sourceRefs ?? []);

    expect(refs.map((ref) => [ref.path[0].index, ref.textStart, ref.textEnd])).toEqual([
      [0, 0, 5],
      [1, 5, 7],
    ]);
  });
});
