import { describe, expect, it } from 'vitest';
import { buildSegments, layoutLines, type LayoutTextSeg } from './line-layout.js';
import type { DocRun } from './types.js';

const ENV = { pageIndex: 0, totalPages: 1 };

function textRun(text: string, underline = false): DocRun {
  return {
    type: 'text',
    text,
    bold: false,
    italic: false,
    underline,
    strikethrough: false,
    fontSize: 12,
    color: null,
    fontFamily: 'serif',
    isLink: false,
    background: null,
    vertAlign: null,
    allCaps: false,
    smallCaps: false,
    doubleStrikethrough: false,
  } as unknown as DocRun;
}

function textSegments(runs: DocRun[]): LayoutTextSeg[] {
  return buildSegments(runs, ENV).filter((seg): seg is LayoutTextSeg => 'text' in seg);
}

describe('buildSegments UAX #14 LB28 boundary glue', () => {
  it.each([
    ['AL × AL', '<', 'a'],
    ['AL × HL', '<', 'א'],
    ['HL × AL', 'א', 'a'],
    ['HL × HL', 'א', 'ב'],
  ])('marks the following segment joinPrev for %s', (_label, prev, next) => {
    expect(textSegments([textRun(prev), textRun(next)]).map((seg) => seg.joinPrev))
      .toEqual([undefined, true]);
  });

  it.each([
    ['trailing whitespace', [textRun('< '), textRun('a')]],
    ['leading whitespace', [textRun('<'), textRun(' a')]],
    ['zero-width space', [textRun('<\u200b'), textRun('a')]],
    ['CJK-breakable text', [textRun('<'), textRun('漢')]],
    ['SEA dictionary text', [textRun('<'), textRun('ก')]],
    [
      'non-text boundary',
      [
        textRun('<'),
        { type: 'break', breakType: 'line' } as DocRun,
        textRun('a'),
      ],
    ],
  ])('does not add glue across %s', (_label, runs) => {
    expect(textSegments(runs).every((seg) => seg.joinPrev === undefined)).toBe(true);
  });

  it('does not create a break inside consecutive spaces split by a source-run boundary (LB7)', () => {
    expect(textSegments([
      textRun('deadline ', true),
      textRun(' ', true),
      textRun('next'),
    ]).map((segment) => segment.joinPrev)).toEqual([undefined, true, undefined]);
  });

  it('keeps a trailing space split by script shaping with the preceding text from the same run', () => {
    expect(textSegments([textRun('日 ')]).map((segment) => ({
      text: segment.text,
      joinPrev: segment.joinPrev,
      sourceRunIndex: segment.sourceRunIndex,
    }))).toEqual([
      { text: '日', joinPrev: undefined, sourceRunIndex: 0 },
      { text: ' ', joinPrev: true, sourceRunIndex: 0 },
    ]);
  });

  it('keeps trailing-space fit invariant when formatting splits the space sequence', () => {
    const context = {
      font: '10px serif',
      letterSpacing: '0px',
      measureText(value: string) {
        return {
          width: [...value].length * 5,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as CanvasRenderingContext2D;
    const lineTexts = (runs: readonly DocRun[]) => layoutLines(
      context,
      buildSegments(runs, ENV),
      20,
      0,
      1,
    ).map((line) => line.segments.map((segment) =>
      'text' in segment ? segment.text : '').join(''));

    const sameRun = lineTexts([
      textRun('X '),
      textRun('AB  ', true),
      textRun('CD'),
    ]);
    const splitRuns = lineTexts([
      textRun('X '),
      textRun('AB ', true),
      textRun(' ', true),
      textRun('CD'),
    ]);

    expect(sameRun).toEqual([
      'X AB  ',
      'CD',
    ]);
    expect(splitRuns).toEqual(sameRun);
  });

});
