import { describe, expect, it } from 'vitest';
import type { DocParagraph, DocxTextRun, FieldRun } from '../types.js';
import {
  paragraphTypographyAcquisitionInput,
  runTypographyAcquisitionInput,
  type InternalParagraphTypographyWire,
  type InternalRunTypographyWire,
} from './typography-input.js';

const value = <T>(raw: string, parsed: T) => ({ status: 'valid' as const, raw, value: parsed });
const missing = () => ({ status: 'missing' as const, raw: null, value: null });

const border = {
  val: value('double', 'double'),
  color: value('Auto', 'auto'),
  themeColor: value('accent1', 'accent1'),
  themeTint: value('80', '80'),
  themeShade: value('40', '40'),
  sizePt: value('24', 3),
  spacePt: value('2', 2),
  shadow: value('1', true),
  frame: value('0', false),
} as const;

const runWire: InternalRunTypographyWire = {
  underline: {
    val: value('words', 'words'),
    color: value('FF0000', 'ff0000'),
    themeColor: value('accent2', 'accent2'),
    themeTint: value('20', '20'),
    themeShade: missing(),
  },
  strike: true,
  doubleStrike: false,
  caps: true,
  smallCaps: false,
  colorAuto: true,
  verticalAlign: value('superscript', 'super'),
  positionPt: value('4', 2),
  snapToGrid: false,
  characterSpacingPt: 1,
  characterScale: 0.8,
  fitText: { valTwips: 2400, id: '-7' },
  kerningThresholdPt: 12,
  emphasis: value('dot', 'dot'),
  languages: { eastAsia: 'ja-jp', bidi: 'ar-sa' },
  eastAsianLayout: {
    vert: true,
    vertCompress: false,
    combine: true,
    combineBrackets: value('round', 'round'),
  },
  border,
  ruby: {
    align: value('distributeSpace', 'distributeSpace'),
    baseFontSizePt: value('24', 12),
    raisePt: value('10', 5),
    language: value('ja-JP', 'ja-jp'),
    guideRuns: [{
      text: 'かん',
      fontFamily: 'Yu Gothic',
      fontSizePt: 6,
      bold: true,
      italic: false,
      color: '112233',
      language: 'ja-jp',
    }],
  },
  revision: {
    kind: 'insertion',
    id: value('7', '7'),
    author: 'A',
    date: '2026-07-14T00:00:00Z',
  },
};

describe('private typography acquisition projection', () => {
  it('projects text and field runs with identical typography facts', () => {
    const textRun = {
      type: 'text', text: 'ABC',
      __typographyAcquisition: runWire,
    } as unknown as DocxTextRun;
    const fieldRun = {
      type: 'field', fallbackText: 'ABC',
      __typographyAcquisition: runWire,
    } as unknown as FieldRun;

    const text = runTypographyAcquisitionInput(textRun);
    const field = runTypographyAcquisitionInput(fieldRun);

    expect(text).toEqual({ sourceText: 'ABC', ...runWire });
    expect(field).toEqual(text);
    expect(text).not.toBe(runWire);
    expect(structuredClone(text)).toEqual(text);
    expect(Object.isFrozen(text)).toBe(true);
    expect(Object.isFrozen(text?.ruby?.guideRuns)).toBe(true);
  });

  it('keeps missing and invalid required values diagnostic-capable', () => {
    const invalidWire: InternalRunTypographyWire = {
      ...runWire,
      underline: {
        ...runWire.underline!,
        val: { status: 'invalid', raw: 'not-an-underline', value: null },
      },
      revision: {
        ...runWire.revision!,
        id: missing(),
      },
    };
    const run = {
      type: 'text', text: 'x', __typographyAcquisition: invalidWire,
    } as unknown as DocxTextRun;

    const input = runTypographyAcquisitionInput(run);

    expect(input?.underline?.val).toEqual({
      status: 'invalid', raw: 'not-an-underline', value: null,
    });
    expect(input?.revision?.id).toEqual(missing());
  });

  it('projects all six paragraph borders without widening DocParagraph', () => {
    const wire: InternalParagraphTypographyWire = {
      borders: {
        top: border,
        right: border,
        bottom: border,
        left: border,
        between: border,
        bar: border,
      },
    };
    const paragraph = {
      type: 'paragraph', runs: [], __paragraphTypographyAcquisition: wire,
    } as unknown as DocParagraph;

    const input = paragraphTypographyAcquisitionInput(paragraph);

    expect(input).toEqual(wire);
    expect(input?.borders.bar?.themeTint.value).toBe('80');
    expect(structuredClone(input)).toEqual(input);
    expect(Object.isFrozen(input?.borders)).toBe(true);
  });

  it('returns undefined for public hand-built values without a private wire', () => {
    expect(runTypographyAcquisitionInput({ type: 'text', text: 'x' } as unknown as DocxTextRun))
      .toBeUndefined();
    expect(paragraphTypographyAcquisitionInput({ runs: [] } as unknown as DocParagraph))
      .toBeUndefined();
  });
});
