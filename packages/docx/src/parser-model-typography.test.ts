import { describe, expect, it } from 'vitest';
import { paragraphAcquisitionInput } from './parser-model.js';
import type { DocParagraph } from './types.js';
import type {
  InternalParagraphTypographyWire,
  InternalRunTypographyWire,
} from './layout/typography-input.js';

const missing = () => ({ status: 'missing' as const, raw: null, value: null });

const runWire: InternalRunTypographyWire = {
  strike: false,
  doubleStrike: false,
  caps: false,
  smallCaps: false,
  colorAuto: false,
  verticalAlign: missing(),
  positionPt: missing(),
  snapToGrid: null,
  characterSpacingPt: null,
  characterScale: null,
  kerningThresholdPt: null,
  emphasis: missing(),
  languages: { eastAsia: null, bidi: null },
  eastAsianLayout: {
    vert: null,
    vertCompress: null,
    combine: null,
    combineBrackets: missing(),
  },
};

const paragraphWire: InternalParagraphTypographyWire = {
  borders: {
    bar: {
      val: { status: 'valid', raw: 'single', value: 'single' },
      color: missing(),
      themeColor: missing(),
      themeTint: missing(),
      themeShade: missing(),
      sizePt: { status: 'valid', raw: '8', value: 1 },
      spacePt: missing(),
      shadow: missing(),
      frame: missing(),
    },
  },
};

describe('paragraph parser-model typography projection', () => {
  it('replaces parser-private wires with named immutable layout inputs', () => {
    const paragraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [],
      runs: [{
        type: 'text', text: 'x', fontSize: 10,
        __typographyAcquisition: runWire,
      }],
      __paragraphTypographyAcquisition: paragraphWire,
    } as unknown as DocParagraph;

    const input = paragraphAcquisitionInput(paragraph, {
      story: 'body', storyInstance: 'body', path: [0],
    }) as unknown as {
      __paragraphTypographyAcquisition?: unknown;
      typographyInput?: unknown;
      runs: readonly {
        __typographyAcquisition?: unknown;
        typographyInput?: unknown;
      }[];
    };

    expect(input.__paragraphTypographyAcquisition).toBeUndefined();
    expect(input.typographyInput).toEqual(paragraphWire);
    expect(input.runs[0].__typographyAcquisition).toBeUndefined();
    expect(input.runs[0].typographyInput).toEqual({ sourceText: 'x', ...runWire });
    expect(structuredClone(input.typographyInput)).toEqual(input.typographyInput);
    expect(Object.isFrozen(input.typographyInput)).toBe(true);
    expect(Object.isFrozen(input.runs[0].typographyInput)).toBe(true);
  });
});
