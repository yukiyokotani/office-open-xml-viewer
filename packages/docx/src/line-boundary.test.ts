import { describe, expect, it } from 'vitest';
import {
  DEFAULT_KINSOKU_RULES,
  type KinsokuRules,
} from '@silurus/ooxml-core';
import {
  layoutLines,
  type LayoutLine,
  type LayoutSeg,
  type LayoutTextSeg,
  type LineBoundary,
} from './line-layout.js';
import type { TabStop } from './types.js';

function makeLinearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const fontSize = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = fontSize();
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function textSeg(text: string, extra: Partial<LayoutTextSeg> = {}): LayoutTextSeg {
  return {
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 10,
    color: null,
    fontFamily: 'Times New Roman',
    vertAlign: null,
    measuredWidth: 0,
    ...extra,
  };
}

const lineBreak = (fontSize = 10): LayoutSeg => ({ lineBreak: true, fontSize, measuredWidth: 0 });

interface Fixture {
  readonly name: string;
  readonly width: number;
  readonly segs: () => LayoutSeg[];
  readonly tabStops?: TabStop[];
  readonly kinsoku?: KinsokuRules;
  readonly baseRtl?: boolean;
  readonly hasTrailingBreak?: boolean;
  readonly expectedMidSegmentIndex?: number;
  readonly fitTextRegion?: { readonly firstSegIndex: number; readonly lastSegIndex: number };
}

const crossRunKinsoku: KinsokuRules = {
  enabled: true,
  lineStartForbidden: new Set(['。'.codePointAt(0)!]),
  lineEndForbidden: new Set(),
};

const fixtures: Fixture[] = [
  {
    name: 'plain Latin words',
    width: 35,
    segs: () => ['alpha ', 'bravo ', 'charlie ', 'delta ', 'echo ', 'foxtrot '].map((text) => textSeg(text)),
  },
  {
    name: 'CJK per-glyph split',
    width: 25,
    segs: () => [textSeg('あ'.repeat(18))],
    expectedMidSegmentIndex: 0,
  },
  {
    name: 'kinsoku cross-run retraction',
    width: 20,
    segs: () => [textSeg('あいうえ'), textSeg('。かきくけこさしす')],
    kinsoku: crossRunKinsoku,
  },
  {
    name: 'manual breaks including a trailing break',
    width: 200,
    segs: () => [textSeg('first'), lineBreak(), textSeg('second'), lineBreak(18)],
    hasTrailingBreak: true,
  },
  {
    name: 'right-aligned custom tab with committed trailing text',
    width: 60,
    tabStops: [{ pos: 40, alignment: 'right', leader: 'dot' }],
    segs: () => [
      textSeg('name'),
      { isTab: true, fontSize: 10, measuredWidth: 0 },
      textSeg('99'),
      lineBreak(),
      textSeg('tail '),
      textSeg('words '),
      textSeg('wrap'),
    ],
  },
  {
    name: 'over-long Latin overflow-wrap',
    width: 20,
    segs: () => [textSeg('abcdefghijklmnopqrst')],
    expectedMidSegmentIndex: 0,
  },
  {
    name: 'small-caps glued group',
    width: 55,
    segs: () => [
      textSeg('Lead '),
      textSeg('I', { smallCaps: true }),
      textSeg('NTRODUCTION', { smallCaps: true, joinPrev: true }),
      textSeg(' tail '),
      textSeg('words'),
    ],
  },
  {
    name: 'bidi RTL ordinary tab',
    width: 70,
    tabStops: [{ pos: 50, alignment: 'left', leader: 'dot' }],
    baseRtl: true,
    segs: () => [
      textSeg('פתיחה', { rtl: true }),
      lineBreak(),
      textSeg('כותרת ', { rtl: true }),
      { isTab: true, fontSize: 10, measuredWidth: 0 },
      textSeg('12', { rtl: true }),
      lineBreak(),
      textSeg('סיום', { rtl: true }),
    ],
  },
  {
    name: 'atomic fitText region',
    width: 30,
    fitTextRegion: { firstSegIndex: 1, lastSegIndex: 2 },
    segs: () => [
      textSeg('lead '),
      textSeg('ab', {
        fitTextRegionIndex: 7,
        fitTextRegionStart: true,
      }),
      textSeg('cd', {
        fitTextRegionIndex: 7,
        fitTextRegionEnd: true,
      }),
      textSeg(' tail tail tail'),
    ],
  },
  {
    name: 'ruby ascent reserve',
    width: 30,
    segs: () => [
      textSeg('lead '),
      textSeg('ruby', { ruby: { text: 'ルビ', fontSizePt: 8 } }),
      textSeg(' tail tail tail'),
    ],
  },
  {
    name: 'astral CJK split',
    width: 12,
    segs: () => [textSeg('𠀋'.repeat(12))],
    expectedMidSegmentIndex: 0,
  },
  {
    name: 'Latin lead with glued breakable CJK follower',
    width: 30,
    segs: () => [
      textSeg('intro'),
      lineBreak(),
      textSeg('Roman'),
      textSeg('、あいうえお', { joinPrev: true }),
      textSeg(' tail'),
    ],
    expectedMidSegmentIndex: 3,
  },
];

// Vertical text does not enter this horizontal breaking loop, and state-sensitive
// fields are resolved before LayoutSegs exist, so neither belongs in this layer's corpus.

function layoutFixture(fixture: Fixture, startBoundary?: LineBoundary): LayoutLine[] {
  const segs = fixture.segs().map((seg) => ({ ...seg })) as LayoutSeg[];
  return layoutLines(
    makeLinearCtx(),
    segs,
    fixture.width,
    0,
    1,
    fixture.tabStops,
    undefined,
    {},
    undefined,
    fixture.kinsoku ?? DEFAULT_KINSOKU_RULES,
    undefined,
    36,
    fixture.width,
    fixture.baseRtl ?? false,
    false,
    false,
    startBoundary,
  );
}

function lineStructure(lines: LayoutLine[]) {
  return lines.map((line) => ({
    segments: line.segments.map((segment) => {
      if ('text' in segment) {
        return { kind: 'text' as const, text: segment.text, measuredWidth: segment.measuredWidth };
      }
      if ('isTab' in segment) {
        return {
          kind: 'tab' as const,
          measuredWidth: segment.measuredWidth,
          leader: segment.leader,
        };
      }
      if ('imagePath' in segment) {
        return { kind: 'image' as const, measuredWidth: segment.measuredWidth };
      }
      return { kind: 'math' as const, measuredWidth: segment.measuredWidth };
    }),
    height: line.height,
    ascent: line.ascent,
    descent: line.descent,
    intendedSingle: line.intendedSingle,
    xOffset: line.xOffset,
    availWidth: line.availWidth,
    hasRuby: line.hasRuby === true,
    endsWithBreak: line.endsWithBreak === true,
    consumedEnd: line.consumedEnd,
  }));
}

describe('resumed text segment geometry', () => {
  it('does not emergency-split a word at an intrinsic-width round-trip boundary', () => {
    const fixture = { name: 'exact-width word', width: 25, segs: () => [textSeg('other')] };
    // Margin subtraction / point conversion can lose one representable step.
    expect(layoutFixture({ ...fixture, width: 25 - Number.EPSILON * 25 })).toHaveLength(1);
    // An actual content deficit still requires a split.
    expect(layoutFixture({ ...fixture, width: 24.99 })).toHaveLength(2);
  });

  it('rebases punctuation compressions after the pagination boundary', () => {
    const resumed = layoutFixture({
      name: 'compressed punctuation resume',
      width: 100,
      segs: () => [textSeg('甲、乙。丙', {
        punctuationCompressions: [
          { end: 2, adjustmentPt: -5 },
          { end: 4, adjustmentPt: -5 },
        ],
      })],
    }, { segIndex: 0, charOffset: 2 });
    const first = resumed[0]?.segments[0];

    expect(first && 'text' in first ? first.text : undefined).toBe('乙。丙');
    expect(first && 'text' in first ? first.punctuationCompressions : undefined)
      .toEqual([{ end: 2, adjustmentPt: -5 }]);
    expect(first?.measuredWidth).toBeCloseTo(10, 5);
  });
});

function textSequence(lines: LayoutLine[]): string[] {
  return lines.map((line) => line.segments
    .filter((segment): segment is LayoutTextSeg => 'text' in segment)
    .map((segment) => segment.text)
    .join(''));
}

describe('strict tab-stop alignment aliases', () => {
  const tabFixture = (alignment: TabStop['alignment']): Fixture => ({
    name: alignment,
    width: 100,
    tabStops: [{ pos: 60, alignment, leader: 'none' }],
    segs: () => [
      textSeg('lead'),
      { isTab: true, fontSize: 10, measuredWidth: 0 },
      textSeg('tail'),
    ],
  });

  it('maps start and num to leading, and end to trailing, in an LTR reading frame', () => {
    const structure = (alignment: TabStop['alignment']) =>
      lineStructure(layoutFixture(tabFixture(alignment)));

    expect(structure('start')).toEqual(structure('left'));
    expect(structure('num')).toEqual(structure('left'));
    expect(structure('end')).toEqual(structure('right'));
  });
});

describe('layoutLines consumed-content boundaries', () => {
  for (const fixture of fixtures) {
    it(`reproduces every suffix for ${fixture.name}`, () => {
      const lines = layoutFixture(fixture);
      expect(lines.length).toBeGreaterThanOrEqual(2);
      if (fixture.expectedMidSegmentIndex !== undefined) {
        const original = fixture.segs();
        expect(lines.some((line) => {
          const boundary = line.consumedEnd;
          if (
            !boundary
            || boundary.segIndex !== fixture.expectedMidSegmentIndex
            || boundary.charOffset <= 0
          ) return false;
          const segment = original[boundary.segIndex];
          return segment !== undefined
            && 'text' in segment
            && boundary.charOffset < segment.text.length;
        })).toBe(true);
      }
      if (fixture.fitTextRegion) {
        const { firstSegIndex, lastSegIndex } = fixture.fitTextRegion;
        expect(lines.every((line) => {
          const segIndex = line.consumedEnd?.segIndex;
          return segIndex === undefined || segIndex <= firstSegIndex || segIndex > lastSegIndex;
        })).toBe(true);
      }
      if (fixture.baseRtl) {
        expect(lines.some((line) => line.segments.some((segment) =>
          'isTab' in segment && segment.measuredWidth > 0,
        ))).toBe(true);
      }
      if (fixture.name === 'ruby ascent reserve') {
        expect(lines.some((line) => line.hasRuby === true)).toBe(true);
      }

      for (let i = 0; i < lines.length; i++) {
        const boundary = lines[i].consumedEnd;
        expect(boundary).toBeDefined();
        if (!boundary) continue;

        const expected = lineStructure(lines.slice(i + 1));
        const suffix = lineStructure(layoutFixture(fixture, boundary));

        expect(suffix).toEqual(expected);
      }
    });
  }

  it('places a cross-run kinsoku boundary inside the retracted source segment', () => {
    const fixture = fixtures.find((candidate) => candidate.name === 'kinsoku cross-run retraction')!;
    const lines = layoutFixture(fixture);

    expect(lines.some((line) =>
      line.consumedEnd?.segIndex === 0 && line.consumedEnd.charOffset > 0,
    )).toBe(true);
  });

  it('re-emits a trailing-break empty line from an earlier original-stream boundary', () => {
    const fixture = fixtures.find((candidate) => candidate.hasTrailingBreak)!;
    const lines = layoutFixture(fixture);
    const boundary = lines[0].consumedEnd;

    expect(boundary).toBeDefined();
    expect(textSequence(lines)).toEqual(['first', 'second', '']);
    expect(textSequence(layoutFixture(fixture, boundary))).toEqual(['second', '']);
  });

  for (const fixture of fixtures) {
    it(`records monotonic boundaries ending at END for ${fixture.name}`, () => {
      const segCount = fixture.segs().length;
      const original = fixture.segs();
      const boundaries = layoutFixture(fixture).map((line) => line.consumedEnd);
      expect(boundaries.every((boundary) => boundary !== undefined)).toBe(true);

      for (let i = 1; i < boundaries.length; i++) {
        const previous = boundaries[i - 1]!;
        const current = boundaries[i]!;
        expect(
          current.segIndex > previous.segIndex
            || (current.segIndex === previous.segIndex && current.charOffset >= previous.charOffset),
        ).toBe(true);
      }
      for (const boundary of boundaries) {
        if (!boundary) continue;
        const segment = original[boundary.segIndex];
        if (!segment || !('text' in segment)) continue;
        const offset = boundary.charOffset;
        const previous = segment.text[offset - 1] ?? '';
        const current = segment.text[offset] ?? '';
        expect(
          /[\uD800-\uDBFF]/.test(previous) && /[\uDC00-\uDFFF]/.test(current),
        ).toBe(false);
      }
      expect(boundaries.at(-1)).toEqual({ segIndex: segCount, charOffset: 0 });
    });
  }
});
