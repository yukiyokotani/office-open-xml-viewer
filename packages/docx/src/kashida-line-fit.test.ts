import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocSettings, DocxTextRun, SectionProps } from './types';

// WORD_KASHIDA_LINE_FIT (#2008): Windows Word reserves 10% (mediumKashida)
// and 20% (highKashida) of the line width in line fit, still justifies the
// lines to the full width, and opens no inter-CJK gaps on soft-wrapped kashida
// and thaiDistribute lines. Synthetic controls blk-0898-kfit, -kfitn and
// -kfit5 in compatibility modes unset, 11, 12, 14 and 15.

/** Stub glyph advance: every character 0.5 em (10 pt at 20 pt). */
function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// A 360 pt band from x = 20 to x = 380.
const section: SectionProps = {
  pageWidth: 400, pageHeight: 400,
  marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
};
const MODE_15: DocSettings = { compatibilityMode: 15 };

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  };
  return { type: 'text', ...run } as DocParagraph['runs'][number];
}

function para(
  runs: DocParagraph['runs'],
  alignment: DocParagraph['alignment'],
  extra: Partial<DocParagraph> = {},
): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', ...extra,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function lines(body: BodyElement, settings?: DocSettings): LineLayout[] {
  const node = layoutBodyModel([body], section, makeCtx(), {}, settings).pages[0]!.layers.body[0]!;
  if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
  return [...node.lines];
}

/** Every visible character of the line with its retained left edge (the
 * placement origin plus its cluster offset), in visual order. */
function characters(line: LineLayout): Array<{ character: string; xPt: number }> {
  return line.placements.filter((p): p is TextPlacement => p.kind === 'text')
    .flatMap((placement) => {
      const characters = [...placement.text];
      expect(placement.clusters).toHaveLength(characters.length);
      return placement.clusters.map((cluster, index) => ({
        character: characters[index]!,
        xPt: placement.origin.xPt + cluster.offset.xPt,
      }));
    })
    .filter(({ character }) => character.trim() !== '')
    .sort((a, b) => a.xPt - b.xPt);
}
const start = (line: LineLayout): number => characters(line)[0]!.xPt;
const end = (line: LineLayout): number => characters(line).at(-1)!.xPt + 10;

const MODES: Array<DocSettings | undefined> = [undefined, { compatibilityMode: 14 }, MODE_15];
// Twenty 40 pt words: a line of n words is 50n - 10 pt wide.
const WORDS = Array.from({ length: 20 }, () => 'aaaa').join(' ');
const wordsOn = (line: LineLayout): number => characters(line).length / 4;

describe('WORD_KASHIDA_LINE_FIT', () => {
  it('fits medium and high kashida lines into 0.9 and 0.8 of the width', () => {
    for (const settings of MODES) {
      const expected: Array<[DocParagraph['alignment'], number]> = [
        ['both', 7], ['distribute', 7], ['lowKashida', 7], ['thaiDistribute', 7],
        // 0.9 × 360 = 324 pt holds 6 words (290 pt), not 7 (340 pt).
        ['mediumKashida', 6],
        // 0.8 × 360 = 288 pt holds 5 words (240 pt), not 6 (290 pt).
        ['highKashida', 5],
      ];
      for (const [alignment, count] of expected) {
        const [first, second] = lines(para([textRun(WORDS)], alignment), settings);
        expect(wordsOn(first!), `${alignment} mode ${settings?.compatibilityMode}`).toBe(count);
        // Paint still justifies the soft-wrapped line to the full width.
        expect(end(first!)).toBeCloseTo(380);
        expect(wordsOn(second!)).toBe(count);
      }
    }
  });

  it('reserves a share of the line width from mode 15 and of the width from the leading indent before it', () => {
    // highKashida with a 50 pt right indent (310 pt line): mode 15 fits
    // 0.8 × 310 = 248 pt (5 words); earlier modes 310 - 0.2 × 360 = 238 pt
    // (4 words). A 50 pt left indent reserves 0.2 × 310 in every mode.
    for (const settings of MODES) {
      const [right] = lines(para([textRun(WORDS)], 'highKashida', { indentRight: 50 }), settings);
      expect(wordsOn(right!), `mode ${settings?.compatibilityMode}`).toBe(settings === MODE_15 ? 5 : 4);
      const [left] = lines(para([textRun(WORDS)], 'highKashida', { indentLeft: 50 }), settings);
      expect(wordsOn(left!), `mode ${settings?.compatibilityMode}`).toBe(5);
    }
  });

  it('keeps first-line and hanging offsets out of the pre-15 reserve', () => {
    // A 50 pt left indent with a 50 pt hanging indent: the first line is
    // 360 pt wide, later lines 310 pt. Before mode 15 both lose
    // 0.2 × (360 - 50) = 62 pt (298 pt: 6 words; 248 pt: 5 words); mode 15
    // fits 0.8 of each (288 pt: 5 words; 248 pt: 5 words).
    for (const settings of MODES) {
      const [first, second] = lines(
        para([textRun(WORDS)], 'highKashida', { indentLeft: 50, indentFirst: -50 }), settings);
      expect(wordsOn(first!), `mode ${settings?.compatibilityMode}`).toBe(settings === MODE_15 ? 5 : 6);
      expect(wordsOn(second!)).toBe(5);
    }
  });

  it('takes the start indent of a right-to-left paragraph for the pre-15 reserve', () => {
    const hebrew = Array.from({ length: 20 }, () => '\u05d0\u05d1\u05d2\u05d3').join(' ');
    const rtl = { bidi: true } as Partial<DocParagraph>;
    for (const settings of MODES) {
      // w:left is the start indent: 305 - 0.2 × 305 = 244 pt (5 words).
      const [start] = lines(para([textRun(hebrew)], 'highKashida', { ...rtl, indentLeft: 55 }), settings);
      expect(wordsOn(start!), `mode ${settings?.compatibilityMode}`).toBe(5);
      // w:right is the end indent: 305 - 0.2 × 360 = 233 pt (4 words) before
      // mode 15, 0.8 × 305 = 244 pt (5 words) from it.
      const [end] = lines(para([textRun(hebrew)], 'highKashida', { ...rtl, indentRight: 55 }), settings);
      expect(wordsOn(end!), `mode ${settings?.compatibilityMode}`).toBe(settings === MODE_15 ? 5 : 4);
    }
  });

  it('force-breaks an over-wide word at the fit width', () => {
    // A 40-character word (400 pt) under highKashida is force-broken where it
    // overflows 288 pt.
    const [first] = lines(para([textRun('b'.repeat(40))], 'highKashida'));
    expect(characters(first!)).toHaveLength(28);
  });

  it('leaves a soft-wrapped space-free CJK line at its natural width', () => {
    // 11 pt ideographs: 32 fill 352 pt of the 360 pt band.
    const cjk = textRun('\u3042'.repeat(70), { fontSize: 22 });
    for (const settings of MODES) {
      for (const alignment of ['lowKashida', 'thaiDistribute'] as const) {
        const [first] = lines(para([cjk], alignment, { defaultFontSize: 22 }), settings);
        expect(characters(first!)).toHaveLength(32);
        expect(end(first!) + 1).toBeCloseTo(372);
      }
      const [medium] = lines(para([cjk], 'mediumKashida', { defaultFontSize: 22 }), settings);
      // 0.9 × 360 = 324 pt: 29 ideographs, unstretched.
      expect(characters(medium!)).toHaveLength(29);
      expect(end(medium!) + 1).toBeCloseTo(20 + 29 * 11);
      for (const alignment of ['both', 'distribute'] as const) {
        const [first] = lines(para([cjk], alignment, { defaultFontSize: 22 }), settings);
        expect(end(first!) + 1).toBeCloseTo(380);
      }
    }
  });

  it('stretches only the spaces of a mixed soft-wrapped kashida line', () => {
    // 'ab ' (30 pt) and 25 ideographs of 13 pt leave 5 pt of slack.
    const runs = [textRun('ab '), textRun('\u3042'.repeat(70), { fontSize: 26 })];
    const [kashida] = lines(para(runs, 'lowKashida'));
    const placed = characters(kashida!);
    const pitches = placed.slice(3).map((char, index) => char.xPt - placed[index + 2]!.xPt);
    expect(Math.max(...pitches)).toBeCloseTo(13);
    expect(end(kashida!) + 3).toBeCloseTo(380);
    const [both] = lines(para(runs, 'both'));
    const bothPlaced = characters(both!);
    expect(bothPlaced[3]!.xPt - bothPlaced[2]!.xPt).toBeGreaterThan(13.1);
  });
});

describe('WORD_RTL_JUSTIFIED_VISIBLE_FIT', () => {
  it('fits a justified right-to-left line by its visible prefix', () => {
    // A 345 pt line: seven words are 340 pt visible, 350 pt with the line-end
    // space.
    const hebrew = Array.from({ length: 20 }, () => '\u05d0\u05d1\u05d2\u05d3').join(' ');
    const rtl = { bidi: true, indentRight: 15 } as Partial<DocParagraph>;
    for (const settings of MODES) {
      const [both] = lines(para([textRun(hebrew)], 'both', rtl), settings);
      expect(wordsOn(both!), `mode ${settings?.compatibilityMode}`).toBe(7);
      // Unmeasured alignments keep charging the line-end space.
      const [start] = lines(para([textRun(hebrew)], 'left', rtl), settings);
      expect(wordsOn(start!)).toBe(6);
    }
  });

  it('hangs the line-end space past the end edge and keeps the visible text in the band', () => {
    // The band runs from x = 35 (15 pt end indent on the left) to x = 380.
    const rtl = { bidi: true, indentRight: 15 } as Partial<DocParagraph>;
    const hebrew = Array.from({ length: 20 }, () => '\u05d0\u05d1\u05d2\u05d3').join(' ');
    for (const settings of MODES) {
      // Seven words justified across the band, the space hanging outside.
      const [justified] = lines(para([textRun(hebrew)], 'both', rtl), settings);
      expect(start(justified!)).toBeCloseTo(35);
      expect(end(justified!)).toBeCloseTo(380);
      // A 340 pt single word with its space (350 pt) fits the 345 pt band and
      // stays at the start edge.
      const word = '\u05d0'.repeat(34);
      const [single] = lines(para([textRun(`${word} ${word}`)], 'both', rtl), settings);
      expect(characters(single!)).toHaveLength(34);
      expect(end(single!)).toBeCloseTo(380);
      expect(start(single!)).toBeCloseTo(40);
      // highKashida: the line-end space hangs as well.
      const [kashida] = lines(para([textRun(hebrew)], 'highKashida', rtl), settings);
      expect(end(kashida!)).toBeCloseTo(380);
      expect(start(kashida!)).toBeCloseTo(35);
      // Latin text keeps its line-end space at the right of its run, where it
      // hangs past the start edge (blk-0898-kfit6 BL lines).
      const [latin] = lines(para([textRun(WORDS)], 'both', rtl), settings);
      expect(wordsOn(latin!)).toBe(7);
      expect(start(latin!)).toBeCloseTo(35);
      expect(end(latin!)).toBeCloseTo(380);
    }
  });
  it('fits a split or hyphenatable right-to-left word by the visible prefix too', () => {
    // A 345 pt band: 150 pt + space + 180 pt is 340 pt visible (350 pt with
    // the line-end space) and stays on the first line.
    const rtl = { bidi: true, indentRight: 15 } as Partial<DocParagraph>;
    const w1 = '\u05d0'.repeat(15);
    const half = '\u05d1'.repeat(9);
    for (const settings of MODES) {
      const [split] = lines(para([textRun(`${w1} ${half}`), textRun(`${half} `, { color: 'FF0000' }), textRun('\u05d2'.repeat(5))], 'both', rtl), settings);
      expect(characters(split!)).toHaveLength(33);
      // The visible text stays in the band at its start edge.
      expect(start(split!)).toBeGreaterThanOrEqual(35 - 1e-6);
      expect(end(split!)).toBeCloseTo(380);
      const soft = textRun('', { optionalHyphen: true } as Partial<DocxTextRun>);
      const [hyphen] = lines(para([textRun(`${w1} ${half}`), soft, textRun(`${half} `), textRun('\u05d2'.repeat(5))], 'both', rtl), settings);
      expect(characters(hyphen!).map(({ character }) => character).join('')).not.toContain('-');
      expect(characters(hyphen!)).toHaveLength(33);
    }
  });

  it('keeps a mixed-direction justified line inside the band when its line-end space is drawn inside', () => {
    // Alternating Latin and Hebrew words; the line ends with a Latin word
    // drawn leftmost, so its line-end space sits between runs.
    const rtl = { bidi: true, indentRight: 15 } as Partial<DocParagraph>;
    const text = Array.from({ length: 20 }, (_, index) => index % 2 === 0 ? 'abcd' : '\u05d0\u05d1\u05d2\u05d3').join(' ');
    for (const settings of MODES) {
      const [line] = lines(para([textRun(text)], 'both', rtl), settings);
      expect(characters(line!)).toHaveLength(28);
      expect(start(line!)).toBeGreaterThanOrEqual(35 - 1e-6);
      expect(end(line!)).toBeLessThanOrEqual(380 + 1e-6);
    }
  });

  it('keeps a final right-to-left line admitted by its visible width inside the band', () => {
    // A 340 pt Latin word and its 10 pt space in a 345 pt band, final line.
    const rtl = { bidi: true, indentRight: 15 } as Partial<DocParagraph>;
    for (const settings of MODES) {
      const [final] = lines(para([textRun(`${'a'.repeat(34)} `)], 'both', rtl), settings);
      expect(start(final!)).toBeCloseTo(35);
      expect(end(final!)).toBeCloseTo(375);
      // With room for the space it stays inside the band, as before.
      const [short] = lines(para([textRun(`${'a'.repeat(20)} `)], 'both', rtl), settings);
      expect(end(short!)).toBeCloseTo(370);
    }
  });
  it('absorbs an overflowing line-end space drawn between segments of a final line', () => {
    // A 250 pt Hebrew word, a space, an 80 pt Latin word and its space: 340 pt
    // visible, 350 pt in all, in a 345 pt band. The Latin word is drawn
    // leftmost, so its space sits between the two words and takes the 5 pt.
    const rtl = { bidi: true, indentRight: 15 } as Partial<DocParagraph>;
    for (const settings of MODES) {
      const [final] = lines(para([textRun(`${'\u05d0'.repeat(25)} ${'a'.repeat(8)} `)], 'both', rtl), settings);
      expect(characters(final!)).toHaveLength(33);
      expect(start(final!)).toBeCloseTo(35);
      expect(end(final!)).toBeCloseTo(380);
    }
  });
  it('charges the right-to-left line-end space where it cannot hang (wordWrap off)', () => {
    const rtl = { bidi: true, indentRight: 15, wordWrap: false } as Partial<DocParagraph>;
    const hebrew = Array.from({ length: 20 }, () => '\u05d0\u05d1\u05d2\u05d3').join(' ');
    for (const settings of MODES) {
      const [first] = lines(para([textRun(hebrew)], 'both', rtl), settings);
      expect(wordsOn(first!)).toBe(6);
      expect(start(first!)).toBeGreaterThanOrEqual(35 - 1e-6);
      expect(end(first!)).toBeLessThanOrEqual(380 + 1e-6);
    }
  });

  it('resolves split line-end spaces on the sides where each is drawn', () => {
    // Seven Latin words fill 340 pt of a 345 pt right-to-left band; the line
    // ends with the seventh word's space and a separately coloured space.
    const rtl = { bidi: true, indentRight: 15 } as Partial<DocParagraph>;
    const seven = Array.from({ length: 7 }, () => 'aaaa').join(' ');
    for (const settings of MODES) {
      const [first] = lines(para([textRun(`${seven} `), textRun(' ', { color: 'FF0000' }), textRun('bbbb cccc')], 'both', rtl), settings);
      expect(wordsOn(first!)).toBe(7);
      expect(start(first!)).toBeCloseTo(35);
      expect(end(first!)).toBeCloseTo(380);
    }
  });

  it('excludes a line-end space suffix split across runs from optional-hyphen fitting', () => {
    // A 345 pt band from x = 20: 150 pt + space + a 180 pt word holding a
    // soft hyphen is 340 pt visible; its space and a separately coloured one follow.
    const ltr = { indentRight: 15 } as Partial<DocParagraph>;
    const soft = textRun('', { optionalHyphen: true } as Partial<DocxTextRun>);
    for (const settings of MODES) {
      const [first] = lines(para([textRun(`${'a'.repeat(15)} ${'b'.repeat(9)}`), soft, textRun(`${'b'.repeat(9)} `),
        textRun(' ', { color: 'FF0000' }), textRun('cccc')], 'both', ltr), settings);
      expect(characters(first!).map(({ character }) => character).join('')).not.toContain('-');
      expect(characters(first!)).toHaveLength(33);
      expect(end(first!)).toBeCloseTo(365);
    }
  });
});
