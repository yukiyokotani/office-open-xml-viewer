import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocSettings, DocxTextRun, SectionProps } from './types';

// WORD_LINE_END_JUSTIFICATION (#2009, #2010): how Windows Word justifies the
// lines of a justified paragraph that end at a manual break, a page or column
// break, or the paragraph end. Synthetic controls in compatibility modes unset,
// 11, 12, 14 and 15, with and without doNotExpandShiftReturn.

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
const BR = { type: 'break', breakType: 'line' } as DocParagraph['runs'][number];

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

describe('WORD_LINE_END_JUSTIFICATION', () => {
  it('justifies a line ended by a manual break unless doNotExpandShiftReturn is set', () => {
    for (const settings of [undefined, MODE_15]) {
      const [broken, final] = lines(para([textRun('aaaa bbbb'), BR, textRun('cc')], 'both'), settings);
      expect(start(broken!)).toBeCloseTo(20);
      expect(end(broken!)).toBeCloseTo(380);
      expect(end(final!)).toBeCloseTo(40);
      const [kept] = lines(para([textRun('aaaa bbbb'), BR, textRun('cc')], 'both'),
        { ...settings, doNotExpandShiftReturn: true });
      expect(end(kept!)).toBeCloseTo(110);
    }
  });

  it('leaves a single word before a manual break at its natural width', () => {
    const [broken] = lines(para([textRun('aaaa'), BR, textRun('cc')], 'both'));
    expect(end(broken!)).toBeCloseTo(60);
  });

  it('justifies only after the last tab', () => {
    // 'a b' has a gap before the tab; the default 36 pt tab stop ends at 56.
    const [broken] = lines(para([textRun('a b\tcc dd'), BR, textRun('x')], 'both'));
    const tab = broken!.placements.find((placement) => placement.kind === 'tab')!;
    expect(tab.bounds!.xPt + tab.bounds!.widthPt).toBeCloseTo(56);
    const chars = characters(broken!);
    expect(chars.map(({ xPt }) => xPt).slice(0, 3)).toEqual([20, 40, 56].map((x) => expect.closeTo(x)));
    expect(end(broken!)).toBeCloseTo(380);
    const [final] = lines(para([textRun('a b\tcc dd')], 'distribute'));
    expect(characters(final!).map(({ xPt }) => xPt).slice(0, 3)).toEqual([20, 40, 56].map((x) => expect.closeTo(x)));
    expect(end(final!)).toBeCloseTo(380);
  });

  it('leaves a space-free CJK line before a break ragged under the kashida types', () => {
    const [kashida] = lines(para([textRun('あいうえ'), BR, textRun('x')], 'lowKashida'));
    expect(end(kashida!)).toBeCloseTo(60);
    const [both] = lines(para([textRun('あいうえ'), BR, textRun('x')], 'both'));
    expect(end(both!)).toBeCloseTo(380);
  });

  it('spreads final and break-ended distribute lines per character before mode 15', () => {
    const [broken, final] = lines(para([textRun('ab cd'), BR, textRun('efg')], 'distribute'));
    for (const line of [broken!, final!]) {
      const chars = characters(line);
      const pitches = chars.slice(1).map((char, index) => char.xPt - chars[index]!.xPt);
      expect(start(line)).toBeCloseTo(20);
      expect(end(line)).toBeCloseTo(380);
      // Equal extra after every character, the space included.
      if (line === final) expect(Math.max(...pitches) - Math.min(...pitches)).toBeCloseTo(0);
    }
    const brokenChars = characters(broken!);
    // b, then a space and c: two equal steps across the space.
    expect(brokenChars[2]!.xPt - brokenChars[1]!.xPt)
      .toBeCloseTo(2 * (brokenChars[1]!.xPt - brokenChars[0]!.xPt));
  });

  it('justifies final distribute lines at their gaps from mode 15 and centres a single word', () => {
    const [broken, final] = lines(para([textRun('aaaa bbbb'), BR, textRun('cccc')], 'distribute'), MODE_15);
    expect(end(broken!)).toBeCloseTo(380);
    expect(characters(broken!)[4]!.xPt).toBeCloseTo(340);
    expect(start(final!)).toBeCloseTo(180);
  });

  it('spreads thaiDistribute lines before a break per character from mode 15', () => {
    const [legacy] = lines(para([textRun('ab cd'), BR, textRun('x')], 'thaiDistribute'));
    expect(characters(legacy!)[1]!.xPt).toBeCloseTo(30);
    const [mode15] = lines(para([textRun('ab cd'), BR, textRun('x')], 'thaiDistribute'), MODE_15);
    expect(characters(mode15!)[1]!.xPt).toBeGreaterThan(30);
    expect(end(mode15!)).toBeCloseTo(380);
  });

  it('keeps the compression a mode-15 line accepted when a tab follows its gaps', () => {
    // A 128 pt band: 'a b c d e' (90 pt) + tab + 'foo' overflows and is
    // accepted by compressing the gaps before the tab.
    const narrow = para([textRun('a b c d e\tfoo'), BR, textRun('x')], 'both', { indentRight: 232 });
    expect(() => lines(narrow, MODE_15)).not.toThrow();
  });

  it('keeps a right-to-left line without gaps at its logical start', () => {
    const rtl = para([textRun('\u05d0\u05d1\u05d2'), BR, textRun('\u05d3')], 'both', { bidi: true } as Partial<DocParagraph>);
    const [broken] = lines(rtl);
    expect(end(broken!)).toBeCloseTo(380);
  });

  it('does not centre a mode-15 final distribute line that has a tab', () => {
    const [final] = lines(para([textRun('aa\tbbbb')], 'distribute'), MODE_15);
    expect(start(final!)).toBeCloseTo(20);
  });

  it('elongates no kashida before the last tab', () => {
    const [broken] = lines(para([textRun('\u0633\u0644\u0627\u0645\tbb cc'), BR, textRun('x')], 'lowKashida'));
    const tab = broken!.placements.find((placement) => placement.kind === 'tab')!;
    expect(tab.bounds!.xPt + tab.bounds!.widthPt).toBeCloseTo(92);
  });

  it('opens a character gap at a formatting seam but not inside a grapheme cluster', () => {
    const red = { color: 'FF0000' };
    const [seam] = lines(para([textRun('a'), textRun('b', red)], 'distribute'));
    expect(characters(seam!).map(({ xPt }) => xPt)).toEqual([20, 370].map((x) => expect.closeTo(x)));
    const [legacyThai] = lines(para([textRun('a'), textRun('b', red), BR, textRun('x')], 'thaiDistribute'), MODE_15);
    expect(characters(legacyThai!).map(({ xPt }) => xPt)).toEqual([20, 370].map((x) => expect.closeTo(x)));
    // 'a' and a combining acute in the next run are one cluster: the slack
    // goes between it, 'b' and 'c' only.
    const [mark] = lines(para([textRun('a'), textRun('\u0301bc', red)], 'distribute'));
    const placed = characters(mark!);
    expect(placed[1]!.xPt - placed[0]!.xPt).toBeCloseTo(10);
    expect(end(mark!)).toBeCloseTo(380);
  });

  it('keeps a fitText region and a grapheme cluster across runs whole, with full context', () => {
    const red = { color: 'FF0000' };
    const origins = (line: LineLayout): number[] => line.placements
      .filter((p): p is TextPlacement => p.kind === 'text' && p.text.trim() !== '')
      .map((p) => p.origin.xPt).sort((x, y) => x - y);
    // Two runs of one 40 pt fitText region stay together.
    const region = { fitTextVal: 800, fitTextId: 42 };
    const [fit] = lines(para([textRun('ab', region), textRun('cd', { ...region, ...red })], 'distribute'));
    expect(origins(fit!)).toEqual([20, 40].map((x) => expect.closeTo(x)));
    // Two flags in separate runs are two clusters: the gap opens.
    const [flags] = lines(para([textRun('\u{1F1FA}\u{1F1F8}'), textRun('\u{1F1E8}\u{1F1E6}', red)], 'distribute'));
    expect(origins(flags!)).toEqual([20, 360].map((x) => expect.closeTo(x)));
    // An emoji ZWJ sequence split across runs stays one cluster.
    const [zwj] = lines(para([textRun('\u{1F469}\u200d'), textRun('\u{1F680}ab', red)], 'distribute'));
    const [woman, rocket] = zwj!.placements.filter((p): p is TextPlacement => p.kind === 'text')
      .sort((x, y) => x.origin.xPt - y.origin.xPt);
    expect(rocket!.origin.xPt).toBeCloseTo(woman!.origin.xPt + woman!.advancePt);
  });

  it('keeps a right-to-left CJK line whose gaps stay closed at its logical start', () => {
    const rtl = para([textRun('\u3042\u3044\u3046\u3048'), BR, textRun('x')], 'distribute',
      { bidi: true } as Partial<DocParagraph>);
    const [broken] = lines(rtl, MODE_15);
    expect(end(broken!)).toBeCloseTo(380);
  });

  it('keeps a hanging fixed-pitch space out of the character spread', () => {
    const [line] = lines(para([textRun('ab'), textRun(' ', { fitTextVal: 200 })], 'distribute'));
    expect(characters(line!).map(({ xPt }) => xPt)).toEqual([20, 370].map((x) => expect.closeTo(x)));
  });

  it('keeps a right-to-left line whose gap stays closed at its logical start', () => {
    const rtl = para([textRun('\u05d0\u05d1\u05d2 \u05d3\u05d4\u05d5'), BR, textRun('\u05d7')], 'distribute',
      { bidi: true } as Partial<DocParagraph>);
    const [broken] = lines(rtl, MODE_15);
    expect(end(broken!)).toBeCloseTo(380);
    expect(start(broken!)).toBeGreaterThan(300);
  });

  it('keeps regional-indicator pairs whole across runs when spreading characters', () => {
    // '\u{1F1FA}' + red '\u{1F1F8}\u{1F1E8}\u{1F1E6}ab': the flags pair as US, CA.
    const [line] = lines(para([textRun('\u{1F1FA}'), textRun('\u{1F1F8}\u{1F1E8}\u{1F1E6}ab', { color: 'FF0000' })],
      'distribute'));
    const [us, flags] = line!.placements.filter((p): p is TextPlacement => p.kind === 'text')
      .sort((a, b) => a.origin.xPt - b.origin.xPt);
    // The seam inside the US pair stays closed, and no cut opens inside CA,
    // although a retained cluster starts there.
    expect(flags!.origin.xPt).toBeCloseTo(us!.origin.xPt + us!.advancePt);
    const [first, second] = flags!.clusters;
    expect(second!.offset.xPt).toBeCloseTo(first!.offset.xPt + first!.advancePt);
    const ab = line!.placements.filter((p): p is TextPlacement => p.kind === 'text' && p.text === 'ab')[0]!;
    const b = ab.clusters.at(-1)!;
    expect(ab.origin.xPt + b.offset.xPt + b.advancePt).toBeCloseTo(380);
  });

  it('gives a text-box paragraph no page or column break of its host', () => {
    const box = {
      type: 'shape', presetGeometry: 'rect', wrapMode: 'inline', widthPt: 300, heightPt: 100,
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0, textAnchor: 't',
      textBlocks: [{ text: 'aa bb', fontSizePt: 20, fontFamily: 'NotInMetrics', alignment: 'both',
        runs: [{ text: 'aa bb', fontSizePt: 20, fontFamily: 'NotInMetrics' }] }],
    } as unknown as DocParagraph['runs'][number];
    const host = para([box], 'left', { endsAtHardBreak: true });
    const page = layoutBodyModel([host], section, makeCtx(), {}, undefined).pages[0]!;
    const found: TextPlacement[] = [];
    const visit = (node: unknown): void => {
      if (!node || typeof node !== 'object') return;
      const record = node as Record<string, unknown>;
      if (record.kind === 'text' && typeof record.text === 'string' && record.text.includes('bb')) {
        found.push(record as unknown as TextPlacement);
      }
      Object.values(record).forEach(visit);
    };
    visit(page);
    expect(found.length).toBeGreaterThan(0);
    const bb = found[0]!;
    const aa = (() => {
      const all: TextPlacement[] = [];
      const collect = (node: unknown): void => {
        if (!node || typeof node !== 'object') return;
        const record = node as Record<string, unknown>;
        if (record.kind === 'text' && typeof record.text === 'string' && record.text.startsWith('aa')) {
          all.push(record as unknown as TextPlacement);
        }
        Object.values(record).forEach(collect);
      };
      collect(page);
      return all[0]!;
    })();
    // A single final line: 'bb' follows 'aa ' at its natural place, unstretched.
    const bbX = bb.text.startsWith('bb') ? bb.origin.xPt : aa.origin.xPt + 30;
    expect(bbX - aa.origin.xPt).toBeCloseTo(30);
  });

  it('justifies a line ended by a page or column break before mode 15 only', () => {
    const piece = para([textRun('aaaa bbbb')], 'both', { endsAtHardBreak: true });
    expect(end(lines(piece)[0]!)).toBeCloseTo(380);
    expect(end(lines(piece, MODE_15)[0]!)).toBeCloseTo(110);
    expect(end(lines(piece, { doNotExpandShiftReturn: true })[0]!)).toBeCloseTo(380);
  });
});
