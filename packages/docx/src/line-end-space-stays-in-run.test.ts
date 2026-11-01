import { describe, it, expect, vi } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocSettings, DocxTextRun, SectionProps } from './types';

// WORD_LINE_END_SPACE_STAYS_IN_RUN (blk-0898-bidi-fit, Windows Word
// 16.0.20527): a line-end U+0020 inside a run of the opposite direction stays
// inside that run, at the start side of the line.

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

function texts(line: LineLayout): TextPlacement[] {
  return line.placements.filter((p): p is TextPlacement => p.kind === 'text')
    .sort((a, b) => a.bounds.xPt - b.bounds.xPt);
}


// 23 visible characters (230 pt) and a closing space (10 pt) in a 235 pt band.
const WORDS = 'אבגדה וזחטי כלמנס עפצקר';
const narrow = { indentRight: 125 } as Partial<DocParagraph>;

function placements(body: BodyElement, settings?: DocSettings): TextPlacement[][] {
  return lines(body, settings).map((line) => texts(line));
}

describe('line-end space in a run of the opposite direction', () => {
  it('narrows the overflowing final space at the start side below mode 15', () => {
    // Mode 14 keeps the line; only the 5 pt overflow of the 10 pt space hangs
    // past the start edge (A0-left-L: ink from 57.1pt, not 60.0pt).
    const body = para([textRun(`${WORDS} `, { rtl: true })], 'left', narrow);
    const [line] = placements(body, { compatibilityMode: 14 });
    expect(line![0]!.origin.xPt).toBeCloseTo(25, 6);
  });

  it('paints the final space at the start side of the line, inside the rtl run', () => {
    const body = para([textRun('אבג דהו ', { rtl: true })], 'left');
    const [line] = placements(body, { compatibilityMode: 15 });
    // The logically last word is visually first; its closing space stays on
    // its left, so the glyphs start one space (10 pt) after the band start.
    const first = line![0]!;
    expect(first.text).toBe('דהו ');
    expect(first.bounds.xPt).toBe(20);
    expect(first.paintOps[0]!.direction).toBe('rtl');
    expect(first.origin.xPt).toBe(30);
  });

  it('hangs a soft-wrapped line-end space of an rtl run past the start edge', () => {
    // The space stays in its run, drawn at the left of the line (no UAX #9
    // L1); Word starts the first line's ink at the band start (A0-left-L).
    const body = para([textRun(`${WORDS} ${WORDS}`, { rtl: true })], 'left', narrow);
    const [first] = placements(body, { compatibilityMode: 15 });
    const leftmost = first![0]!;
    expect(leftmost.paintOps[0]!.direction).toBe('rtl');
    expect(leftmost.text.endsWith(' ')).toBe(true);
    expect(leftmost.origin.xPt).toBe(20);
    expect(leftmost.bounds.xPt).toBe(10);
  });

  it('puts the space between an Arabic word and following Latin inside a non-rtl run (e02)', () => {
    // bidi paragraph, one non-rtl run "A1 A2 Xray": Word draws A2 A1 Xray, one space between each.
    const body = para([textRun('כתב קלם Xray')], 'left', { bidi: true } as Partial<DocParagraph>);
    const [line] = placements(body);
    const words = line!.map((p) => ({ text: p.text, x: p.bounds.xPt }));
    expect(words.map((w) => w.text.trim())).toEqual(['קלם', 'כתב', 'Xray']);
    // Each slice keeps its own trailing space at its right: 4 + 4 + 4 characters.
    expect(words.map((w) => w.x)).toEqual([words[0]!.x, words[0]!.x + 40, words[0]!.x + 80]);
  });

  it('keeps line-level UAX #9 order on a right-to-left line spread per character (pre-15 distribute)', () => {
    // The character-spread gap model has not adopted WORD_RUN_DIRECTION_ISOLATE.
    const bidi = { bidi: true } as Partial<DocParagraph>;
    const order = (alignment: DocParagraph['alignment']) => placements(
      para([textRun('Xray כתב קלם')], alignment, bidi), { compatibilityMode: 14 },
    )[0]!.map((p) => p.text.trim());
    expect(order('left')).toEqual(['Xray', 'קלם', 'כתב']);
    expect(order('distribute')).toEqual(['קלם', 'כתב', 'Xray']);
  });

  it('resolves many separately formatted line-end spaces in linear time', () => {
    // An LTR word, rtl Hebrew and 400 separately coloured rtl spaces: the
    // spaces are drawn inside the line, each an interior line-end space.
    const spaces = Array.from({ length: 400 }, (_, index) =>
      textRun(' ', { rtl: true, color: index % 2 === 0 ? 'FF0000' : '00FF00' }));
    const body = para([textRun('abc '), textRun('אבג', { rtl: true }), ...spaces], 'left');
    const spy = vi.spyOn(Array.prototype, 'indexOf');
    lines(body, { compatibilityMode: 15 });
    const calls = spy.mock.contexts.filter((context, index) =>
      (context as unknown[]).length >= 400 && typeof spy.mock.calls[index]![0] === 'number').length;
    spy.mockRestore();
    expect(calls).toBeLessThan(10);
  });

  it('absorbs interior line-end spaces with one copy of the stretch plan', () => {
    const run = (count: number) => {
      const spaces = Array.from({ length: count }, (_, index) =>
        textRun(' ', { rtl: true, color: index % 2 === 0 ? 'FF0000' : '00FF00' }));
      const body = para([textRun('abc '), textRun('אבג', { rtl: true }), ...spaces], 'left');
      const set = vi.spyOn(Map.prototype, 'set');
      lines(body, { compatibilityMode: 15 });
      const calls = set.mock.calls.length;
      set.mockRestore();
      return calls;
    };
    // Linear: doubling the spaces at most doubles the map writes (plus slack);
    // copying the plan per absorbed space grows them about fourfold.
    expect(run(800)).toBeLessThan(2.5 * run(400));
  });

  it('keeps the visible text of a mode-15 justified line inside the band when its line-end spaces are drawn inside', () => {
    // An 85 pt band (20 to 105): the rtl spaces after the Hebrew stay in
    // their run, drawn between "abc " and the Hebrew; the line must still
    // end inside the band.
    const hebrewEnd = (runs: DocParagraph['runs']) => {
      const [first] = placements(para(runs, 'both', { indentRight: 275 } as Partial<DocParagraph>),
        { compatibilityMode: 15 });
      const hebrew = first!.find((p) => p.text.includes('אבג'))!;
      return hebrew.bounds.xPt + hebrew.bounds.widthPt;
    };
    const base = [textRun('abc '), textRun('אבג', { rtl: true }), textRun('   ', { rtl: true, color: 'FF0000' })];
    expect(hebrewEnd([...base, textRun('next')])).toBeLessThanOrEqual(105 + 1e-6);
    // With an edge-hanging space as well, the interior ones are still absorbed.
    expect(hebrewEnd([...base, textRun(' ', { color: '00FF00' }), textRun('next')]))
      .toBeLessThanOrEqual(105 + 1e-6);
  });

  it('counts absorbed interior spaces once in the drawn line width', () => {
    // Ten rtl spaces drawn inside an 85 pt mode-15 justified line: the line's
    // width spans its drawn content and never goes negative.
    for (const count of [3, 10]) {
      const [line] = lines(para([
        textRun('abc '), textRun('אבג', { rtl: true }),
        textRun(' '.repeat(count), { rtl: true, color: 'FF0000' }),
      ], 'both', { indentRight: 275 } as Partial<DocParagraph>), { compatibilityMode: 15 });
      const right = Math.max(...texts(line!).filter((p) => p.text.trim() !== '')
        .map((p) => p.bounds.xPt + p.bounds.widthPt));
      expect(line!.bounds.widthPt).toBeGreaterThanOrEqual(0);
      expect(line!.bounds.xPt + line!.bounds.widthPt).toBeGreaterThanOrEqual(right - 1e-6);
    }
  });

  it('never absorbs more than a short decorated neighbour holds', () => {
    // 12 pt text in a 42 pt band; the spaces share the rtl word's run, and
    // absorption must not give the highlighted "c " a negative extent.
    const size = { fontSize: 12 };
    // A single closing space, and a double one compressed inside its segment.
    for (const [decorated, indentRight] of [['c ', 318], ['c  ', 315]] as const) {
      const [line] = lines(para([
        textRun('ab', size), textRun(decorated, { ...size, underline: true, background: 'FFFF00' }),
        textRun('אבג   ', { ...size, rtl: true }),
      ], 'both', { indentRight, defaultFontSize: 12 } as Partial<DocParagraph>), { compatibilityMode: 15 });
      const widths: number[] = [];
      JSON.stringify(line, (key, value) => {
        if (key === 'widthPt' && typeof value === 'number') widths.push(value);
        return value;
      });
      expect(Math.min(...widths)).toBeGreaterThanOrEqual(0);
    }
  });
});
