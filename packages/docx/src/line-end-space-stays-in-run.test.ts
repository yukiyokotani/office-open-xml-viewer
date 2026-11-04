import { describe, it, expect } from 'vitest';
import { sourceOwnedTextPlacements } from './layout/text-source-ownership.js';
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
    const words = line!.filter(p => p.text.trim()).map((p) => ({ text: p.text, x: p.bounds.xPt }));
    expect(words.map((w) => w.text.trim())).toEqual(['קלם', 'כתב', 'Xray']);
    // The retained source spaces occupy the physical slots between words.
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

  it('retains separately formatted whitespace through a many-run layout', () => {
    const spaces = Array.from({ length: 400 }, (_, index) =>
      textRun(' ', { rtl: true, color: index % 2 === 0 ? 'FF0000' : '00FF00' }));
    const result = lines(para([textRun('abc '), textRun('אבג', { rtl: true }), ...spaces], 'left'),
      { compatibilityMode: 15 });
    const placed = result.flatMap(texts);
    expect(placed.map(p => p.text).join('').replace(/\s/g, '')).toBe('abcאבג');
    expect(placed.reduce((sum, p) => sum + p.text.length, 0)).toBe(407);
    for (const placement of placed) {
      expect(placement.bounds.widthPt).toBeGreaterThanOrEqual(0);
      for (const cluster of placement.clusters) expect(cluster.advancePt).toBeGreaterThanOrEqual(0);
    }
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
  it('keeps visible glyphs inside the band when the preceding decorated run cannot absorb all spaces', () => {
    const size = { fontSize: 12 };
    const [line] = lines(para([
      textRun('ab', size), textRun('c ', { ...size, underline: true, background: 'FFFF00' }),
      textRun('אבג ABC    ', { ...size, rtl: true }),
    ], 'both', { indentRight: 291, defaultFontSize: 12 } as Partial<DocParagraph>), { compatibilityMode: 15 });
    const right = Math.max(...texts(line!).filter(p => p.text.trim()).map(p => p.bounds.xPt + p.bounds.widthPt));
    expect(right).toBeLessThanOrEqual(89 + 1e-6);
  });

  it('compresses U+0020 without taking advance from a preceding nonbreaking space', () => {
    const size = { fontSize: 12 };
    const [line] = lines(para([
      textRun('ab', size), textRun('c ', { ...size, underline: true }),
      textRun('אבג ABC\u00a0    ', { ...size, rtl: true }),
    ], 'both', { indentRight: 285, defaultFontSize: 12 } as Partial<DocParagraph>), { compatibilityMode: 15 });
    const label = texts(line!).find(p => p.text.startsWith('ABC'))!;
    const nbsp = label.clusters.find(cluster =>
      label.text.slice(cluster.range.start - label.range.start, cluster.range.end - label.range.start) === '\u00a0')!;
    expect(nbsp.advancePt).toBe(6);
    expect(label.clusters.reduce((sum, cluster) => sum + cluster.advancePt, 0)).toBeCloseTo(label.advancePt, 6);
    expect(label.trailingSpaceCompressionPt).toBeGreaterThan(0);
  });

  it('places word-space expansion on the physical leading side of RTL-painted slices', () => {
    const [line] = lines(para([
      textRun('Xray ', { rtl: true }), textRun('כתב', { rtl: true }),
      textRun(' ', { rtl: true }), textRun('Yoke', { rtl: true }),
    ], 'distribute'),
      { compatibilityMode: 15 });
    const placed = texts(line!);
    expect(placed.filter(p => p.text.trim()).map(p => p.text.trim())).toEqual(['Yoke', 'כתב', 'Xray']);
    // Natural content takes 130pt; each of the two source spaces gets 115pt
    // of the 230pt slack. RTL-owned spaces precede their word physically.
    expect(placed.find(p => p.text.trim() === 'כתב')!.origin.xPt).toBeCloseTo(185, 6);
    expect(placed.find(p => p.text.trim() === 'Xray')!.origin.xPt).toBeCloseTo(340, 6);
    const arabic = placed.find(p => p.text.trim() === 'כתב')!;
    const owners = sourceOwnedTextPlacements(arabic);
    const space = owners.find(p => p.text === ' ')!;
    const word = owners.find(p => p.text === 'כתב')!;
    expect(space.bounds.widthPt).toBeCloseTo(125, 6);
    expect(space.trailingSpaceCompressionPt).toBeCloseTo(-115, 6);
    expect(word.bounds.xPt).toBeCloseTo(185, 6);
    expect(word.origin.xPt).toBeCloseTo(185, 6);
    for (const placement of placed) {
      expect(placement.bounds.widthPt).toBeGreaterThanOrEqual(0);
      for (const cluster of placement.clusters) expect(cluster.advancePt).toBeGreaterThanOrEqual(0);
    }

  });

  it('counts opposite-run spaces at explicit line ends across source seams in mode 15', () => {
    const content = `${WORDS} ${WORDS} `;
    const variants = [
      [textRun(content, { rtl: true })],
      [textRun(content.slice(0, -1), { rtl: true }), textRun(' ', { rtl: true })],
      [textRun(content.slice(0, -3), { rtl: true }), textRun(content.slice(-3), { rtl: true, bold: true })],
      [textRun(content.replaceAll('א', '😀'), { rtl: true })],
    ];
    for (const runs of variants) {
      const body = para(runs, 'left', narrow);
      expect(lines(body, { compatibilityMode: 14 })).toHaveLength(2);
      const current = lines(body, { compatibilityMode: 15 });
      expect(current).toHaveLength(3);
      // A source-interior soft wrap still hangs: only the explicit end counts.
      const first = texts(current[0]!).sort((a, b) => a.range.start - b.range.start)
        .map(p => p.text).join('').trim().split(/ +/u);
      expect(first).toHaveLength(4);
    }
  });

  it('counts an unmarked Latin run space against a bidi paragraph end', () => {
    const latin = 'abcde fghij klmno pqrst';
    const body = para([textRun(`${latin} ${latin} `)], 'left', { ...narrow, bidi: true });
    expect(lines(body, { compatibilityMode: 14 })).toHaveLength(2);
    expect(lines(body, { compatibilityMode: 15 })).toHaveLength(3);
  });

  it('counts source spaces before a manual break and distinguishes mixed-direction suffixes', () => {
    const content = `${WORDS} ${WORDS} `;
    const br = { type: 'break', breakType: 'line' } as DocParagraph['runs'][number];
    const broken = lines(para([textRun(content, { rtl: true }), br, textRun('tail')],
      'left', narrow), { compatibilityMode: 15 });
    expect(broken).toHaveLength(4);
    const three = lines(para([textRun(content + '  ', { rtl: true })], 'left',
      { indentRight: 105 }), { compatibilityMode: 15 });
    expect(three).toHaveLength(3);
    const mixed = lines(para([
      textRun(content + ' ', { rtl: true }), textRun(' ', { rtl: false }),
    ], 'left', { indentRight: 105 }), { compatibilityMode: 15 });
    expect(mixed).toHaveLength(2);
  });

  it('hangs the physical leading space on soft-wrapped RTL distribute lines', () => {
    const [first] = lines(para([textRun(`${WORDS} ${WORDS}`, { rtl: true })],
      'distribute', { bidi: true }), { compatibilityMode: 15 });
    const ink = texts(first!).filter(p => p.text.trim());
    expect(Math.min(...ink.map(p => p.origin.xPt))).toBeCloseTo(20, 6);
    expect(Math.max(...ink.map(p => p.origin.xPt + p.text.trim().length * 10)))
      .toBeCloseTo(380, 6);
  });

  it('expands a logically interior label space even when it paints at the physical right edge', () => {
    const [line] = lines(para([
      textRun('Xray '), textRun('כתב קלם בית', { rtl: true }),
    ], 'distribute', { bidi: true }), { compatibilityMode: 15 });
    const placed = texts(line!);
    const ink = placed.filter(p => p.text.trim());
    expect(ink.map(p => p.text.trim())).toEqual(['בית', 'קלם', 'כתב', 'Xray']);
    // Three logical source spaces share 200pt. The label's space follows its
    // glyphs physically, so the visible edge ends before the allocated edge.
    expect(ink.map(p => p.origin.xPt)).toEqual([
      expect.closeTo(20, 6), expect.closeTo(126.666666667, 6),
      expect.closeTo(233.333333333, 6), expect.closeTo(263.333333333, 6),
    ]);
    const label = ink.find(p => p.text.trim() === 'Xray')!;
    const closing = label.clusters.find(c => c.range.end === label.range.end)!;
    expect(label.bounds.xPt + label.bounds.widthPt).toBeCloseTo(380, 6);
    expect(closing.advancePt).toBeCloseTo(76.666666667, 6);
    expect(label.clusters.reduce((sum, c) => sum + c.advancePt, 0)).toBeCloseTo(label.advancePt, 6);
  });

  it('retains expansion inside a consecutive trailing-space sequence', () => {
    const [line] = lines(para([
      textRun('Xray ', { rtl: true }), textRun('כתב', { rtl: true }),
      textRun('   ', { rtl: true }), textRun('Yoke', { rtl: true }),
    ], 'distribute'),
      { compatibilityMode: 15 });
    const placed = texts(line!);
    // Four U+0020 opportunities share 210pt: all three spaces before the
    // Hebrew word physically retain their 52.5pt expansion.
    expect(placed.find(p => p.text.trim() === 'כתב')!.origin.xPt).toBeCloseTo(247.5, 6);
    expect(placed.find(p => p.text.trim() === 'Xray')!.origin.xPt).toBeCloseTo(340, 6);
    const hebrew = placed.find(p => p.text.trim() === 'כתב')!;
    const spaceOwner = sourceOwnedTextPlacements(hebrew).find(p => p.text === '   ')!;
    expect(spaceOwner.bounds.widthPt).toBeCloseTo(187.5, 6);
    expect(spaceOwner.trailingSpaceCompressionPt).toBeCloseTo(-157.5, 6);
    expect(hebrew.clusters.reduce((sum, cluster) => sum + cluster.advancePt, 0))
      .toBeCloseTo(hebrew.advancePt, 6);

  });

  it('hangs the physical end space of a non-rtl run on a soft-wrapped bidi line', () => {
    const [first] = lines(para([textRun(`${WORDS} ${WORDS}`)], 'left',
      { ...narrow, bidi: true } as Partial<DocParagraph>), { compatibilityMode: 14 });
    const right = Math.max(...texts(first!).map(p => p.origin.xPt
      + p.clusters.filter(c => c.range.start < p.range.start + p.text.trimEnd().length)
        .reduce((sum, c) => sum + c.advancePt, 0)));
    expect(right).toBeCloseTo(380, 6);
  });

  it('retains differently formatted closing spaces outside their content-level block', () => {
    const tailWord = WORDS.split(' ').at(-1)!;
    const [first] = lines(para([
      textRun(WORDS.slice(0, -tailWord.length)),
      textRun(`${tailWord}   `, { fontSize: 12, background: 'FFFF00' }),
      textRun(WORDS),
    ], 'left', { bidi: true, indentRight: 95 } as Partial<DocParagraph>), { compatibilityMode: 14 });
    const placed = texts(first!);
    const closing = placed.find(p => p.text === '   ')!;
    expect(closing.bounds.xPt).toBeCloseTo(380, 6);
    expect(closing.bounds.widthPt).toBeCloseTo(18, 6);
    expect(closing.fontSizePt).toBe(12);
    expect(closing.highlightBounds?.xPt).toBeCloseTo(380, 6);
    expect(closing.highlightBounds?.widthPt).toBeCloseTo(18, 6);
    expect([...placed].sort((a, b) => a.range.start - b.range.start).map(p => p.text).join(''))
      .toBe(`${WORDS}   `);
  });

});
