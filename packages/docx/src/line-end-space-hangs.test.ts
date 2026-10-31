import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocSettings, DocxTextRun, SectionProps } from './types';

// WORD_LINE_END_SPACE_HANGS: the U+0020 that ends a soft-wrapped LTR line
// stands outside justification and centre/right alignment. Windows Word ends
// every justified line flush at the margin in compatibility modes unset, 11,
// 12 and 14 (mode 15 already fitted the visible width), and places
// centre/right-aligned wrapped lines by their visible text.

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

function para(runs: DocParagraph['runs'], alignment: DocParagraph['alignment']): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 20, defaultFontFamily: 'NotInMetrics',
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

const visibleStart = (line: LineLayout): number =>
  texts(line).find((p) => p.text.trim() !== '')!.bounds.xPt;

/** Right edge of the line's visible text (each character 10 pt). */
function visibleEnd(line: LineLayout): number {
  const last = texts(line).filter((p) => p.text.trim() !== '').at(-1)!;
  const trailing = last.text.length - last.text.replace(/[ \u3000]+$/, '').length;
  return last.bounds.xPt + last.advancePt - trailing * 10;
}

// Seven 4-letter words (340 pt visible, 350 pt with the line-end space) fit
// the 360 pt band; the eighth wraps.
const WORDS = 'aaaa bbbb cccc dddd eeee ffff gggg hhhh';

describe('WORD_LINE_END_SPACE_HANGS', () => {
  it('justifies a wrapped line flush to the band edge, the line-end space hanging', () => {
    const [first, last] = lines(para([textRun(WORDS)], 'both'));
    expect(texts(first!).map((p) => p.text).join('')).toBe('aaaa bbbb cccc dddd eeee ffff gggg ');
    expect(visibleStart(first!)).toBeCloseTo(20, 6);
    expect(visibleEnd(first!)).toBeCloseTo(380, 6);
    // The final line is not justified.
    expect(visibleStart(last!)).toBeCloseTo(20, 6);
  });

  it('centres and right-aligns a wrapped line by its visible text', () => {
    const [centred, centredLast] = lines(para([textRun(WORDS)], 'center'));
    expect(visibleStart(centred!)).toBeCloseTo(20 + (360 - 340) / 2, 6);
    expect(visibleStart(centredLast!)).toBeCloseTo(20 + (360 - 40) / 2, 6);
    const [right, rightLast] = lines(para([textRun(WORDS)], 'right'));
    expect(visibleEnd(right!)).toBeCloseTo(380, 6);
    expect(visibleEnd(rightLast!)).toBeCloseTo(380, 6);
  });

  it('looks through a zero-width anchor host after the line-end space', () => {
    const host = { type: 'anchorHost', fontSize: 20, fontFamily: 'NotInMetrics', bold: false, italic: false } as DocParagraph['runs'][number];
    const runs = [textRun('aaaa bbbb cccc dddd eeee ffff gggg '), host, textRun('hhhh')];
    const [right] = lines(para(runs, 'right'));
    expect(visibleEnd(right!)).toBeCloseTo(380, 6);
    const [centred] = lines(para(runs, 'center'));
    expect(visibleStart(centred!)).toBeCloseTo(30, 6);
    const [justified] = lines(para(runs, 'both'));
    expect(visibleEnd(justified!)).toBeCloseTo(380, 6);
    // Hosts between trailing space runs are looked through too.
    const interleaved = [textRun('aaaa bbbb cccc dddd eeee ffff gggg '), host, textRun(' ', { bold: true }), textRun('hhhh')];
    const [spread] = lines(para(interleaved, 'both'));
    expect(visibleEnd(spread!)).toBeCloseTo(380, 6);
  });

  it('excludes the hanging space from mode-15 CJK expansion behind an anchor host', () => {
    const host = { type: 'anchorHost', fontSize: 20, fontFamily: 'NotInMetrics', bold: false, italic: false } as DocParagraph['runs'][number];
    const words = 'ああ bbbb cccc dddd eeee ffff gggg ';
    for (const runs of [
      [textRun(words), host, textRun('hhhhhhhh')],
      [textRun(words), textRun('hhhhhhhh')],
      // A host between the last ideograph and the hanging space.
      [textRun('ああ bbbb cccc dddd eeee ffff ああ'), host, textRun(' '), textRun('hhhhhhhh')],
    ]) {
      const [line] = lines(para(runs, 'both'), { compatibilityMode: 15 });
      expect(visibleEnd(line!)).toBeCloseTo(380, 6);
    }
  });

  it('leaves right-to-left lines to their previous geometry', () => {
    // A line with right-to-left text is outside the rule: its line-end space
    // and a trailing anchor host keep their cells, so seven gaps share the
    // 10 pt of slack (350 pt natural width) as before the rule.
    const host = { type: 'anchorHost', fontSize: 20, fontFamily: 'NotInMetrics', bold: false, italic: false } as DocParagraph['runs'][number];
    const [line] = lines(para([textRun('אאאא bbbb cccc dddd eeee ffff gggg '), host, textRun('hhhh')], 'both'));
    const last = texts(line!).find((p) => p.text.startsWith('gggg'))!;
    expect(last.bounds.xPt).toBeCloseTo(20 + 6 * 40 + 6 * (10 + 10 / 7), 6);
  });

  it('keeps the line-end spaces of a wordWrap-off paragraph in the line', () => {
    // §17.3.1.45: with wordWrap off Word counts every trailing space (Windows
    // Word controls: right-aligned lines end 3, 9 and 12 pt short of the margin
    // for one space, three spaces and an ideographic space).
    for (const [tail, widthPt] of [[' ', 10], ['   ', 30], ['\u3000', 10]] as const) {
      const p = { ...(para([textRun(`aaaa${tail}`)], 'right') as DocParagraph), wordWrap: false } as unknown as BodyElement;
      const [line] = lines(p);
      expect(visibleEnd(line!)).toBeCloseTo(380 - widthPt, 6);
    }
  });

  it('keeps hanging spaces decorated within the band under ulTrailSpace', () => {
    // §17.15.3.8: with ulTrailSpace Windows Word decorates trailing spaces,
    // clipped at the right edge of the line's band; alignment is unchanged.
    const decorated = { underline: true, highlight: 'yellow' } as Partial<DocxTextRun>;
    const end = (line: LineLayout) => Math.max(...texts(line).flatMap((p) => [
      ...(p.decorations ?? []).map((d) => d.to.xPt),
      ...(p.highlightFragments ?? []).map((f) => f.rect.xPt + f.rect.widthPt),
    ]));
    // Right-aligned: the hanging space lies past the band edge, nothing shows.
    const [right] = lines(para([textRun(WORDS, decorated)], 'right'), { ulTrailSpace: true });
    expect(visibleEnd(right!)).toBeCloseTo(380, 6);
    expect(end(right!)).toBeCloseTo(380, 6);
    // A space-only run wholly past the edge keeps no decoration at all.
    const wavy = { underline: true, underlineStyle: 'wave' } as Partial<DocxTextRun>;
    const [tail] = lines(para([textRun('aaaa '), textRun('\u3000', wavy)], 'right'), { ulTrailSpace: true });
    expect(texts(tail!).at(-1)!.decorations ?? []).toEqual([]);
    // Centred final line: the space is decorated inside the band.
    const [centred] = lines(para([textRun('aaaa\u3000 ', decorated)], 'center'), { ulTrailSpace: true });
    expect(visibleStart(centred!)).toBeCloseTo(20 + (360 - 40) / 2, 6);
    expect(end(centred!)).toBeCloseTo(20 + (360 - 40) / 2 + 40 + 20, 6);
  });

  it('paints no decoration over the hanging space', () => {
    // Windows Word ends underline, strike-through, highlight and shading at the
    // visible text of a soft-wrapped line.
    const decorated = { underline: true, highlight: 'yellow' } as Partial<DocxTextRun>;
    for (const alignment of ['both', 'right'] as const) {
      for (const extra of [decorated, { strikethrough: true }] as Partial<DocxTextRun>[]) {
        const [first] = lines(para([textRun(WORDS, extra)], alignment));
        const last = texts(first!).at(-1)!;
        expect(last.text.endsWith(' ')).toBe(true);
        // Adjacent decorated runs merge into one rule; it ends at the text.
        const ends = texts(first!).flatMap((p) => p.decorations ?? []).map((d) => d.to.xPt);
        expect(ends.length).toBeGreaterThan(0);
        expect(Math.max(...ends)).toBeCloseTo(380, 6);
        if (extra.highlight) {
          const fill = last.highlightFragments![0]!.rect;
          expect(fill.xPt + fill.widthPt).toBeCloseTo(380, 6);
        }
      }
    }
    // A hanging U+3000 is not decorated either.
    const [ideographic] = lines(para([textRun('aaaa\u3000', decorated)], 'center'));
    const fills = texts(ideographic!).flatMap((p) => p.highlightFragments ?? []);
    expect(Math.max(...fills.map((f) => f.rect.xPt + f.rect.widthPt))).toBeCloseTo(visibleEnd(ideographic!), 6);
    // A line-end space in its own decorated run carries none.
    const runs = [textRun('aaaa bbbb cccc dddd eeee ffff gggg'), textRun(' ', decorated), textRun('hhhh')];
    const [first] = lines(para(runs, 'right'));
    const space = texts(first!).at(-1)!;
    expect(space.text).toBe(' ');
    expect(space.decorations ?? []).toEqual([]);
    // An empty list: an absent one would make paint fall back to the bounds.
    expect(space.highlightFragments).toEqual([]);
  });

  it('hangs paragraph-final spaces and ideographic spaces too', () => {
    for (const text of ['aaaa bbbb ', 'aaaa bbbb   ', 'aaaa\u3000bbbb\u3000', 'aaaa bbbb \u3000 ']) {
      const [right] = lines(para([textRun(text)], 'right'));
      expect(visibleEnd(right!)).toBeCloseTo(380, 6);
      const [centred] = lines(para([textRun(text)], 'center'));
      expect(visibleStart(centred!)).toBeCloseTo(20 + (360 - 90) / 2, 6);
    }
    // Ideographic-space separators wrap and hang like U+0020.
    const [wrapped] = lines(para([textRun(WORDS.replaceAll(' ', '\u3000'))], 'right'));
    expect(visibleEnd(wrapped!)).toBeCloseTo(380, 6);
  });

  it('hangs the spaces before a manual break', () => {
    const br = { type: 'break', breakType: 'line' } as DocParagraph['runs'][number];
    for (const tail of ['aaaa ', 'aaaa   ', 'aaaa\u3000']) {
      const [right] = lines(para([textRun(tail), br, textRun('bbbb')], 'right'));
      expect(visibleEnd(right!)).toBeCloseTo(380, 6);
      const [centred] = lines(para([textRun(tail), br, textRun('bbbb')], 'center'));
      expect(visibleStart(centred!)).toBeCloseTo(20 + (360 - 40) / 2, 6);
    }
  });
});
