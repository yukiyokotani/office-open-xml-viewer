import { describe, expect, it } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import type { DocumentLayout } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, DocxTextRun, SectionProps } from './types.js';

// WORD_KERNED_SPACE_PAIRS / WORD_KERN_THRESHOLD_AUTHORITY. The synthetic face
// has pair adjustments only on pairs that contain U+0020 (and AV), so a line
// tightens exactly by the pairs the layout acquires from it.
const SCALAR: Record<string, number> = { x: 6, A: 9, B: 8, V: 9, ' ': 3, '’': 4 };
const PAIRS: Record<string, number> = { ' A': -0.75, 'A ': -0.75, '’ ': -0.5, AV: -1 };

function context(): CanvasRenderingContext2D {
  const ctx = { font: '12px Arial', fontKerning: 'auto', letterSpacing: '0px',
    measureText(text: string) {
      const size = Number(/([\d.]+)px/u.exec(ctx.font)?.[1] ?? 12);
      const chars = [...text];
      let units = chars.reduce((sum, c) => sum + (SCALAR[c] ?? 7), 0);
      if (ctx.fontKerning === 'normal') {
        for (let i = 1; i < chars.length; i++) units += PAIRS[chars[i - 1]! + chars[i]!] ?? 0;
      }
      const width = units * size / 12;
      return { width, fontBoundingBoxAscent: size * .8, fontBoundingBoxDescent: size * .2,
        actualBoundingBoxLeft: 0, actualBoundingBoxRight: width,
        actualBoundingBoxAscent: size * .8, actualBoundingBoxDescent: size * .2 } as TextMetrics;
    } };
  return ctx as unknown as CanvasRenderingContext2D;
}

const borders = { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
function run(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return { type: 'text', text, fontSize: 12, fontFamily: 'Arial',
    bold: false, italic: false, underline: false, strikethrough: false, color: null,
    isLink: false, background: null, vertAlign: null, hyperlink: null, ...extra } as DocxTextRun;
}
function layout(runs: DocxTextRun[], width: number, compatibilityMode?: number,
  alignment: DocParagraph['alignment'] = 'left'): DocumentLayout {
  const paragraph: DocParagraph = { runs, alignment, indentLeft: 0, indentRight: 0,
    indentFirst: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops: [], defaultFontSize: 12, defaultFontFamily: 'Arial', widowControl: false } as unknown as DocParagraph;
  const body: BodyElement[] = [{ type: 'paragraph', ...paragraph }];
  const doc = { body, settings: compatibilityMode === undefined ? {} : { compatibilityMode },
    section: { pageWidth: width, pageHeight: 1000, marginTop: 0, marginRight: 0, marginBottom: 0,
      marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false,
      evenAndOddHeaders: false } as SectionProps,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {}, footnotes: [] } as unknown as DocxDocumentModel;
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
}
/** Painted glyph origins of each line, keyed by the painted text. */
function lines(result: DocumentLayout) {
  const out: { text: string; x: number; advance: number }[][] = [];
  const visit = (value: unknown, seen = new WeakSet<object>()): void => {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    if ('kind' in value && value.kind === 'paragraph' && 'lines' in value) {
      for (const line of (value as { lines: readonly { placements: readonly unknown[] }[] }).lines) {
        out.push(line.placements.flatMap((placement) => {
          const p = placement as { kind: string; origin: { xPt: number }; advancePt: number;
            paintOps: readonly { text: string; offset: { xPt: number } }[] };
          return p.kind === 'text'
            ? p.paintOps.map(op => ({ text: op.text, x: p.origin.xPt + op.offset.xPt, advance: p.advancePt }))
            : [];
        }));
      }
      return;
    }
    for (const [key, child] of Object.entries(value)) if (key !== 'source') {
      if (Array.isArray(child)) child.forEach(c => visit(c, seen)); else visit(child, seen);
    }
  };
  result.pages.forEach(page => visit(page));
  return out;
}
const SPAN_TEXT = 'x A x A x B';
/** Left edge of the final glyph: every preceding advance and seam pair. */
const lastGlyphX = (line: { text: string; x: number }[]) => line.at(-1)!.x;

describe('kerned U+0020 pairs (WORD_KERNED_SPACE_PAIRS)', () => {
  it.each([undefined, 15])('charges the space+glyph pair across word pieces (mode %s)', (mode) => {
    const [line] = lines(layout([run(SPAN_TEXT, { kerning: 2 })], 400, mode));
    expect(line!.map(p => p.text).join('')).toBe(SPAN_TEXT);
    // x(6) sp(3) A(9) sp(3) x(6) sp(3) A(9) sp(3) x(6) sp(3) = 51 natural,
    // two A+space and two space+A pairs at -0.75 each.
    expect(lastGlyphX(line!)).toBeCloseTo(51 - 4 * 0.75, 6);
  });

  it('keeps every pair, including space pairs, off with an explicit zero threshold and no mode', () => {
    const [line] = lines(layout([run(`${SPAN_TEXT} AV`, { kerning: 0 })], 400));
    const av = line!.at(-1)!;
    expect(av.text).toBe('AV');
    expect(av.x).toBeCloseTo(51 + 8 + 3, 6);
    expect(av.advance).toBeCloseTo(18, 6);
  });

  it('applies the pair at a font-slot seam between a glyph and its space', () => {
    // U+2019 resolves to the highAnsi slot and U+0020 to ascii: same face,
    // separate pieces. Its pair with the following space still applies.
    const text = 'x ’ x ’ x';
    const [kerned] = lines(layout([run(text, { kerning: 2 })], 400, 15));
    const [plain] = lines(layout([run(text)], 400, 15));
    expect(lastGlyphX(plain!)).toBeCloseTo(6 + 3 + 4 + 3 + 6 + 3 + 4 + 3, 6);
    expect(lastGlyphX(kerned!)).toBeCloseTo(lastGlyphX(plain!) - 2 * 0.5, 6);
  });

  it('hangs a slot-seam pair with its line-end space under right alignment', () => {
    const [line] = lines(layout([run('x \u2019 ', { kerning: 2 })], 100, 15, 'right'));
    const quote = line!.find(p => p.text === '\u2019')!;
    expect(quote.x + SCALAR['\u2019']!).toBeCloseTo(100, 6);
  });

  it('fits a line by its kerned spaces and takes no pair at a line start', () => {
    // Natural 'x A x A' is 6+3+9+3+6+3+9 = 39 wide; with its three pairs 36.75.
    const fitted = lines(layout([run('x A x A', { kerning: 2 })], 37, 15));
    expect(fitted.map(line => line.map(p => p.text).join(''))).toEqual(['x A x A']);
    const wrapped = lines(layout([run('x A x A', { kerning: 2 })], 30, 15));
    expect(wrapped.map(line => line.map(p => p.text).join(''))).toEqual(['x A x ', 'A']);
    expect(wrapped[1]![0]!.x).toBeCloseTo(0, 6);
  });

  it('takes no pair across a change of size', () => {
    const [line] = lines(layout([run('x ', { kerning: 2 }), run('A', { kerning: 2, fontSize: 13 })], 400, 15));
    expect(line!.at(-1)!.x).toBeCloseTo(9, 6);
  });
});
