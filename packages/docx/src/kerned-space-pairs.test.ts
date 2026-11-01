import { describe, expect, it } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import type { DocumentLayout } from './layout/types.js';
import { measureParagraphIntrinsicWidths } from './layout/intrinsic-width.js';
import type { ParagraphLayoutContext } from './layout-context.js';
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
interface LayoutOptions {
  readonly tabStops?: DocParagraph['tabStops'];
  readonly settings?: Record<string, unknown>;
  readonly section?: Partial<SectionProps>;
}
function paragraphOf(runs: DocxTextRun[], alignment: DocParagraph['alignment'] = 'left',
  tabStops: DocParagraph['tabStops'] = []): DocParagraph {
  return { runs, alignment, indentLeft: 0, indentRight: 0,
    indentFirst: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops, defaultFontSize: 12, defaultFontFamily: 'Arial', widowControl: false } as unknown as DocParagraph;
}
function modelOf(paragraph: DocParagraph, width: number, compatibilityMode?: number,
  options: LayoutOptions = {}): DocxDocumentModel {
  const body: BodyElement[] = [{ type: 'paragraph', ...paragraph }];
  return { body, settings: { ...(compatibilityMode === undefined ? {} : { compatibilityMode }), ...options.settings },
    section: { pageWidth: width, pageHeight: 1000, marginTop: 0, marginRight: 0, marginBottom: 0,
      marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false,
      evenAndOddHeaders: false, ...options.section } as SectionProps,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {}, footnotes: [] } as unknown as DocxDocumentModel;
}
function layout(runs: DocxTextRun[], width: number, compatibilityMode?: number,
  alignment: DocParagraph['alignment'] = 'left', options: LayoutOptions = {}): DocumentLayout {
  const doc = modelOf(paragraphOf(runs, alignment, options.tabStops), width, compatibilityMode, options);
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

  it('folds the pair into the zero advance floor of over-condensed text', () => {
    // WORD_CONDENSED_SPACING_ADVANCE_FLOOR: every character floors at zero.
    const glyphs = (line: { text: string; x: number }[]) => line.filter(p => p.text.trim());
    const [floored] = lines(layout([run('x A x', { kerning: 2, charSpacing: -20 })], 400, 15));
    expect(glyphs(floored!).map(p => p.x)).toEqual([0, 0, 0]);
    expect(floored!.every(p => p.advance >= 0)).toBe(true);
    // Spaces of 0.5 floor with their pair: the space+A pair can only remove
    // that 0.5 (A then x at 3.5 and 10), not its full 0.75.
    const [partial] = lines(layout([run('x A x', { kerning: 2, charSpacing: -2.5 })], 400, 15));
    expect(glyphs(partial!).map(p => p.x)).toEqual([0, 3.5, 10]);
  });

  it('aligns a right-tab cell by its kerned widths', () => {
    const [line] = lines(layout([run('\tx A x', { kerning: 2 })], 400, 15, 'left',
      { tabStops: [{ pos: 100, alignment: 'right', leader: 'none' }] }));
    expect(line!.at(-1)!.x + SCALAR.x!).toBeCloseTo(100, 6);
  });

  it('admits a hyphen prefix by its kerned width', () => {
    // 'x A-' is 25 natural and 24.25 with its space+A pair.
    const wrapped = lines(layout([run('x A-BC', { kerning: 2 })], 24.5, 15));
    expect(wrapped.map(line => line.map(p => p.text).join(''))).toEqual(['x A-', 'BC']);
  });

  it('admits an optional hyphen prefix by its kerned width', () => {
    const shy = { ...run('', { kerning: 2 }), __optionalHyphen: true } as DocxTextRun;
    const wrapped = lines(layout([run('x A', { kerning: 2 }), shy, run('BC', { kerning: 2 })], 24.5, 15));
    expect(wrapped.map(line => line.map(p => p.text).join(''))).toEqual(['x A-', 'BC']);
  });

  it('keeps the seam pair on a head kept by kinsoku retraction', () => {
    const result = lines(layout([run('x ABC\u6f22', { kerning: 2 })], 34, 15, 'left',
      { settings: { kinsoku: true, noLineBreaksBefore: '\u6f22' } }));
    expect(result.map(line => line.map(p => p.text).join(''))).toEqual(['x AB', 'C\u6f22']);
    const head = result[0]!.at(-1)!;
    expect(head.x).toBeCloseTo(9 - 0.75, 6);
    expect(head.advance).toBeCloseTo(17 - 0.75, 6);
  });

  it('keeps tate-chu-yoko cells one em along the column', () => {
    const [line] = lines(layout([run('x A', { kerning: 2, eastAsianVert: true }),
      run('\u6f22', { kerning: 2 })], 400, 15, 'left', { section: { textDirection: 'tbRl' } as Partial<SectionProps> }));
    const cells = line!.filter(p => p.text !== '\u6f22');
    expect(cells.map(p => p.advance)).toEqual([12, 12]);
  });

  it('sizes max-content with the same slot-seam pair as bounded lines', () => {
    const paragraph = paragraphOf([run('x \u2019 x', { kerning: 2 })]);
    const doc = modelOf(paragraph, 400, 15);
    const measure = context();
    const intrinsic = measureParagraphIntrinsicWidths(paragraph, INTRINSIC_CONTEXT, 400,
      { context: measure, fontFamilyClasses: {} }, {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false, compatibilityMode: 15,
        layoutServices: createLayoutServices(doc, { measureContext: measure }),
      });
    expect(intrinsic.maxWidthPt).toBeCloseTo(6 + 3 + 4 + 3 + 6 - 0.5, 6);
  });

  it('hangs only the seam pair, never snap-grid padding, of a line-end space', () => {
    const grid = { section: { docGridType: 'snapToChars', docGridLinePitch: 18,
      docGridCharSpace: 4096 * 3 } as Partial<SectionProps> };
    // Unkerned, so no seam pair: the ideograph keeps the position it has
    // without WORD_KERNED_SPACE_PAIRS (its snapped space cell's leading
    // padding is not hanging advance).
    const [line] = lines(layout([run('\u6f22 ')], 100, 15, 'right', grid));
    expect(line![0]!.x).toBeCloseTo(80, 6);
  });

  it('folds the pair into the committed contextual advance of the seam space', () => {
    // 'A ' commits A 6.5 and a space of 3 - 0.75 - 2.5, floored to 0: the
    // following space+A pair has no advance left to remove.
    const [line] = lines(layout([run('A A', { kerning: 2, charSpacing: -2.5 })], 400, 15));
    expect(line!.filter(p => p.text.trim()).map(p => p.x)).toEqual([0, 6.5]);
  });

  it('admits an early explicit split before a glued slot suffix by its kerned width', () => {
    const wrapped = lines(layout([run('x A-BC\u2019', { kerning: 2 })], 24.5, 15));
    expect(wrapped.map(line => line.map(p => p.text).join(''))).toEqual(['x A-', 'BC\u2019']);
  });

  it('gives max-content the bounded ownership of a floored seam pair', () => {
    const runs = [run('x \u2019 A', { kerning: 2, charSpacing: -2.5 })];
    const paragraph = paragraphOf(runs);
    const doc = modelOf(paragraph, 400, 15);
    const measure = context();
    const { maxWidthPt } = measureParagraphIntrinsicWidths(paragraph, INTRINSIC_CONTEXT, 400,
      { context: measure, fontFamilyClasses: {} }, {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: false, compatibilityMode: 15,
        layoutServices: createLayoutServices(doc, { measureContext: measure }),
      });
    const [bounded] = lines(layout(runs, 400, 15));
    expect(maxWidthPt).toBeCloseTo(bounded!.at(-1)!.x + 9 - 2.5, 6);
    expect(lines(layout(runs, maxWidthPt, 15))).toHaveLength(1);
  });

  it('takes no pair in any piece of a ruby run', () => {
    const [line] = lines(layout([run('x A x A', { kerning: 2, ruby: { text: 'r', fontSizePt: 6 } } as Partial<DocxTextRun>)], 400, 15));
    // Only the A+space pair measured inside the 'A ' piece applies.
    expect(lastGlyphX(line!)).toBeCloseTo(30 - 0.75, 6);
  });

  it('hangs trailing punctuation after a visible prefix measured with its seam pair', () => {
    // 'x AB' is 26 natural and 25.25 with its space+A pair; the comma hangs.
    const text = (width: number) => lines(layout([run('x AB,', { kerning: 2 })], width, 15))
      .map(line => line.map(p => p.text).join(''));
    expect(text(25.25)).toEqual(['x AB,']);
    expect(text(25.2)).toEqual(['x ', 'AB,']);
  });

  it.each([
    ['ruby', { ruby: { text: 'r', fontSizePt: 6 } }],
    ['right-to-left', { rtl: true }],
  ] as const)('keeps max-content unkerned at %s seams that bounded lines leave unkerned', (_, extra) => {
    const runs = [run('x A x A', { kerning: 2, ...extra } as Partial<DocxTextRun>)];
    const maxWidthPt = maxContent(runs);
    // Natural 39; only the A+space pairs inside the 'A ' pieces apply.
    expect(maxWidthPt).toBeCloseTo(39 - 0.75, 6);
    expect(lines(layout(runs, maxWidthPt, 15))).toHaveLength(1);
  });
});

function maxContent(runs: DocxTextRun[]): number {
  const paragraph = paragraphOf(runs);
  const doc = modelOf(paragraph, 400, 15);
  const measure = context();
  return measureParagraphIntrinsicWidths(paragraph, INTRINSIC_CONTEXT, 400,
    { context: measure, fontFamilyClasses: {} }, {
      pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
      documentHasEastAsianText: false, compatibilityMode: 15,
      layoutServices: createLayoutServices(doc, { measureContext: measure }),
    }).maxWidthPt;
}

const INTRINSIC_CONTEXT: ParagraphLayoutContext = {
  lineGrid: { active: false, pitchPt: null },
  characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
  rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
  physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
  lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: false,
  isJustified: false, stretchLastLine: false, tabStops: [], hasRuby: false, hasEastAsianText: false,
  kinsoku: { enabled: true, lineStartForbidden: new Set(), lineEndForbidden: new Set() }, defaultTabPt: 36,
};
