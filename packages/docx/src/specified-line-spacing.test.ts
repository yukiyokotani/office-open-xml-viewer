import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { DocParagraph, DocxDocumentModel, DocRun, LineSpacing } from './types.js';
import type { ParagraphLayout } from './layout/types.js';

// Word printing PDFs (Mac modes 14/15; Windows mode 15 adjudication for FE
// usWin boxes): coordinates are quantized to
// 0.24pt. Expectations below are the measured nominal geometry, not Canvas
// font boxes (the deliberately different stub cannot supply the Word rule).
function render(font: string, size: number, spacing: LineSpacing, grid = false, hostSize?: number, mode = 15, underline = false, markSize = size, wrap = false, localAlias: boolean | 'provided' = false, decoratedLast = false, markCsFace?: string, numbered?: { markerSize: number; empty?: boolean }) {
  let canvasFont = '';
  const context = {
    get font() { return canvasFont; }, set font(v: string) { canvasFont = v; },
    measureText(text: string) {
      const px = Number(/([\d.]+)px/.exec(canvasFont)?.[1] ?? 10);
      return { width: text.length * px / 3, fontBoundingBoxAscent: px * .7,
        fontBoundingBoxDescent: px * .3, actualBoundingBoxAscent: px * .7,
        actualBoundingBoxDescent: px * .3 } as TextMetrics;
    },
  } as CanvasRenderingContext2D;
  const text = (value: string): DocRun => ({ type: 'text', text: value,
    fontFamily: font, fontSize: size, bold: false, italic: false,
    underline, strikethrough: false, color: null, isLink: false,
    background: null, vertAlign: null } as DocRun);
  const runs: DocRun[] = [];
  for (let i = 0; i < 3; i++) {
    if (hostSize !== undefined) runs.push({ type: 'anchorHost', fontFamily: font,
      fontFamilyEastAsia: font, fontSize: hostSize, bold: false, italic: false } as DocRun);
    runs.push({ ...text('PROBE Hgyp'), ...(decoratedLast && i === 2 ? { underline: true } : {}) } as DocRun);
    if (i < 2) runs.push({ type: 'break', breakType: 'line', fontSize: size } as DocRun);
  }
  if (wrap) runs.push({ type: 'image', imagePath: 'word/media/test.png',
    mimeType: 'image/png', widthPt: 80, heightPt: 60, anchor: true,
    anchorXPt: 80, anchorYPt: 0, anchorXRelativeFrom: 'margin',
    anchorYRelativeFrom: 'margin', anchorYFromPara: false,
    wrapMode: 'square', wrapSide: 'bothSides' } as DocRun);
  const paragraph = { type: 'paragraph', runs: numbered?.empty ? [] : runs, alignment: 'left',
    indentLeft: numbered ? 36 : 0,
    indentRight: 0, indentFirst: numbered ? -36 : 0, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: spacing, tabStops: [],
    numbering: numbered ? {
      numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 36, tab: 36, suff: 'tab',
      fontFamily: font, fontFacts: { fontSize: numbered.markerSize, fontFamily: font },
    } : null,
    defaultFontSize: markSize,
    defaultFontFamily: font, widowControl: false,
    ...(markCsFace ? { paragraphMarkFontFacts: { fontFamily: font, fontFamilyCs: markCsFace,
      fontSize: markSize } } : {}),
  } as DocParagraph;
  const empty = { default: null, first: null, even: null };
  const model = { body: [paragraph], headers: empty, footers: empty, settings: { compatibilityMode: mode },
    section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginLeft: 72,
      marginRight: 72, marginBottom: 72, headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
      ...(grid ? { docGridType: 'lines', docGridLinePitch: 20 } : {}) },
  } as DocxDocumentModel;
  const layout = layoutDocument(model, createLayoutServices(model, {
    measureContext: context,
    ...(localAlias ? { localMetrics: { [font.toLowerCase()]: {
      family: 'Installed test face', requestedFamily: font,
      weight: 400, style: 'normal' as const, sourceIdentity: localAlias === 'provided'
        ? 'provided-sfnt:test-face' : 'office-local:test-installed-face',
      lineHeightRatio: 1.14990234375, designAscentRatio: .9052734375,
      designDescentRatio: .2119140625,
    } } } : {}),
  }));
  return layout.pages[0].layers.body.find(n => n.kind === 'paragraph') as ParagraphLayout;
}

describe('Word specified spacing for numbered lines (#1884)', () => {
  // Windows Word 16.0 #1884 controls: numbered text lines and empty items keep
  // the plain partition; a taller marker raises an atLeast line by its ascent.
  it.each([false, true])('keeps four fifths of an exact line, empty=%s', (empty) => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 18, explicit: true },
      false, undefined, 15, false, 10, false, false, false, undefined, { markerSize: 10, empty });
    expect(paragraph.lines[0].baselinePt - paragraph.lines[0].bounds.yPt).toBeCloseTo(14.4, 10);
  });

  it.each([
    // The stub's 30pt marker ascent (21pt) over the text's design descent.
    { empty: false, baseline: 21 + 24 - 21 - 434 / 2048 * 10 },
    // An empty item stands on its mark's own (stub) 3pt descent.
    { empty: true, baseline: 21 },
  ])('raises an atLeast line by a taller marker ascent, empty=$empty', ({ empty, baseline }) => {
    const tall = render('Arial', 10, { rule: 'atLeast', value: 18, explicit: true },
      false, undefined, 15, false, 10, false, false, false, undefined, { markerSize: 30, empty });
    const line = tall.lines[0];
    expect(line.baselinePt - line.bounds.yPt).toBeCloseTo(baseline, 6);
  });
});

describe('numbered specified lines outside compatibility mode 15 (#1884)', () => {
  it.each([
    { empty: false, rule: 'exact' as const }, { empty: true, rule: 'exact' as const },
    { empty: false, rule: 'atLeast' as const }, { empty: true, rule: 'atLeast' as const },
  ])('keeps the established $rule line, empty=$empty', ({ empty, rule }) => {
    // The values main produced before the mode-15 numbered class was measured;
    // an exact line keeps four fifths in every mode (WORD_EXACT_LINE_BASELINE).
    const paragraph = render('Arial', 10, { rule, value: 18, explicit: true },
      false, undefined, 14, false, 10, false, false, false, undefined, { markerSize: 30, empty });
    expect(paragraph.lines[0].baselinePt - paragraph.lines[0].bounds.yPt).toBeCloseTo(rule === 'exact' ? 14.4 : 11, 9);
    expect(paragraph.lines[0].advancePt).toBe(18);
  });
});

describe('atLeast with a raised marker outside the specified class (#1884)', () => {
  it('takes the larger of the minimum and the raised box', () => {
    // Underlined text leaves WORD_SPECIFIED_TEXT_LINE_BOX; the stub's 30pt
    // marker ascent (21pt) over the 3pt text descent still raises the box.
    const paragraph = render('Arial', 10, { rule: 'atLeast', value: 18, explicit: true },
      false, undefined, 15, true, 10, false, false, false, undefined, { markerSize: 30 });
    expect(paragraph.lines[0].advancePt).toBeCloseTo(24, 6);
  });
});

describe('Word specified spacing for ordinary Latin text', () => {
  it.each([
    ['Arial', 10, 30, false], ['Verdana', 10, 24, true],
    ['Yu Mincho', 30, 12, true],
  ] as const)('places %s exact text independently of its font box and anchor hosts', (font, size, height, grid) => {
    const free = render(font, size, { rule: 'exact', value: height, explicit: true }, grid);
    const hosted = render(font, size, { rule: 'exact', value: height, explicit: true }, grid, 30);
    for (const line of free.lines) expect(line.baselinePt - line.bounds.yPt).toBeCloseTo(height * 4 / 5, 10);
    expect(hosted.lines.map(l => l.baselinePt)).toEqual(free.lines.map(l => l.baselinePt));
    expect(hosted.flowBounds.heightPt).toBe(3 * height);
  });

  it('keeps four fifths of an exact line in compatibility modes 12 and 14 (WORD_EXACT_LINE_BASELINE)', () => {
    const old = render('Arial', 10, { rule: 'exact', value: 30, explicit: true }, false, undefined, 12);
    expect(old.lines[0].baselinePt - 72).toBe(24);
    const fourteen = render('Arial', 10, { rule: 'exact', value: 30, explicit: true }, false, undefined, 14);
    expect(fourteen.lines[0].baselinePt - 72).toBe(24);
  });

  it('keeps four fifths of exact heights outside 12-30pt (WORD_EXACT_LINE_BASELINE)', () => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 40, explicit: true });
    expect(paragraph.lines[0].baselinePt - 72).toBe(32);
    // Windows Word line1 L1a: text far taller than the line and a tall line
    // over small text keep the same partition.
    for (const [size, height] of [[24, 6], [9, 96], [24, 24]] as const) {
      const lines = render('Arial', size, { rule: 'exact', value: height, explicit: true }).lines;
      expect(lines.map(line => line.baselinePt - line.bounds.yPt))
        .toEqual(lines.map(() => expect.closeTo(height * 4 / 5, 9)));
      expect(lines[0]!.advancePt).toBe(height);
    }
    expect(paragraph.flowBounds.heightPt).toBe(120);
  });

  it('keeps four fifths of an exact line beside a wrapping float (WORD_EXACT_LINE_BASELINE)', () => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 30, explicit: true },
      false, undefined, 15, false, 10, true);
    expect(paragraph.lines[0].baselinePt - paragraph.lines[0].bounds.yPt).toBe(24);
  });

  it('keeps installed local aliases in the authored-face baseline class', () => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 30, explicit: true },
      false, 30, 15, false, 10, false, true);
    expect(paragraph.lines[0].baselinePt - 72).toBe(24);
  });

  it('uses admitted host sides for atLeast baselines instead of the authored catalog', () => {
    const paragraph = render('Times New Roman', 10,
      { rule: 'atLeast', value: 1, explicit: true },
      false, undefined, 15, false, 10, false, true);
    // The installed fixture has its own 2355-unit line and 434-unit descent;
    // the authored catalog's 2268-unit line / 443-unit descent cannot replace it.
    expect(paragraph.lines[0].advancePt).toBeCloseTo(10 * 2355 / 2048, 10);
    expect(paragraph.lines[0].baselinePt - 72).toBeCloseTo(10 * (2355 - 434) / 2048, 10);
  });

  it('preserves selected application-provided resources instead of borrowing authored reference geometry', () => {
    const spacing = { rule: 'exact', value: 30, explicit: true } as const;
    const paragraph = render('Arial', 10, spacing, false, undefined, 15, false, 10, false, 'provided');
    const previous = render('Arial', 10, spacing, false, undefined, 12, false, 10, false, 'provided');
    expect(paragraph.lines.map(line => line.baselinePt)).toEqual(previous.lines.map(line => line.baselinePt));
  });

  it('puts non-grid atLeast leading before text without reserving a tall floating host', () => {
    const spacing = { rule: 'atLeast', value: 24, explicit: true } as const;
    const free = render('Arial', 10, spacing);
    const hosted = render('Arial', 10, spacing, false, 30);
    expect(free.lines[0].baselinePt - 72).toBeCloseTo(21.880859375, 10);
    expect(hosted.lines.map(l => l.baselinePt)).toEqual(free.lines.map(l => l.baselinePt));
    expect(hosted.flowBounds.heightPt).toBe(72);
  });

  it.each([
    [10, 24, 19.20],
    [30, 50.16, 35.64],
  ])('uses the Windows usWin descent for %spt Far East atLeast text', (size, advance, baseline) => {
    const paragraph = render('Yu Mincho', size, { rule: 'atLeast', value: 24, explicit: true });
    expect(Math.abs(paragraph.lines[0].advancePt - advance)).toBeLessThan(.24);
    expect(Math.abs(paragraph.lines[0].baselinePt - 72 - baseline)).toBeLessThan(.24);
  });

  it('keeps four fifths for decorated text and a differently-sized paragraph mark (WORD_EXACT_LINE_BASELINE)', () => {
    const spacing = { rule: 'exact', value: 30, explicit: true } as const;
    const underlined = render('Arial', 10, spacing, false, undefined, 15, true);
    const differentMark = render('Arial', 10, spacing, false, undefined, 15, false, 12);
    expect(underlined.lines[0].baselinePt - 72).toBe(24);
    expect(differentMark.lines[0].baselinePt - 72).toBe(24);
  });

  it('keeps every line of a paragraph decorated on a later line at four fifths', () => {
    const paragraph = render('Arial', 10, { rule: 'exact', value: 30, explicit: true },
      false, undefined, 15, false, 10, false, false, true);
    expect(paragraph.lines.map(line => line.baselinePt - line.bounds.yPt)).toEqual([24, 24, 24]);
  });

  it('keeps four fifths beside a same-size paragraph mark of another face', () => {
    // Windows Word 16.0 #1884: a mark of another face at the text size keeps
    // the exact partition of the homogeneous controls.
    const paragraph = render('Arial', 10, { rule: 'exact', value: 30, explicit: true },
      false, undefined, 15, false, 10, false, false, false, 'Times New Roman');
    expect(paragraph.lines[0].baselinePt - 72).toBe(24);
  });

  it('allocates whole grid cells for the explicit single-line diagnostic', () => {
    const paragraph = render('Yu Mincho', 30, { rule: 'auto', value: 1, explicit: true }, true);
    expect(paragraph.lines[0].advancePt).toBe(60);
    expect(Math.abs(paragraph.lines[0].baselinePt - 72 - 40.56)).toBeLessThan(.24);
  });

  it('keeps grid-atLeast floating hosts out of visible-text allocation', () => {
    const hosted = render('Arial', 10, { rule: 'atLeast', value: 24, explicit: true }, true, 30);
    const free = render('Arial', 10, { rule: 'atLeast', value: 24, explicit: true }, true);
    expect(hosted.lines.map(line => line.baselinePt)).toEqual(free.lines.map(line => line.baselinePt));
    expect(hosted.flowBounds.heightPt).toBe(72);
  });

  it.each([
    ['Arial', 10, 24, 24, 17.52],
    ['Arial', 10, 40, 40, 33.60],
    ['Arial', 30, 24, 40, 30.96],
    ['Yu Mincho', 30, 24, 60, 40.56],
  ] as const)('places %s grid-atLeast leading before the normal cell box', (font, size, minimum, advance, baseline) => {
    const paragraph = render(font, size, { rule: 'atLeast', value: minimum, explicit: true }, true);
    expect(paragraph.lines.map(line => line.advancePt)).toEqual([advance, advance, advance]);
    expect(Math.abs(paragraph.lines[0].baselinePt - 72 - baseline)).toBeLessThan(.24);
  });
});
