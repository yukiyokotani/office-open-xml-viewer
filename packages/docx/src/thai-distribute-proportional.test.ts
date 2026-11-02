import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServicesRuntimeView } from './layout/runtime-state.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocSettings, DocxDocumentModel, DocxTextRun, SectionProps } from './types';

// word-thai-distribute-proportional-spread (#2045): synthetic Windows Word
// controls, Tahoma Thai with Latin labels, modes unset and 15.

/** Stub advances in em: Thai SARA AA and the space 0.25, Thai above/below
 * marks 0, everything else 0.5. */
function advanceEm(character: string): number {
  if (/[ัิ-ฺ็-๎]/u.test(character)) return 0;
  return character === 'า' || character === ' ' ? 0.25 : 0.5;
}
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
        width: [...s].reduce((sum, character) => sum + advanceEm(character) * p, 0),
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

const EMPTY_STORY = Object.freeze({ default: null, first: null, even: null });

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
const BR = { type: 'break', breakType: 'line' } as DocParagraph['runs'][number];

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

/** The x of every non-space shaped cluster on the line, left to right. */
function xs(line: LineLayout): number[] {
  return line.placements.filter((p): p is TextPlacement => p.kind === 'text')
    .flatMap((placement) => placement.clusters
      .filter((cluster) => placement.text.slice(
        cluster.range.start - placement.range.start, cluster.range.end - placement.range.start).trim() !== '')
      .map((cluster) => placement.origin.xPt + cluster.offset.xPt))
    .sort((a, b) => a - b);
}
const near = (values: readonly number[]) => values.map((x) => expect.closeTo(x, 6));
/** A break-ended pre-15 thaiDistribute line holding `runs`. */
const brLine = (...runs: DocParagraph['runs']) =>
  lines(para([...runs, BR, textRun('x')], 'thaiDistribute'))[0]!;

describe('word-thai-distribute-proportional-spread', () => {
  it('spreads a pre-15 thaiDistribute line by cluster advance, not per cluster', () => {
    // กากา: advances 10, 5, 10, 5 pt; 330 pt of slack over the first three
    // clusters' 25 pt (the last cluster keeps its advance): factor 13.2.
    expect(xs(brLine(textRun('กากา')))).toEqual(near([20, 162, 233, 375]));
  });

  it('keeps Latin letters at their advance and widens the space and the Thai clusters', () => {
    // a b ␠ ก า ก า: Latin 10 + 10, space 5, Thai 10 5 10 5 (55 pt). The space
    // and the first three Thai clusters (30 pt) share 305 pt.
    const k = 305 / 30;
    const thai = 20 + 25 + 5 * k;
    expect(xs(brLine(textRun('ab กากา'))))
      .toEqual(near([20, 30, thai, thai + 10 * (1 + k), thai + 15 * (1 + k), 375]));
  });

  it('widens the Thai cluster before a space as well as the space', () => {
    // ก า ␠ ก า: 10 5 5 10 5 pt; the four cells before the last share 325 pt.
    const k = 325 / 30;
    expect(xs(brLine(textRun('กา กา')))).toEqual(near([20, 30 + 10 * k, 40 + 20 * k, 375]));
  });

  it('spreads a soft-wrapped pre-15 thaiDistribute line the same way', () => {
    // Twelve กากา words (30 pt) with spaces: line 1 holds ten (345 pt). Every
    // cell but the last า (340 pt) widens by the factor 15 / 340; the
    // line-end space hangs.
    const [first] = lines(para([textRun(Array(12).fill('กากา').join(' '))], 'thaiDistribute'));
    const k = 15 / 340;
    const expected: number[] = [];
    let x = 20;
    for (let word = 0; word < 10; word++) {
      for (const advance of [10, 5, 10, 5]) { expected.push(x); x += advance * (1 + k); }
      x += 5 * (1 + k);
    }
    expect(xs(first!)).toEqual(near(expected));
  });

  it('weighs a cluster continued across a formatting seam by its whole advance', () => {
    // กี่ (10 pt, marks 0) า ก า: the run split after ก must not drop ก's weight.
    const whole = xs(brLine(textRun('กี่ากา')));
    const k = 330 / 25;
    expect(whole).toEqual(near([20, 30 + 10 * k, 35 + 15 * k, 375]));
    // Split, the zero-advance marks shape on their own at the end of ก (30 pt)
    // and every later cluster keeps its unsplit position.
    expect(xs(brLine(textRun('ก'), textRun('ี่ากา', { color: 'FF0000' }))))
      .toEqual(near([20, 30, ...whole.slice(1)]));
  });

  it('measures the cluster weights without reshaping the run', () => {
    // 100 spaced words in one run. Shaping work (request text plus its
    // shaping context, UTF-16 units) must match mode 15, whose equal spread
    // takes no weights: reshaping each cluster in its run context would cost
    // a whole run per cluster.
    const text = Array(100).fill('กากา').join(' ');
    const shapedUnits = (settings?: DocSettings): number => {
      const model: DocxDocumentModel = {
        body: [para([textRun(text)], 'thaiDistribute')], section,
        headers: EMPTY_STORY, footers: EMPTY_STORY, footnotes: [], endnotes: [], fontFamilyClasses: {},
        ...(settings ? { settings } : {}),
      };
      const services = createLayoutServices(model, { measureContext: makeCtx() });
      let units = 0;
      const view = createLayoutServicesRuntimeView(services, { text: {
        ...services.text,
        shape(request) {
          units += request.text.length + (request.substituteContext?.text.length ?? 0);
          return services.text.shape(request);
        },
      } });
      layoutDocument(model, view, { currentDateMs: 0 });
      return units;
    };
    expect(shapedUnits()).toBe(shapedUnits({ compatibilityMode: 15 }));
  });

  it('keeps the equal spread on a piece ended by a page or column break', () => {
    const body = para([textRun('กากา')], 'thaiDistribute') as BodyElement & { endsAtHardBreak?: boolean };
    body.endsAtHardBreak = true;
    expect(xs(lines(body)[0]!)).toEqual(near([20, 140, 255, 375]));
  });

  it('keeps the equal spread above mode 15 and under a character grid', () => {
    const body = para([textRun('กากา'), BR, textRun('x')], 'thaiDistribute');
    expect(xs(lines(body, { compatibilityMode: 16 })[0]!)).toEqual(near([20, 140, 255, 375]));
    const grid = { ...section, docGridType: 'linesAndChars', docGridCharSpace: 0 } as SectionProps;
    const node = layoutBodyModel([body], grid, makeCtx(), {}).pages[0]!.layers.body[0]!;
    if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(xs(node.lines[0]!)).not.toEqual(near([20, 162, 233, 375]));
  });

  it('keeps the equal per-cluster spread from mode 15', () => {
    const [line] = lines(para([textRun('กากา'), BR, textRun('x')], 'thaiDistribute'), { compatibilityMode: 15 });
    expect(xs(line!)).toEqual(near([20, 140, 255, 375]));
  });

  // Unmeasured line classes keep the equal spread of v0.89 (the values below
  // are its output).
  it.each([
    ['Thai with Myanmar', [textRun('กาကိကိ')], [20, 131.66666666666669, 238.33333333333334, 360]],
    ['Thai with supplementary Han', [textRun('กา𠀀 กา')], [20, 135, 140, 260, 375]],
    ['Thai with fixed-pitch CJK', [textRun('กา'), textRun('漢', { fitTextVal: 200, fitTextId: 1 }), textRun(' กา')], [20, 135, 140, 260, 375]],
    ['Thai directly before Latin', [textRun('กากาab')], [20, 133.33333333333331, 241.66666666666666, 355, 360, 370]],
    ['a Thai mark after a space', [textRun('กา \u0e34กา')], [20, 111.25, 202.5, 283.75, 375]],
  ])('keeps the equal spread on a line with %s', (_name, runs, base) => {
    expect(xs(brLine(...runs))).toEqual(near(base));
  });
});
