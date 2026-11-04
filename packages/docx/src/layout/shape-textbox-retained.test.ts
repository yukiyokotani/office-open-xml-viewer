import { describe, expect, it } from 'vitest';
import type { ParagraphLayoutContext } from '../layout-context.js';
import type { ShapeRun, ShapeTextRun } from '../types.js';
import type { DocxDocumentModel } from '../types.js';
import { createLayoutServices } from '../layout-runtime.js';
import { paintPlacedTextBoxLayout } from '../paint/canvas-text.js';
import type { CanvasPaintResourcePainter } from '../paint/types.js';
import { acquireShapeTextBoxLayout } from './paragraph.js';

const context: ParagraphLayoutContext = {
  lineGrid: { active: false, pitchPt: null },
  characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
  rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
  physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
  lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0,
  baseRtl: false, isJustified: false, stretchLastLine: false,
  tabStops: [], hasRuby: false, hasEastAsianText: false,
  kinsoku: { enabled: true, lineStartForbidden: new Set(), lineEndForbidden: new Set() },
  defaultTabPt: 36,
};

function canvas(): CanvasRenderingContext2D {
  let font = '10px sans-serif';
  const ctx = {
    get font() { return font; }, set font(value: string) { font = value; },
    fillStyle: '', strokeStyle: '', lineWidth: 1, textAlign: 'left',
    textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px', fontKerning: 'auto',
    measureText(text: string) {
      const size = Number.parseFloat(/([\d.]+)px/u.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * .8, fontBoundingBoxDescent: size * .2,
        actualBoundingBoxAscent: size * .8, actualBoundingBoxDescent: size * .2,
      } as TextMetrics;
    },
    save() {}, restore() {}, translate() {}, rotate() {}, scale() {},
    setLineDash() {}, beginPath() {}, moveTo() {}, lineTo() {}, stroke() {},
    fillRect() {}, strokeRect() {}, fillText() {}, drawImage() {}, rect() {}, clip() {},
  } as unknown as CanvasRenderingContext2D;
  return ctx;
}

describe('retained shape text-box acquisition', () => {
  it.each([
    { label: 'omitted CS companions retain public Latin formatting', facts: { rtl: true }, weight: 700, style: 'italic', size: 12, family: 'Latin Face' },
    { label: 'explicit CS values override Latin formatting', facts: { rtl: true, boldCs: false, italicCs: false, fontSizeCsPt: 18, fontFamilyCs: 'CS Face' }, weight: 400, style: 'normal', size: 18, family: 'CS Face' },
    { label: 'cs selects its companions without rtl', facts: { cs: true, boldCs: false, fontSizeCsPt: 16, fontFamilyCs: 'CS Face' }, weight: 400, style: 'italic', size: 16, family: 'CS Face' },
  ])('$label', ({ facts, weight, style, size, family }) => {
    const measureContext = canvas();
    const services = createLayoutServices({
      section: {
        pageWidth: 612, pageHeight: 792,
        marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
        headerDistance: 36, footerDistance: 36,
        titlePage: false, evenAndOddHeaders: false,
      },
      body: [], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as DocxDocumentModel, { measureContext });
    const run: ShapeTextRun = {
      text: 'אבג', fontSizePt: 12, fontFamily: 'Latin Face', bold: true, italic: true, ...facts,
    };
    const shape = {
      textBlocks: [{ text: run.text, fontSizePt: 12, alignment: 'left', runs: [run] }],
    } as unknown as ShapeRun;
    const layout = acquireShapeTextBoxLayout(shape, { xPt: 0, yPt: 0, widthPt: 300, heightPt: 100 }, {
      id: 'cs-shape', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      flowDomainId: 'body', context,
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, documentHasEastAsianText: false,
        pageWritingMode: 'horizontal-tb', layoutServices: services,
      },
    });
    const paragraph = layout?.story.blocks[0];
    if (paragraph?.kind !== 'paragraph') throw new Error('expected retained paragraph');
    const placements = paragraph.lines.flatMap((line) => line.placements);
    expect(placements.length).toBeGreaterThan(0);
    for (const placement of placements) {
      if (placement.kind !== 'text') throw new Error('expected retained text');
      expect(placement.fontWeight).toBe(weight);
      expect(placement.fontStyle).toBe(style);
      expect(placement.fontSizePt).toBe(size);
      expect(placement.fontRoute.familyList.split(',')[0]).toBe(JSON.stringify(family));
    }
  });

  it.each([
    { anchor: undefined, expectedTopPt: 30, label: 'defaults to top' },
    { anchor: 'ctr', expectedTopPt: 65, label: 'centers' },
    { anchor: 'b', expectedTopPt: 100, label: 'bottom-aligns' },
  ])('$label the retained text story from bodyPr@anchor', ({ anchor, expectedTopPt }) => {
    const measureContext = canvas();
    const services = createLayoutServices({
      section: {
        pageWidth: 612, pageHeight: 792,
        marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
        headerDistance: 36, footerDistance: 36,
        titlePage: false, evenAndOddHeaders: false,
      },
      body: [], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as DocxDocumentModel, { measureContext });
    const shape = {
      textInsetL: 0, textInsetT: 10, textInsetR: 0, textInsetB: 10,
      ...(anchor ? { textAnchor: anchor } : {}),
      textBlocks: [{
        text: 'abcd', fontSizePt: 10, color: '112233', alignment: 'left',
        runs: [{ text: 'abcd', fontSizePt: 10, color: '112233' }],
      }],
    } as unknown as ShapeRun;

    const layout = acquireShapeTextBoxLayout(
      shape,
      { xPt: 10, yPt: 20, widthPt: 40, heightPt: 100 },
      {
        id: `anchor-${anchor ?? 'default'}`,
        source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        flowDomainId: 'body', context,
        measurer: { context: measureContext, fontFamilyClasses: {} },
        environment: {
          pageIndex: 0, totalPages: 1, documentHasEastAsianText: false,
          pageWritingMode: 'horizontal-tb',
          layoutServices: services,
        },
      },
    );

    expect(layout?.story.blocks[0]?.flowBounds.yPt).toBe(expectedTopPt);
    expect(layout?.contentBounds).toEqual({ xPt: 10, yPt: 20, widthPt: 40, heightPt: 100 });
  });

  it('partitions once and paint at scale 1/2 never measures', () => {
    const measureContext = canvas();
    const services = createLayoutServices({
      section: {
        pageWidth: 612, pageHeight: 792,
        marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
        headerDistance: 36, footerDistance: 36,
        titlePage: false, evenAndOddHeaders: false,
      },
      body: [], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as DocxDocumentModel, { measureContext });
    const shape = {
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
      textAnchor: 't', textBlocks: [{
        text: 'abcdefghij', fontSizePt: 10, color: '112233', alignment: 'left',
        runs: [{ text: 'abcdefghij', fontSizePt: 10, color: '112233' }],
      }],
    } as unknown as ShapeRun;
    const layout = acquireShapeTextBoxLayout(shape, { xPt: 10, yPt: 20, widthPt: 40, heightPt: 100 }, {
      id: 'shape', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      flowDomainId: 'body', context,
      measurer: { context: measureContext, fontFamilyClasses: {} },
      environment: {
        pageIndex: 0, totalPages: 1, documentHasEastAsianText: false,
        pageWritingMode: 'horizontal-tb',
        layoutServices: services,
      },
    });

    expect(layout?.story.blocks[0]?.kind).toBe('paragraph');
    const paragraph = layout?.story.blocks[0];
    expect(paragraph?.kind === 'paragraph'
      ? paragraph.lines.map((line) => line.range)
      : []).toEqual([
      { start: 0, end: 4 }, { start: 4, end: 8 }, { start: 8, end: 10 },
    ]);
    const resources: CanvasPaintResourcePainter = { paint() { throw new Error('unexpected resource'); } };
    for (const scale of [1, 2]) {
      const paintContext = canvas();
      paintContext.measureText = () => { throw new Error('paint must not measure'); };
      expect(() => paintPlacedTextBoxLayout(layout!, {
        ctx: paintContext, scale, dpr: 1, resources,
      })).not.toThrow();
    }
  });

  it('resolves spAutoFit from the retained line stack during acquisition', () => {
    const measureContext = canvas();
    const services = createLayoutServices({
      section: {
        pageWidth: 612, pageHeight: 792,
        marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
        headerDistance: 36, footerDistance: 36,
        titlePage: false, evenAndOddHeaders: false,
      },
      body: [], headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as DocxDocumentModel, { measureContext });
    const shape = {
      textAutofit: 'sp',
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
      textAnchor: 't', textBlocks: [{
        text: 'abcdefghij', fontSizePt: 10, color: '112233', alignment: 'left',
        runs: [{ text: 'abcdefghij', fontSizePt: 10, color: '112233' }],
      }],
    } as unknown as ShapeRun;
    const layout = acquireShapeTextBoxLayout(
      shape,
      { xPt: 10, yPt: 20, widthPt: 40, heightPt: 100 },
      {
        id: 'autofit-shape', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        flowDomainId: 'body', context,
        measurer: { context: measureContext, fontFamilyClasses: {} },
        environment: {
          pageIndex: 0, totalPages: 1, documentHasEastAsianText: false,
          pageWritingMode: 'horizontal-tb',
          layoutServices: services,
        },
      },
    );

    expect(layout?.flowBounds).toEqual({ xPt: 10, yPt: 20, widthPt: 40, heightPt: 30 });
    expect(layout?.contentBounds).toEqual(layout?.flowBounds);
  });
});
