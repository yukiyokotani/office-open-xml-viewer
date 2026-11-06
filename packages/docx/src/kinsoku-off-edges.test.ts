import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { layoutLines, type LayoutTextSeg } from './line-layout.js';
import { resolveDocumentLayoutSettings, resolveParagraphLayoutContext, resolveSectionLayoutContext } from './layout-context.js';
import { paragraphAcquisitionInput } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { measureParagraphIntrinsicWidths } from './layout/intrinsic-width.js';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { BodyElement, DocParagraph, DocSettings, DocxDocumentModel, DocxTextRun, SectionProps } from './types.js';

const section: SectionProps = {
  pageWidth: 20, pageHeight: 400,
  marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
};
function lines(text: string, alignment = 'thaiDistribute', overflowPunct = true, settings: DocSettings = {}, runFacts: { langEastAsia?: string } = {}): string[] {
  const run: DocxTextRun = {
    text, fontSize: 10, fontFamily: 'UnregisteredTestFace', bold: false, italic: false,
    underline: false, strikethrough: false, color: null, isLink: false,
    background: null, vertAlign: null, hyperlink: null, ...runFacts,
  };
  const body: BodyElement = {
    type: 'paragraph', alignment, kinsoku: false, overflowPunct,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...run }],
  };
  const ctx = {
    font: '10px serif', letterSpacing: '0px',
    measureText: (value: string) => ({ width: [...value].length * 10,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
  } as unknown as CanvasRenderingContext2D;
  return layoutBodyModel([body], section, ctx, {}, {
    characterSpacingControl: 'compressPunctuation', ...settings,
  }).pages.flatMap(page => page.layers.body.flatMap(node => node.kind === 'paragraph'
    ? node.lines.map(line => line.placements.filter(p => p.kind === 'text').map(p => p.text).join('')) : []));
}

describe('paragraph-off semantics and observed Japanese Thai line edges', () => {
  it('allows the measured closing head without globally re-enabling paragraph rules', () => {
    expect(lines('あい。う')).toEqual(['あい', '。う']);
  });
  it.each([
    ['あい々う', true, ['あ', 'い々', 'う']],
    ['あ（いう', true, ['あ', '（い', 'う']],
    ['あい。う', false, ['あ', 'い。', 'う']],
  ] as const)('retains the measured forbidden edge in %s', (text, overflow, expected) => {
    expect(lines(text, 'thaiDistribute', overflow)).toEqual(expected);
  });
  it('does not grant small-kana heads under the kana compression option', () => {
    expect(lines('あいぁう', 'thaiDistribute', true, { characterSpacingControl: 'compressPunctuationAndJapaneseKana' }))
      .toEqual(['あ', 'いぁ', 'う']);
  });
  it.each([
    { compatibilityMode: 15 },
    { noLineBreaksAfter: '' },
  ])('keeps normative OFF for the unmeasured profile or a present empty custom table', settings => {
    expect(lines('あい々う', 'thaiDistribute', false, settings)).toEqual(['あい', '々う']);
  });
  it.each(['zh-CN', 'ko-KR'])('does not transfer Japanese permissions to %s', language => {
    expect(lines('あい々う', 'thaiDistribute', false, {}, { langEastAsia: language }))
      .toEqual(['あい', '々う']);
  });
  it('honors paragraph-off semantics outside the observed edge exception', () => {
    expect(lines('あい々う', 'left', false)).toEqual(['あい', '々う']);
  });
});


it('retains paragraph one-mark refusal when the separate edge view permits the following closing head', () => {
  const ctx = { measureText: (text: string) => ({ width: [...text].length * 10,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
  const segment: LayoutTextSeg = { text: 'あい。』う', fontSize: 10, fontFamily: 'UnregisteredTestFace',
    bold: false, italic: false, underline: false, strikethrough: false,
    color: null, vertAlign: null, measuredWidth: 0 };
  // Both cuts are legal under this deliberately independent edge view. Only
  // the paragraph-owned one-mark refusal distinguishes hanging from not hanging.
  const permitted = { ...DEFAULT_KINSOKU_RULES, lineStartForbidden: new Set<number>() };
  const result = layoutLines(ctx, [segment], 20, 0, 1, [], undefined, {}, 0,
    DEFAULT_KINSOKU_RULES, undefined, 36, 20, false, false, false,
    undefined, 'bounded', undefined, true, false, undefined, 'one', undefined, permitted);
  expect(result[0]!.segments.flatMap(piece => 'text' in piece ? [piece.text] : []).join('')).toBe('あい');
});

it('keeps the OFF intrinsic atom and max-content widths with an active bounded edge exception', () => {
  const run: DocxTextRun = { text: 'あ。い', fontSize: 10, fontFamily: 'UnregisteredTestFace',
    bold: false, italic: false, underline: false, strikethrough: false,
    color: null, isLink: false, background: null, vertAlign: null, hyperlink: null };
  const paragraph: DocParagraph = { alignment: 'thaiDistribute', kinsoku: false,
    indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: null, numbering: null, tabStops: [], defaultFontSize: 10,
    runs: [{ type: 'text', ...run }] };
  const doc: DocxDocumentModel = { section, body: [{ type: 'paragraph', ...paragraph }],
    settings: { characterSpacingControl: 'compressPunctuation' },
    headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null } };
  const input = paragraphAcquisitionInput(paragraph, { story: 'body', storyInstance: 'body', path: [0] });
  const settings = resolveDocumentLayoutSettings(doc);
  const context = resolveParagraphLayoutContext(settings, resolveSectionLayoutContext(settings, section),
    { story: 'body', containers: [], lineNumberingEligible: true }, input);
  const measure = { font: '10px serif', letterSpacing: '0px',
    measureText: (text: string) => ({ width: [...text].length * 10,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
  const result = measureParagraphIntrinsicWidths(input, context, 100,
    { context: measure, fontFamilyClasses: {} }, {
      pageIndex: 0, totalPages: 1, documentHasEastAsianText: true,
      characterSpacingControl: 'compressPunctuation', pageWritingMode: 'horizontal-tb',
      layoutServices: createLayoutServices(doc, { measureContext: measure }),
    });
  // Existing intrinsic atoms consult paragraph tables even when enabled=false.
  // This preservation test does not extend normative OFF into that separate
  // sizing contract. The bounded edge view permits 。, so using it here would
  // incorrectly reduce this existing two-character atom to one character.
  expect(result).toEqual({ minWidthPt: 20, maxWidthPt: 30 });
});
