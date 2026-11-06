import { expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { layoutLines, type LayoutTextSeg } from './line-layout.js';
import { resolveParagraphLayoutContext, resolveSectionLayoutContext } from './layout-context.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { createLayoutServices } from './layout-runtime.js';
import { measureParagraphIntrinsicWidths } from './layout/intrinsic-width.js';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { BodyElement, DocParagraph, DocSettings, DocxDocumentModel, DocxTextRun, SectionProps } from './types.js';

const section: SectionProps = { pageWidth: 20, pageHeight: 400,
  marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false };
function lines(text: string, alignment = 'left', settings: DocSettings = {},
  overflowPunct = true, langEastAsia?: string, pieces?: readonly string[]): string[] {
  const run: DocxTextRun = { text, fontSize: 10, fontFamily: 'UnregisteredTestFace',
    bold: false, italic: false, underline: false, strikethrough: false,
    color: null, isLink: false, background: null, vertAlign: null, hyperlink: null,
    ...(langEastAsia === undefined ? {} : { langEastAsia }) };
  const paragraph: BodyElement = { type: 'paragraph', alignment, kinsoku: false, overflowPunct,
    indentLeft: 0, indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0,
    lineSpacing: null, numbering: null, tabStops: [],
    runs: (pieces ?? [text]).map((piece, index) => ({ type: 'text', ...run, text: piece, color: index ? 'FF0000' : null })) };
  const ctx = { font: '10px serif', letterSpacing: '0px', measureText: (value: string) => ({
    width: [...value].length * 10, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
  const result = layoutBodyModel([paragraph], section, ctx, {}, settings).pages.flatMap(page => page.layers.body.flatMap(node =>
    node.kind === 'paragraph' ? node.lines.map(line => line.placements.filter(p => p.kind === 'text').map(p => p.text).join('')) : []));
  expect(result.join('')).toBe(text);
  return result;
}

it.each(['left', 'both', 'distribute'])('neither hangs nor pushes the mode15 Japanese OFF closing mark under %s', alignment => {
  expect(lines('あい。う', alignment, { compatibilityMode: 15 })).toEqual(['あい', '。う']);
});
it('retains the pre15 OFF one-mark budget across adjacent closing marks', () => {
  expect(lines('あい。う', 'both', { compatibilityMode: 14 })).toEqual(['あい。', 'う']);
  expect(lines('あい。」う', 'both', { compatibilityMode: 14 })).toEqual(['あい。', '」う']);
  // The closing mark queued from another actual paint run cannot renew the
  // punctuation allowance already used by the first run.
  expect(lines('あい。」う', 'both', { compatibilityMode: 14 }, true, undefined, ['あい。', '」う']))
    .toEqual(['あい。', '」う']);
});
it.each(['lowKashida', 'thaiDistribute'])('does not push an OFF closing head in ordinary %s', alignment => {
  expect(lines('あい。う', alignment, { compatibilityMode: 15 })).toEqual(['あい', '。う']);
});
it('keeps the independent legacy Thai iteration, opening and overflow-off prohibitions', () => {
  const settings = { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' };
  expect(lines('あい。う', 'thaiDistribute', settings)).toEqual(['あい', '。う']);
  expect(lines('あい々う', 'thaiDistribute', settings)).toEqual(['あ', 'い々', 'う']);
  expect(lines('あ（いう', 'thaiDistribute', settings)).toEqual(['あ', '（い', 'う']);
  expect(lines('あい。う', 'thaiDistribute', settings, false)).toEqual(['あ', 'い。', 'う']);
});
it('preserves the previous out-of-scope language and present-empty custom-table geometry', () => {
  for (const lang of ['zh-CN', 'ko-KR']) expect(lines('あい々う', 'left', { compatibilityMode: 15 }, false, lang))
    .toEqual(['あ', 'い々', 'う']);
  expect(lines('あい。う', 'thaiDistribute', { compatibilityMode: 14, noLineBreaksAfter: '' }, false))
    .toEqual(['あ', 'い。', 'う']);
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
  const source = layoutSourceStore(doc);
  const input = source.blocks.body[0];
  if (input?.type !== 'paragraph') throw Error('Expected canonical paragraph');
  const settings = source.documentLayoutSettings;
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
