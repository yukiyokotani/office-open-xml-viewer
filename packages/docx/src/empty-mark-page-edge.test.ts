import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// WORD_TRAILING_EMPTY_MARK_BASELINE_ADMISSION, one case per measured Windows
// Word cell: an empty Times New Roman 11pt mark (12.65pt single line) at the
// page bottom, an ink paragraph after it, compatibility mode 14 with
// useFELayout. Off the grid the whole line must fit; on a 20pt lines grid the
// centered 12.65pt box (bottom 16.325pt into the cell) must fit. keepNext
// carries the mark with the paragraph that does not fit.

function para(text: string, extra: Partial<DocParagraph> = {}): BodyElement {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: text ? [{ type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 11, color: null, fontFamily: 'Times New Roman', isLink: false, background: null,
      vertAlign: null }] : [],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman', widowControl: false, ...extra,
  } as unknown as BodyElement;
}

function linearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * px() * 0.5,
      fontBoundingBoxAscent: px() * 0.9, fontBoundingBoxDescent: px() * 0.25,
      actualBoundingBoxAscent: px() * 0.7, actualBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function probePage(roomPt: number, keepNext: boolean, grid: boolean,
  mark: Partial<DocParagraph> = {}): number {
  const section = {
    pageWidth: 400, pageHeight: 72 + 100 + roomPt + 72,
    marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    ...(grid ? { docGridType: 'lines', docGridLinePitch: 20 } : {}),
  } as SectionProps;
  const model = {
    body: [para('Filler', { lineSpacing: { rule: 'exact', value: 100, explicit: true } }),
      para('', { keepNext, ...mark }), para('After'), para('Tail')],
    section,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [], endnotes: [], fontFamilyClasses: {},
    settings: { compatibilityMode: 14, useFeLayout: true },
  } as unknown as DocxDocumentModel;
  const pages = layoutDocument(model, createLayoutServices(model, {
    measureContext: linearCtx(),
    localMetrics: testFontSnapshot([
      { family: 'Times New Roman', lineHeightRatio: 1.15 },
      // Word's single lines in the controls: 21.6pt at 16pt, 14.3pt at 11pt.
      { family: 'Angsana New', lineHeightRatio: 1.35 },
      { family: 'MS Mincho', lineHeightRatio: 1.3 },
    ]),
  }), { currentDateMs: 0 }).pages;
  return pages.findIndex((page) => page.layers.body.some((node) => node.source.path[0] === 1));
}

describe('empty-mark page-edge admission (Windows Word cells)', () => {
  it.each([
    // excess4 006/010/014 (no grid) and 018/022/026 (no grid, keepNext)
    { roomPt: 13.15, keepNext: false, grid: false, stays: true },
    { roomPt: 12.15, keepNext: false, grid: false, stays: false },
    { roomPt: 9.65, keepNext: false, grid: false, stays: false },
    { roomPt: 13.15, keepNext: true, grid: false, stays: false },
    { roomPt: 12.15, keepNext: true, grid: false, stays: false },
    { roomPt: 9.65, keepNext: true, grid: false, stays: false },
    // excess4 030/034/038 and excess-height 025-040 (20pt grid)
    { roomPt: 20.5, keepNext: false, grid: true, stays: true },
    { roomPt: 17, keepNext: false, grid: true, stays: true },
    { roomPt: 16.5, keepNext: false, grid: true, stays: true },
    { roomPt: 16.25, keepNext: false, grid: true, stays: false },
    { roomPt: 16, keepNext: false, grid: true, stays: false },
    // excess4 042/046/050 (20pt grid, keepNext)
    { roomPt: 20.5, keepNext: true, grid: true, stays: false },
    { roomPt: 16.5, keepNext: true, grid: true, stays: false },
    { roomPt: 16, keepNext: true, grid: true, stays: false },
  ])('$roomPt pt left, keepNext $keepNext, grid $grid: stays $stays', ({ roomPt, keepNext, grid, stays }) => {
    expect(probePage(roomPt, keepNext, grid)).toBe(stays ? 0 : 1);
  });

  // excess7 (no grid): 0.5pt above and below the full line, 0.6pt below the
  // baseline. Only an automatic multiple's extra space may hang.
  const auto = (value: number) => ({ lineSpacing: { rule: 'auto' as const, value, explicit: true } });
  const exact = (value: number) => ({ lineSpacing: { rule: 'exact' as const, value, explicit: true } });
  const atLeast = (value: number) => ({ lineSpacing: { rule: 'atLeast' as const, value, explicit: true } });
  const font = (family: string, size: number) => ({ defaultFontFamily: family, defaultFontSize: size });
  it.each([
    ['auto 1.15', auto(1.15), 14.55, 9.8, [true, true, false]],
    ['auto 1.5', auto(1.5), 18.98, 9.8, [true, true, false]],
    ['auto 2.0', auto(2), 25.3, 9.8, [true, true, false]],
    ['exact 12pt', exact(12), 12, 9.6, [true, false, false]],
    ['exact 20pt', exact(20), 20, 16, [true, false, false]],
    ['atLeast 12pt', atLeast(12), 12.65, 9.8, [true, false, false]],
    ['atLeast 20pt', atLeast(20), 20, 17.6, [true, false, false]],
    ['Angsana New 16pt (Thai), single', font('Angsana New', 16), 21.6, 15.1, [true, false, false]],
    ['MS Mincho 11pt (CJK), single', font('MS Mincho', 11), 14.3, 11.1, [true, false, false]],
  ] as const)('%s', (_, mark, linePt, baselinePt, stays) => {
    const rooms = [linePt + 0.5, linePt - 0.5, baselinePt - 0.6];
    rooms.forEach((roomPt, index) => {
      expect(probePage(Math.round(roomPt * 100) / 100, false, false, mark as Partial<DocParagraph>), `room ${roomPt}`)
        .toBe(stays[index] ? 0 : 1);
    });
  });
});
