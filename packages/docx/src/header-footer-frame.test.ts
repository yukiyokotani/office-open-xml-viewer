import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { PaintNode } from './layout/types.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  FramePr,
  SectionProps,
} from './types';

// ECMA-376 §17.3.1.11 text frames in header/footer stories. A frame paragraph
// is positioned relative to its anchors and occupies no ordinary story flow;
// the following non-frame paragraph wraps around it. Word's classic footer page
// number is exactly this shape: `framePr wrap="around" vAnchor="text"
// hAnchor="margin" xAlign="right"` followed by an empty paragraph.
//
// Stub metrics: glyph advance = charCount × fontPx, font box = 0.8/0.2 em.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// Page 200×140, margins 20 ⇒ text band x 20..180. Footer bottom edge at 130.
const section: SectionProps = {
  pageWidth: 200, pageHeight: 140,
  marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
  headerDistance: 10, footerDistance: 10, titlePage: false, evenAndOddHeaders: false,
};

function textRun(text: string, fontSize = 10): BodyElement extends never ? never : DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocParagraph['runs'][number];
}

function paragraph(text: string, framePr?: FramePr): BodyElement {
  const value: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: text ? [textRun(text)] : [],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics',
    ...(framePr ? { framePr } : {}),
  };
  return { type: 'paragraph', ...value } as BodyElement;
}

function frame(over: Partial<FramePr> = {}): FramePr {
  return {
    dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'margin', vAnchor: 'text',
    hRule: 'auto', hSpace: 0, vSpace: 0, xAlign: 'right', y: 0.05, ...over,
  };
}

function layout(footer: BodyElement[], header: BodyElement[] = []) {
  const story = (body: BodyElement[]) => body.length ? { body } : null;
  const model: DocxDocumentModel = {
    body: [paragraph('body')],
    section,
    headers: { default: story(header), first: null, even: null },
    footers: { default: story(footer), first: null, even: null },
    footnotes: [],
    endnotes: [],
    fontFamilyClasses: {},
  };
  const services = createLayoutServices(model, { measureContext: makeCtx() });
  return layoutDocument(model, services, { currentDateMs: 0 }).pages[0]!.layers;
}

const textOf = (node: PaintNode): string =>
  node.kind === 'paragraph'
    ? node.lines.flatMap((line) => line.placements)
        .filter((placement) => placement.kind === 'text')
        .map((placement) => placement.text).join('')
    : '';

const paragraphs = (nodes: readonly PaintNode[]) =>
  nodes.filter((node): node is Extract<PaintNode, { kind: 'paragraph' }> =>
    node.kind === 'paragraph');

describe('header/footer text frames (§17.3.1.11)', () => {
  it('positions a text-anchored footer frame beside its anchor without advancing the story', () => {
    const framed = layout([paragraph('12', frame()), paragraph('')]);
    const footer = paragraphs(framed.footer);
    const number = footer.find((node) => textOf(node) === '12')!;
    expect(number.ordinaryFlow).toBe(false);
    // xAlign="right" against the margin band [20, 180]: a 20pt-wide frame.
    expect(number.flowBounds.xPt).toBeCloseTo(160);
    expect(number.flowBounds.widthPt).toBeCloseTo(20);

    // The in-flow fallback stacked the number above the empty anchor line; a
    // positioned frame adds no story advance, so the footer is one line tall
    // and its frame sits on the anchor line (y = 0.05pt below its top).
    const inFlow = layout([paragraph('12'), paragraph('')]);
    const inFlowFooter = paragraphs(inFlow.footer);
    const inFlowNumber = inFlowFooter.find((node) => textOf(node) === '12')!;
    const anchor = footer.find((node) => node !== number)!;
    expect(number.flowBounds.yPt).toBeCloseTo(anchor.flowBounds.yPt + 0.05);
    expect(anchor.flowBounds.yPt).toBeCloseTo(inFlowNumber.flowBounds.yPt + 10);
  });

  it('wraps the anchor paragraph text around a header frame', () => {
    const header = layout([], [paragraph('77', frame({ xAlign: 'left' })), paragraph('abc')]);
    const nodes = paragraphs(header.header);
    const anchor = nodes.find((node) => textOf(node) === 'abc')!;
    const text = anchor.lines[0]!.placements.find((placement) => placement.kind === 'text')!;
    // The 20pt frame at the left margin excludes x 20..40 on the anchor line.
    expect(text.bounds.xPt).toBeGreaterThanOrEqual(40 - 1e-6);
  });

  it('keeps page-anchored story frames in ordinary flow', () => {
    const footer = paragraphs(layout([paragraph('12', frame({ vAnchor: 'page', y: 5 })), paragraph('')]).footer);
    expect(footer.find((node) => textOf(node) === '12')!.ordinaryFlow).toBe(true);
  });
});
